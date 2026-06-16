from __future__ import annotations

import hashlib
import importlib.util
import sqlite3
import sys
import time
from pathlib import Path

import pythoncom
import win32com.client

ROOT_DIR = Path(__file__).resolve().parent
MONITOR_PATH = ROOT_DIR / "XML Monitoring.py"
PROCESSED_KEYS_PATH = ROOT_DIR / "outlook_rebusca_processados.txt"


def _carregar_monitor_module():
    spec = importlib.util.spec_from_file_location("xml_monitoring_module", MONITOR_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError("Nao foi possivel carregar XML Monitoring.py")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _carregar_chaves_processadas() -> set[str]:
    if not PROCESSED_KEYS_PATH.exists():
        return set()

    return {
        linha.strip()
        for linha in PROCESSED_KEYS_PATH.read_text(encoding="utf-8").splitlines()
        if linha.strip()
    }


def _registrar_chave_processada(chave: str) -> None:
    with PROCESSED_KEYS_PATH.open("a", encoding="utf-8") as handle:
        handle.write(f"{chave}\n")


def _montar_destino_xml(pasta_xml: Path, nome_arquivo: str, chave_outlook: str) -> Path:
    nome = Path(nome_arquivo)
    base = nome.stem or "xml"
    suffix = nome.suffix.lower() or ".xml"
    digest = hashlib.sha1(chave_outlook.encode("utf-8")).hexdigest()[:12]
    return pasta_xml / f"{base}__olk_{digest}{suffix}"


def _inserir_itens_novos(conn, itens: list[tuple], stamp: str) -> None:
    for item in itens:
        conn.execute(
            """
            INSERT OR REPLACE INTO notas_itens
                (chave_acesso, nf, seq, cod_material, descricao,
                 ordem_producao, qtd, un, vlr_unit, vlr_total_item,
                 cfop, nf_vinculada_5124, nf_origem, serie_origem,
                 vlr_total_nf, emissao, data_importacao)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            item + (stamp,),
        )


def main() -> int:
    monitor = _carregar_monitor_module()
    monitor.inicializar_db()

    pythoncom.CoInitialize()
    processadas = _carregar_chaves_processadas()
    pasta_xml = Path(monitor.pasta_xml)
    pasta_xml.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(monitor.caminho_db)
    try:
        chaves_banco = {
            row[0]
            for row in conn.execute("SELECT DISTINCT chave_acesso FROM notas_itens")
            if row[0]
        }

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        root_folder = outlook.DefaultStore.GetRootFolder()
        pasta = monitor._buscar_pasta_outlook(root_folder, "XML Coonagro")
        if not pasta:
            print("[ERRO] Pasta 'XML Coonagro' nao encontrada no Outlook.")
            return 1

        total_emails = int(getattr(pasta.Items, "Count", 0) or 0)
        anexos_lidos = 0
        anexos_salvos = 0
        xmls_novos_banco = 0
        ja_processados = 0
        ja_no_banco = 0
        ignorados = 0

        stamp = time.strftime("%Y-%m-%d %H:%M:%S")
        itens = pasta.Items
        for msg_idx in range(total_emails, 0, -1):
            try:
                msg = itens.Item(msg_idx)
                if getattr(msg, "Class", None) != 43:
                    continue

                entry_id = str(getattr(msg, "EntryID", "") or f"mail-{msg_idx}")
                attachments = msg.Attachments
                total_anexos = int(getattr(attachments, "Count", 0) or 0)
                for anexo_idx in range(1, total_anexos + 1):
                    anexo = attachments.Item(anexo_idx)
                    nome_arquivo = str(getattr(anexo, "FileName", "") or "").strip()
                    if not nome_arquivo.lower().endswith(".xml"):
                        continue

                    anexos_lidos += 1
                    tamanho = str(getattr(anexo, "Size", "") or "")
                    chave_outlook = (
                        f"{entry_id}|{anexo_idx}|{nome_arquivo.lower()}|{tamanho}"
                    )
                    if chave_outlook in processadas:
                        ja_processados += 1
                        continue

                    destino = _montar_destino_xml(
                        pasta_xml, nome_arquivo, chave_outlook
                    )
                    if not destino.exists():
                        anexo.SaveAsFile(str(destino))
                        anexos_salvos += 1

                    dados = monitor.extrair_detalhes_xml(str(destino))
                    if not dados:
                        ignorados += 1
                        processadas.add(chave_outlook)
                        _registrar_chave_processada(chave_outlook)
                        continue

                    chave_xml = dados[0][0]
                    if chave_xml in chaves_banco:
                        ja_no_banco += 1
                        processadas.add(chave_outlook)
                        _registrar_chave_processada(chave_outlook)
                        continue

                    _inserir_itens_novos(conn, dados, stamp)
                    chaves_banco.add(chave_xml)
                    xmls_novos_banco += 1
                    processadas.add(chave_outlook)
                    _registrar_chave_processada(chave_outlook)
                    monitor._log(
                        f"[REBUSCA OUTLOOK] NF importada novamente do Outlook: {nome_arquivo}"
                    )
            except Exception as exc:
                monitor._log(
                    f"[REBUSCA OUTLOOK] Falha ao processar e-mail #{msg_idx}: {exc}"
                )

        conn.commit()

        if xmls_novos_banco > 0:
            pendentes = monitor._somar_xmls_pendentes_recarregar(xmls_novos_banco)
            monitor._log(
                f"[REBUSCA OUTLOOK] {xmls_novos_banco} XML(s) novo(s) reencontrado(s). Execute Recarregar."
            )
            print(
                f"[OK] {xmls_novos_banco} XML(s) novo(s) inserido(s) no banco. "
                f"Pendentes para Recarregar: {pendentes}."
            )
        else:
            print("[INFO] Nenhum XML novo foi encontrado na revarredura do Outlook.")

        print(
            "[RESUMO] "
            f"E-mails: {total_emails} | "
            f"Anexos XML lidos: {anexos_lidos} | "
            f"Salvos localmente: {anexos_salvos} | "
            f"Ja processados: {ja_processados} | "
            f"Ja no banco: {ja_no_banco} | "
            f"Ignorados: {ignorados}"
        )
        return 0
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
