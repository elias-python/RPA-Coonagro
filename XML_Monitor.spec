# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


PROJECT_DIR = Path(globals().get('SPECPATH') or Path.cwd()).resolve()
ICON_CANDIDATES = (
    PROJECT_DIR / 'robo.ico',
    PROJECT_DIR / 'robo.png',
    PROJECT_DIR / 'assets' / 'robo.ico',
    PROJECT_DIR / 'assets' / 'robo.png',
)


def _resolver_fonte_icone():
    for caminho in ICON_CANDIDATES:
        if caminho.exists():
            return caminho
    return None


def _resolver_icone_exe(icon_source):
    if icon_source is None:
        return None
    if icon_source.suffix.lower() == '.ico':
        return str(icon_source)

    destino = PROJECT_DIR / 'build' / 'icon' / 'robo.ico'
    destino.parent.mkdir(parents=True, exist_ok=True)

    try:
        from PIL import Image

        with Image.open(icon_source) as imagem:
            imagem = imagem.convert('RGBA')
            lado = max(imagem.size)
            canvas = Image.new('RGBA', (lado, lado), (0, 0, 0, 0))
            posicao = ((lado - imagem.width) // 2, (lado - imagem.height) // 2)
            canvas.paste(imagem, posicao, imagem)
            canvas.save(
                destino,
                format='ICO',
                sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)],
            )
    except Exception as exc:
        print(f'[build] Falha ao gerar icone a partir de {icon_source}: {exc}')
        return None

    return str(destino)


ICON_SOURCE = _resolver_fonte_icone()
ICON_DATA = [(str(ICON_SOURCE), '.')] if ICON_SOURCE else []
ICON_EXE = _resolver_icone_exe(ICON_SOURCE)

if ICON_SOURCE:
    print(f'[build] Icone localizado: {ICON_SOURCE}')
    if ICON_EXE:
        print(f'[build] Icone do EXE: {ICON_EXE}')
    else:
        print('[build] Icone encontrado, mas o build seguira sem icone customizado no EXE.')
else:
    print('[build] Nenhum arquivo robo.png/robo.ico encontrado para o EXE.')


a = Analysis(
    ['XML Monitoring.py'],
    pathex=[],
    binaries=[],
    datas=ICON_DATA,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='XML_Monitor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=ICON_EXE,
)
