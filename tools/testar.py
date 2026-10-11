#!/usr/bin/env python3
"""Executa as regras do jogo fora do Excel e confere que a fuga tem solucao.

Os modulos Texto, Mapa, Jogo e Visual nao dependem do Office e rodam sem
alteracao no modo VBA do LibreOffice Basic. O roteiro em tests/Roteiro.bas
joga a fuga inteira, testa a criatura e grava quadros; este script confere o
resultado e desenha os quadros em docs/.

    python tools/testar.py            # confere e gera as imagens
    python tools/testar.py --testar   # so confere (usado no CI)

Requisitos: LibreOffice (soffice); para as imagens, Pillow e uma fonte com
os simbolos (DejaVu Sans).
"""
import html
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
MODULOS = ["Texto", "Mapa", "Jogo", "Visual"]
XBA = ('<?xml version="1.0" encoding="UTF-8"?>\n'
       '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
       '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="{nome}" '
       'script:language="StarBasic">{codigo}</script:module>\n')
XLB = ('<?xml version="1.0" encoding="UTF-8"?>\n'
       '<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
       '<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" '
       'library:readonly="false" library:passwordprotected="false">\n{itens}</library:library>\n')


def fonte(caminho):
    linhas = caminho.read_text(encoding="ascii").splitlines()
    return "\n".join(l for l in linhas if not l.startswith("Attribute "))


def executar_vba(pasta):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        sys.exit("soffice nao encontrado: instale o LibreOffice")
    perfil = "-env:UserInstallation=" + (pasta / "perfil").as_uri()
    subprocess.run([soffice, "--headless", perfil, "--terminate_after_init"], timeout=180,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    lib = pasta / "perfil" / "user" / "basic" / "Standard"
    lib.mkdir(parents=True, exist_ok=True)
    codigos = {m: "Option VBASupport 1\n" + fonte(RAIZ / "src" / f"{m}.bas") for m in MODULOS}
    codigos["Roteiro"] = fonte(RAIZ / "tests" / "Roteiro.bas")
    for nome, codigo in codigos.items():
        (lib / f"{nome}.xba").write_text(XBA.format(nome=nome, codigo=html.escape(codigo)), encoding="utf-8")
    itens = "".join(f' <library:element library:name="{n}"/>\n' for n in codigos)
    (lib / "script.xlb").write_text(XLB.format(itens=itens), encoding="utf-8")

    saida = pasta / "jogo.txt"
    try:
        subprocess.run([soffice, "--headless", perfil, "macro:///Standard.Roteiro.Principal"], timeout=180,
                       env={**os.environ, "JOGO_SAIDA": str(saida)},
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except subprocess.TimeoutExpired:
        # Sem tela, um erro de compilacao do Basic abre um aviso invisivel e trava.
        subprocess.run(["pkill", "-x", "soffice.bin"])
        sys.exit("o LibreOffice travou: quase sempre e erro de compilacao em algum modulo .bas")
    if not saida.exists():
        sys.exit("o roteiro VBA nao gerou saida")
    return saida.read_text(encoding="utf-8", errors="replace").splitlines()


def desenhar(nome, casas, destino, lado=22):
    from PIL import Image, ImageDraw, ImageFont
    largura = max(c[0] for c in casas) + 1
    altura = max(c[1] for c in casas) + 1
    img = Image.new("RGB", (largura * lado, altura * lado), (6, 6, 9))
    d = ImageDraw.Draw(img)
    try:
        f = ImageFont.truetype("DejaVuSans-Bold.ttf", int(lado * 0.72))
    except OSError:
        f = ImageFont.load_default()
    rgb = lambda n: (n & 255, (n >> 8) & 255, (n >> 16) & 255)      # o VBA guarda como BGR
    for x, y, fundo, simbolo, cor in casas:
        d.rectangle([x * lado, y * lado, (x + 1) * lado - 1, (y + 1) * lado - 1], fill=rgb(fundo))
        if simbolo:
            d.text((x * lado + lado / 2, y * lado + lado / 2), chr(simbolo), font=f, fill=rgb(cor), anchor="mm")
    img.save(destino)


if __name__ == "__main__":
    with tempfile.TemporaryDirectory() as tmp:
        linhas = executar_vba(Path(tmp))
    notas, quadros = [], {}
    for linha in linhas:
        c = linha.split("\t")
        if c[0] == "#":
            notas.append(c[1])
        elif c[0] == "Q":
            atual = quadros.setdefault(c[1], [])
        elif c[0] == "C":
            atual.append(tuple(int(v) for v in c[1:6]))
    for n in notas:
        print(" ", n)
    falhas = [n for n in notas if n.startswith(("FALHA", "ERRO"))]
    if not notas or notas[-1] != "FIM" or falhas:
        sys.exit("teste falhou")
    print(f"ok: {sum(n.startswith('ok ') for n in notas)} conferencias, {len(quadros)} quadros")
    if "--testar" not in sys.argv:
        for nome, casas in quadros.items():
            desenhar(nome, casas, RAIZ / "docs" / f"{nome}.png")
        print("imagens em docs/:", ", ".join(quadros))
