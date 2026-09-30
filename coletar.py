#!/usr/bin/env python3
"""
Coleta empresas de seringueira e eucalipto com telefone DDD 17 (SP)
na base aberta de CNPJ da Receita Federal. Gera um painel (HTML) e uma
planilha (Excel) e envia os dois para o Telegram.
"""
import csv
import datetime
import html
import json
import os
import re
import sys
import time
import zipfile
from collections import defaultdict
from pathlib import Path
from urllib.parse import quote

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# ================== CONFIGURAÇÃO (pode editar) ==================
CNAES = {
    "0139306": "Seringueira",   # Cultivo de seringueira
    "0210101": "Eucalipto",     # Cultivo de eucalipto
}
DDDS = {"17"}
UF = "SP"
APENAS_ATIVAS = True
# ================================================================

MODO_TESTE = os.getenv("MODO_TESTE", "false").lower() == "true"
TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()
CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "").strip()

RF_HOST = "https://arquivos.receitafederal.gov.br"
# Códigos da pasta compartilhada da Receita (o primeiro é o atual; os outros são reserva)
RF_SHARES = [x for x in os.getenv("RF_SHARE", "YggdBLfdninEJX9,gn672Ad4CF8N6TK").split(",") if x]
RF_LEGADO = os.getenv("RF_URL_BASE", "https://arquivos.receitafederal.gov.br/dados/cnpj/dados_abertos_cnpj/")

TRAB = Path("trabalho")
SAIDA = Path("saida")
HIST = Path("historico/cnpjs.txt")

S = requests.Session()
S.headers["User-Agent"] = "Mozilla/5.0 (coleta-produtores-ddd17)"

QUALIF = {"49": "Sócio-administrador", "22": "Sócio", "05": "Administrador",
          "16": "Presidente", "10": "Diretor", "65": "Titular", "08": "Conselheiro"}
NATUREZA = {"2135": "Empresário individual", "4120": "Produtor rural (PF)",
            "2062": "Sociedade limitada", "2305": "Eireli", "2143": "Cooperativa",
            "2046": "S.A. fechada", "2054": "S.A. aberta", "2240": "Sociedade simples limitada",
            "2232": "Sociedade simples pura", "2313": "Eireli (simples)"}
PORTE = {"00": "Não informado", "01": "Microempresa", "03": "Pequeno porte", "05": "Demais"}
NATUREZAS_PESSOA = {"2135", "4120", "2305", "2313"}


def log(msg):
    print(f"[{datetime.datetime.now():%H:%M:%S}] {msg}", flush=True)


# ---------------------------------------------------------------- Telegram
def telegram(metodo, **kw):
    if not TOKEN or not CHAT_ID:
        log("Telegram não configurado, pulando envio.")
        return
    try:
        r = requests.post(f"https://api.telegram.org/bot{TOKEN}/{metodo}", timeout=180, **kw)
        if not r.ok:
            log(f"Telegram {metodo} falhou: {r.text[:300]}")
    except Exception as e:
        log(f"Telegram {metodo} erro: {e}")


def enviar_msg(texto):
    telegram("sendMessage", data={"chat_id": CHAT_ID, "text": texto, "parse_mode": "HTML"})


def enviar_arquivo(caminho, legenda=""):
    with open(caminho, "rb") as f:
        telegram("sendDocument", data={"chat_id": CHAT_ID, "caption": legenda},
                 files={"document": (caminho.name, f)})


# ---------------------------------------------------------------- Receita
PROBLEMAS = []


def _variantes(share, caminho):
    """Formas conhecidas de acessar a pasta compartilhada da Receita (Nextcloud)."""
    return [
        ("webdav", f"{RF_HOST}/public.php/webdav{caminho}", (share, "")),
        ("dav", f"{RF_HOST}/public.php/dav/files/{share}{caminho}", None),
    ]


def _itens(xml, caminho):
    """Lê a resposta PROPFIND e devolve [(nome, é_pasta)] sem a própria pasta."""
    itens = []
    for bloco in re.findall(r"<(?:\w+:)?response>(.*?)</(?:\w+:)?response>", xml, re.S | re.I):
        h = re.search(r"<(?:\w+:)?href>([^<]+)</(?:\w+:)?href>", bloco, re.I)
        if not h:
            continue
        nome = requests.utils.unquote(h.group(1)).rstrip("/").split("/")[-1]
        pasta = bool(re.search(r"<(?:\w+:)?collection\s*/?>", bloco, re.I))
        if nome and nome != caminho.rstrip("/").split("/")[-1]:
            itens.append((nome, pasta))
    return itens


def listar(share, caminho, variante=None):
    for nome, url, auth in _variantes(share, caminho):
        if variante and nome != variante:
            continue
        try:
            r = S.request("PROPFIND", url, auth=auth, headers={"Depth": "1"}, timeout=90)
            if r.status_code in (200, 207):
                return nome, _itens(r.text, caminho)
            PROBLEMAS.append(f"{nome} {caminho or '/'}: HTTP {r.status_code}")
        except Exception as e:
            PROBLEMAS.append(f"{nome} {caminho or '/'}: {type(e).__name__}")
    return None, []


def achar_fonte():
    prioridade = ("dados", "cadastros", "cnpj", "dados_abertos_cnpj")
    for share in RF_SHARES:
        log(f"Procurando na pasta compartilhada {share}...")
        variante, raiz = listar(share, "")
        if not variante:
            continue
        fila, vistos = [("", raiz, 0)], set()
        while fila:
            caminho, itens, prof = fila.pop(0)
            meses = sorted((n for n, p in itens if p and re.fullmatch(r"\d{4}-\d{2}", n)), reverse=True)
            for mes in meses[:3]:
                pasta = f"{caminho}/{mes}"
                _, arquivos = listar(share, pasta, variante)
                if any(n.lower().startswith("estabelecimentos") for n, _ in arquivos):
                    log(f"Encontrei: {pasta}")
                    return {"tipo": variante, "share": share, "pasta": pasta, "mes": mes}
            if prof >= 4:
                continue
            subpastas = [n for n, p in itens if p and not re.fullmatch(r"\d{4}-\d{2}", n)]
            subpastas.sort(key=lambda n: (n.lower() not in prioridade, n.lower()))
            for n in subpastas[:8]:
                sub = f"{caminho}/{n}"
                if sub in vistos:
                    continue
                vistos.add(sub)
                _, filhos = listar(share, sub, variante)
                fila.append((sub, filhos, prof + 1))

    try:
        r = S.get(RF_LEGADO, timeout=90)
        r.raise_for_status()
        meses = sorted(set(re.findall(r'href="(\d{4}-\d{2})/"', r.text)))
        if meses:
            return {"tipo": "http", "pasta": RF_LEGADO.rstrip("/") + "/" + meses[-1], "mes": meses[-1]}
    except Exception as e:
        PROBLEMAS.append(f"endereço antigo: {type(e).__name__}")
    detalhes = "; ".join(PROBLEMAS[:6]) or "sem resposta"
    raise RuntimeError(f"Não encontrei os arquivos da Receita ({detalhes}).")


def url_arquivo(fonte, nome):
    if fonte["tipo"] == "http":
        return f"{fonte['pasta']}/{nome}", None
    for variante, url, auth in _variantes(fonte["share"], f"{fonte['pasta']}/{nome}"):
        if variante == fonte["tipo"]:
            return url, auth


def baixar(fonte, nome):
    destino = TRAB / nome
    url, auth = url_arquivo(fonte, nome)
    tamanho_antes = -1
    for tentativa in range(1, 11):
        try:
            feito = destino.stat().st_size if destino.exists() else 0
            cab = {"Range": f"bytes={feito}-"} if feito else {}
            with S.get(url, auth=auth, headers=cab, stream=True, timeout=180) as r:
                if r.status_code == 416:
                    pass
                else:
                    r.raise_for_status()
                    modo = "ab" if (feito and r.status_code == 206) else "wb"
                    with open(destino, modo) as f:
                        for bloco in r.iter_content(1 << 20):
                            f.write(bloco)
            with zipfile.ZipFile(destino) as z:
                z.namelist()
            log(f"  baixado {nome} ({destino.stat().st_size / 1e6:.0f} MB)")
            return destino
        except zipfile.BadZipFile:
            atual = destino.stat().st_size if destino.exists() else 0
            if atual == tamanho_antes:
                destino.unlink(missing_ok=True)
            tamanho_antes = atual
            log(f"  {nome} incompleto, tentativa {tentativa}")
        except Exception as e:
            log(f"  {nome} tentativa {tentativa}: {e}")
        time.sleep(min(90, 10 * tentativa))
    raise RuntimeError(f"Não consegui baixar {nome} depois de 10 tentativas.")


def linhas(caminho, basicos=None, pre=None):
    with zipfile.ZipFile(caminho) as z:
        for nome in z.namelist():
            with z.open(nome) as f:
                for bruta in f:
                    if basicos is not None and bruta[1:9] not in basicos:
                        continue
                    if pre is not None and not pre(bruta):
                        continue
                    try:
                        campos = next(csv.reader([bruta.decode("latin-1")], delimiter=";"))
                    except Exception:
                        continue
                    yield [c.strip() for c in campos]


# ---------------------------------------------------------------- Formatação
def fmt_tel(ddd, num):
    ddd = ddd.lstrip("0")
    num = re.sub(r"\D", "", num)
    if not ddd or len(num) < 8:
        return None
    cel = len(num) == 9 and num[0] == "9"
    if len(num) == 8 and num[0] in "6789":
        num, cel = "9" + num, True
    return {"f": f"({ddd}) {num[:-4]}-{num[-4:]}", "d": f"{ddd}{num}", "cel": cel}


def fmt_cnpj(c):
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"


def fmt_data(d):
    return f"{d[6:8]}/{d[4:6]}/{d[:4]}" if len(d) == 8 else ""


def fmt_dinheiro(v):
    try:
        n = float(v.replace(",", "."))
    except ValueError:
        return ""
    return "R$ " + f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def nome_limpo(razao):
    return re.sub(r"\s+\d{11}$", "", razao).strip()


# ---------------------------------------------------------------- Coleta
def coletar():
    TRAB.mkdir(exist_ok=True)
    fonte = achar_fonte()
    log(f"Base da Receita: {fonte['mes']} ({fonte['tipo']})")
    faixa = range(1) if MODO_TESTE else range(10)

    municipios = {}
    z = baixar(fonte, "Municipios.zip")
    for c in linhas(z):
        if len(c) >= 2:
            municipios[c[0]] = c[1].title()
    z.unlink()

    cnaes_b = [c.encode() for c in CNAES]
    uf_b = f'"{UF}"'.encode()

    def pre(b):
        return uf_b in b and any(c in b for c in cnaes_b)

    estab = {}
    for i in faixa:
        z = baixar(fonte, f"Estabelecimentos{i}.zip")
        for c in linhas(z, pre=pre):
            if len(c) < 28 or c[19] != UF:
                continue
            if APENAS_ATIVAS and c[5].lstrip("0") != "2":
                continue
            if not ({c[21].lstrip("0"), c[23].lstrip("0")} & DDDS):
                continue
            sec = [s for s in c[12].split(",") if s]
            culturas = [n for cod, n in CNAES.items() if cod == c[11] or cod in sec]
            if not culturas:
                continue
            estab[c[0] + c[1] + c[2]] = c + [culturas]
        z.unlink()
        log(f"Estabelecimentos{i}: {len(estab)} encontrados até agora")

    basicos = {k[:8].encode() for k in estab}
    empresas, socios = {}, defaultdict(list)
    if basicos:
        for i in faixa:
            z = baixar(fonte, f"Empresas{i}.zip")
            for c in linhas(z, basicos=basicos):
                empresas[c[0]] = c
            z.unlink()
        for i in faixa:
            z = baixar(fonte, f"Socios{i}.zip")
            for c in linhas(z, basicos=basicos):
                if len(c) > 4 and c[2]:
                    socios[c[0]].append({"n": c[2].title(), "q": QUALIF.get(c[4], "")})
            z.unlink()
    return fonte["mes"], montar(estab, empresas, socios, municipios)


def montar(estab, empresas, socios, municipios):
    leads = []
    for cnpj, c in estab.items():
        b = cnpj[:8]
        emp = empresas.get(b, [])
        razao = emp[1] if len(emp) > 1 else ""
        natureza = emp[2] if len(emp) > 2 else ""
        lista_socios = socios.get(b, [])
        if lista_socios:
            dono = ", ".join(s["n"] for s in lista_socios[:3])
        elif natureza in NATUREZAS_PESSOA or not natureza:
            dono = nome_limpo(razao).title()
        else:
            dono = ""
        cidade = municipios.get(c[20], c[20])
        rua = " ".join(x for x in (c[13], c[14]) if x).title()
        endereco = ", ".join(x for x in (rua, c[15], c[16].title(), c[17].title()) if x)
        tels = [t for t in (fmt_tel(c[21], c[22]), fmt_tel(c[23], c[24])) if t]
        leads.append({
            "cnpj": fmt_cnpj(cnpj), "id": cnpj,
            "dono": dono, "razao": nome_limpo(razao).title(), "fantasia": c[4].title(),
            "socios": lista_socios, "culturas": c[-1],
            "cidade": cidade, "endereco": endereco, "cep": c[18],
            "tels": tels, "email": c[27].lower(),
            "porte": PORTE.get(emp[5].zfill(2), "") if len(emp) > 5 else "",
            "natureza": NATUREZA.get(natureza, natureza),
            "capital": fmt_dinheiro(emp[4]) if len(emp) > 4 else "",
            "inicio": fmt_data(c[10]),
            "filial": c[3] == "2",
        })
    leads.sort(key=lambda d: (d["cidade"], d["dono"] or d["razao"]))
    return leads


# ---------------------------------------------------------------- Saídas
def marcar_novos(leads):
    anteriores = set(HIST.read_text().split()) if HIST.exists() else set()
    primeira = not anteriores
    for d in leads:
        d["novo"] = (not primeira) and d["id"] not in anteriores
    if not MODO_TESTE:
        HIST.parent.mkdir(exist_ok=True)
        HIST.write_text("\n".join(sorted(d["id"] for d in leads)) + "\n")
    return primeira


def gerar_excel(leads, caminho):
    wb = Workbook()
    ws = wb.active
    ws.title = "Produtores DDD 17"
    cab = ["Dono / sócios", "Razão social", "Nome fantasia", "Cultura", "Cidade",
           "Telefone 1", "Telefone 2", "WhatsApp", "E-mail", "Endereço", "CEP", "CNPJ",
           "Porte", "Natureza jurídica", "Capital social", "Início da atividade",
           "Novo nesta coleta", "Área (ha) - preencher", "Observações"]
    ws.append(cab)
    for d in leads:
        cel = next((t for t in d["tels"] if t["cel"]), None)
        ws.append([
            d["dono"], d["razao"], d["fantasia"], " + ".join(d["culturas"]), d["cidade"],
            d["tels"][0]["f"] if d["tels"] else "", d["tels"][1]["f"] if len(d["tels"]) > 1 else "",
            f"https://wa.me/55{cel['d']}" if cel else "", d["email"], d["endereco"], d["cep"],
            d["cnpj"], d["porte"], d["natureza"], d["capital"], d["inicio"],
            "Sim" if d.get("novo") else "", "", "",
        ])
    verde = PatternFill("solid", fgColor="2F5D46")
    for cel in ws[1]:
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = verde
    larguras = [32, 34, 24, 22, 20, 16, 16, 30, 28, 44, 11, 20, 16, 22, 16, 14, 10, 14, 30]
    for i, w in enumerate(larguras):
        ws.column_dimensions[chr(65 + i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(caminho)


def gerar_painel(leads, mes, caminho):
    info = {"mes": mes, "gerado": datetime.datetime.now().strftime("%d/%m/%Y"), "teste": MODO_TESTE}
    dados = json.dumps(leads, ensure_ascii=False).replace("</", "<\\/")
    pagina = PAINEL.replace("__DADOS__", dados).replace("__INFO__", json.dumps(info, ensure_ascii=False))
    caminho.write_text(pagina, encoding="utf-8")


def main():
    SAIDA.mkdir(exist_ok=True)
    if "--demo" in sys.argv:
        mes, leads = "demo", demo()
    else:
        mes, leads = coletar()
    primeira = marcar_novos(leads)

    painel = SAIDA / "painel.html"
    planilha = SAIDA / f"produtores_ddd17_{mes}.xlsx"
    gerar_painel(leads, mes, painel)
    gerar_excel(leads, planilha)

    ser = sum("Seringueira" in d["culturas"] for d in leads)
    euc = sum("Eucalipto" in d["culturas"] for d in leads)
    novos = sum(d["novo"] for d in leads)
    linha_novos = "Primeira coleta, a partir da próxima aviso quem é novo." if primeira else f"{novos} novas desde a última coleta."
    texto = (
        ("🧪 <b>Modo teste</b> (só parte da base)\n" if MODO_TESTE else "")
        + f"🌳 <b>Coleta DDD 17 concluída</b>\n"
        f"Base da Receita: {html.escape(mes)}\n"
        f"Seringueira: <b>{ser}</b>   Eucalipto: <b>{euc}</b>\n"
        f"Total: <b>{len(leads)}</b> empresas. {linha_novos}\n\n"
        "Abra o <b>painel.html</b> abaixo para filtrar e chamar no WhatsApp."
    )
    log(texto.replace("<b>", "").replace("</b>", ""))
    enviar_msg(texto)
    enviar_arquivo(painel, "Painel: toque para abrir no navegador")
    enviar_arquivo(planilha, "Planilha completa (coluna de hectares para preencher)")


def demo():
    import random
    random.seed(4)
    nomes = ["Antonio Carlos Ferreira", "Maria Aparecida Souza", "Jose Roberto Lima", "Helena Martins",
             "Luiz Fernando Baptista", "Sebastiao Rocha", "Claudia Nogueira", "Paulo Sergio Tavares"]
    cidades = ["Sao Jose Do Rio Preto", "Nhandeara", "Votuporanga", "Monte Aprazivel", "Poloni", "Jose Bonifacio"]
    out = []
    for i in range(40):
        n = random.choice(nomes)
        cult = random.choice([["Seringueira"], ["Seringueira"], ["Eucalipto"], ["Seringueira", "Eucalipto"]])
        cid = f"{random.randint(10, 99)}{random.randint(100000, 999999)}0001{random.randint(10, 99)}"
        out.append({"cnpj": fmt_cnpj(cid), "id": cid, "dono": n, "razao": f"Seringal {n.split()[-1]} Ltda",
                    "fantasia": "", "socios": [{"n": n, "q": "Sócio-administrador"}], "culturas": cult,
                    "cidade": random.choice(cidades), "endereco": "Estrada Vicinal, Km 12, Zona Rural",
                    "cep": "15000000", "tels": [fmt_tel("17", f"99{random.randint(1000000, 9999999)}")],
                    "email": "", "porte": "Microempresa", "natureza": "Sociedade limitada",
                    "capital": "R$ 50.000,00", "inicio": "12/03/2015", "filial": False})
    return out


PAINEL = r"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Produtores DDD 17</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,500;12..96,700&family=Source+Sans+3:wght@400;600&display=swap" rel="stylesheet">
<style>
:root{
  --fundo:#EDF1EE; --sup:#FFFFFF; --tinta:#17221D; --suave:#5A6A62; --linha:#D3DCD6;
  --ser:#2F5D46; --euc:#5B8E94; --amb:#9A5B3F; --novo:#B5462E;
  --titulo:"Bricolage Grotesque", "Avenir Next", "Segoe UI", system-ui, sans-serif;
  --texto:"Source Sans 3", "Segoe UI", system-ui, -apple-system, sans-serif;
  box-sizing:border-box;
  padding-top:env(safe-area-inset-top,0px); padding-bottom:env(safe-area-inset-bottom,0px);
}
@media (prefers-color-scheme: dark){:root{
  --fundo:#0F1512; --sup:#17201B; --tinta:#E2EAE4; --suave:#93A39A; --linha:#29352F;
  --ser:#72B38F; --euc:#8FC3C8; --amb:#D58C69; --novo:#EB7F63;
}}
*,*::before,*::after{box-sizing:inherit}
html{scroll-padding-top:env(safe-area-inset-top,0px)}
body{margin:0;background:var(--fundo);color:var(--tinta);font:16px/1.5 var(--texto)}
.wrap{max-width:720px;margin:0 auto;padding:28px 18px 60px}
h1{font:700 clamp(30px,8vw,46px)/1.02 var(--titulo);letter-spacing:-.02em;margin:0 0 10px}
.meta{color:var(--suave);margin:0 0 22px;font-size:15px}
.teste{display:inline-block;background:var(--novo);color:#fff;border-radius:4px;padding:1px 8px;font-size:13px;margin-bottom:12px}
.barra{display:flex;height:34px;border-radius:17px;overflow:hidden;gap:3px;background:var(--linha)}
.barra button{border:0;padding:0;cursor:pointer;min-width:6px;transition:filter .15s}
.barra button:hover,.barra button:focus-visible{filter:brightness(1.12)}
.barra button[aria-pressed=true]{outline:3px solid var(--tinta);outline-offset:-3px}
.legenda{display:flex;flex-wrap:wrap;gap:6px 18px;list-style:none;padding:0;margin:10px 0 0;font-size:15px}
.legenda b{font:700 20px var(--titulo);margin-right:4px}
.pt{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:6px;vertical-align:1px}
.controles{position:sticky;top:env(safe-area-inset-top,0px);z-index:2;background:var(--fundo);padding:16px 0 10px;margin-top:18px;border-bottom:1px solid var(--linha)}
input[type=search],select{width:100%;font:inherit;color:inherit;background:var(--sup);border:1px solid var(--linha);border-radius:10px;padding:11px 14px}
.linha2{display:flex;gap:10px;margin-top:10px;flex-wrap:wrap;align-items:center}
.linha2 select{flex:1 1 180px;width:auto}
.check{display:flex;align-items:center;gap:6px;font-size:15px;white-space:nowrap}
.check input{width:18px;height:18px;accent-color:var(--ser)}
.contagem{color:var(--suave);font-size:14px;margin:12px 0 4px}
.lead{background:var(--sup);border-radius:12px;padding:14px 16px 12px 18px;margin-top:10px;border-left:5px solid var(--ser)}
.lead.euc{border-left-color:var(--euc)} .lead.amb{border-left-color:var(--amb)}
.lead h2{font:600 19px/1.25 var(--titulo);margin:0}
.novo{color:var(--novo);font-size:13px;font-weight:600;margin-left:6px;vertical-align:2px}
.emp{color:var(--suave);margin:2px 0 0;font-size:15px}
.acoes{display:flex;flex-wrap:wrap;gap:8px;margin-top:10px}
.acoes a{text-decoration:none;font-size:15px;font-weight:600;padding:7px 12px;border-radius:8px;border:1px solid var(--linha);color:var(--tinta)}
.acoes a.wa{background:var(--ser);border-color:var(--ser);color:#fff}
@media (prefers-color-scheme: dark){.acoes a.wa{color:#0F1512}}
details{margin-top:8px;font-size:15px}
summary{cursor:pointer;color:var(--suave)}
dl{display:grid;grid-template-columns:auto 1fr;gap:3px 12px;margin:8px 0 2px}
dt{color:var(--suave)} dd{margin:0;overflow-wrap:anywhere}
.mais,.exportar{display:block;width:100%;margin-top:14px;font:600 16px var(--texto);padding:12px;border-radius:10px;cursor:pointer}
.mais{background:var(--sup);border:1px solid var(--linha);color:var(--tinta)}
.exportar{background:transparent;border:1px dashed var(--suave);color:var(--tinta)}
.vazio{padding:30px 0;color:var(--suave)}
footer{margin-top:28px;color:var(--suave);font-size:13px}
:focus-visible{outline:3px solid var(--euc);outline-offset:2px}
@media (prefers-reduced-motion: reduce){*{transition:none!important}}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div id="aviso"></div>
    <h1>Seringueira e eucalipto no DDD 17</h1>
    <p class="meta" id="meta"></p>
    <div class="barra" id="barra" role="group" aria-label="Filtrar por cultura"></div>
    <ul class="legenda" id="legenda"></ul>
  </header>

  <section class="controles">
    <input type="search" id="busca" placeholder="Buscar nome, empresa, cidade ou CNPJ" aria-label="Buscar">
    <div class="linha2">
      <select id="cidade" aria-label="Cidade"><option value="">Todas as cidades</option></select>
      <label class="check"><input type="checkbox" id="cel"> Com celular</label>
      <label class="check"><input type="checkbox" id="novos"> Só novos</label>
    </div>
  </section>

  <p class="contagem" id="contagem"></p>
  <main id="lista"></main>
  <button class="mais" id="mais" hidden>Mostrar mais</button>
  <button class="exportar" id="exportar">Baixar esta lista em CSV</button>

  <footer>Dados públicos do CNPJ (Receita Federal). Só aparecem produtores com CNPJ; quem planta como pessoa física sem CNPJ não está aqui. Ao entrar em contato, apresente-se e respeite quem pedir para não ser procurado (LGPD).</footer>
</div>

<script>
const DADOS = __DADOS__;
const INFO = __INFO__;
const COR = {ser:"var(--ser)", euc:"var(--euc)", amb:"var(--amb)"};
const ROTULO = {ser:"Só seringueira", euc:"Só eucalipto", amb:"As duas"};
const f = {texto:"", grupo:"", cidade:"", cel:false, novos:false};
let limite = 60;

const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const norm = s => (s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
const grupo = d => d.culturas.length > 1 ? "amb" : (d.culturas[0] === "Seringueira" ? "ser" : "euc");

DADOS.forEach(d => {
  d._g = grupo(d);
  d._b = norm([d.dono, d.razao, d.fantasia, d.cidade, d.cnpj, d.id, (d.socios || []).map(s => s.n).join(" ")].join(" "));
});

if (INFO.teste) $("aviso").innerHTML = '<span class="teste">Modo teste: só parte da base</span>';
$("meta").textContent = `${DADOS.length} empresas com CNPJ ativo. Base da Receita de ${INFO.mes}, gerado em ${INFO.gerado}.`;

const cont = {ser:0, euc:0, amb:0};
DADOS.forEach(d => cont[d._g]++);
["ser","euc","amb"].forEach(g => {
  if (!cont[g]) return;
  const b = document.createElement("button");
  b.style.flex = cont[g]; b.style.background = COR[g];
  b.setAttribute("aria-pressed", "false");
  b.setAttribute("aria-label", `${ROTULO[g]}: ${cont[g]}`);
  b.onclick = () => { f.grupo = f.grupo === g ? "" : g; atualizarBarra(); render(true); };
  b.dataset.g = g; $("barra").appendChild(b);
  $("legenda").insertAdjacentHTML("beforeend", `<li><span class="pt" style="background:${COR[g]}"></span><b>${cont[g]}</b>${ROTULO[g].toLowerCase()}</li>`);
});
function atualizarBarra(){ [...$("barra").children].forEach(b => b.setAttribute("aria-pressed", String(b.dataset.g === f.grupo))); }

[...new Set(DADOS.map(d => d.cidade))].sort((a,b) => a.localeCompare(b, "pt")).forEach(c => {
  const n = DADOS.filter(d => d.cidade === c).length;
  $("cidade").insertAdjacentHTML("beforeend", `<option value="${esc(c)}">${esc(c)} (${n})</option>`);
});

function filtrados(){
  const t = norm(f.texto).trim();
  return DADOS.filter(d =>
    (!f.grupo || d._g === f.grupo) &&
    (!f.cidade || d.cidade === f.cidade) &&
    (!f.cel || d.tels.some(x => x.cel)) &&
    (!f.novos || d.novo) &&
    (!t || d._b.includes(t)));
}

function card(d){
  const nome = d.dono || d.razao || d.fantasia || d.cnpj;
  const emp = [d.razao && d.razao !== nome ? d.razao : "", d.cidade].filter(Boolean).join(", em ");
  const cel = d.tels.find(x => x.cel);
  const mapa = "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(`${d.endereco}, ${d.cidade}, SP`);
  const acoes = [
    ...d.tels.map(x => `<a href="tel:+55${x.d}">Ligar ${esc(x.f)}</a>`),
    cel ? `<a class="wa" href="https://wa.me/55${cel.d}" target="_blank" rel="noopener">WhatsApp</a>` : "",
    d.endereco ? `<a href="${mapa}" target="_blank" rel="noopener">Mapa</a>` : ""
  ].join("");
  const socios = (d.socios || []).map(s => esc(s.n) + (s.q ? ` <span style="color:var(--suave)">(${esc(s.q)})</span>` : "")).join("<br>");
  const linhas = [
    ["Cultura", d.culturas.join(" e ")], ["Sócios", socios, true], ["Fantasia", d.fantasia],
    ["Endereço", [d.endereco, d.cep].filter(Boolean).join(", CEP ")], ["E-mail", d.email],
    ["CNPJ", d.cnpj + (d.filial ? " (filial)" : "")], ["Porte", d.porte], ["Natureza", d.natureza],
    ["Capital", d.capital], ["Desde", d.inicio]
  ].filter(l => l[1]).map(l => `<dt>${l[0]}</dt><dd>${l[2] ? l[1] : esc(l[1])}</dd>`).join("");
  return `<article class="lead ${d._g}">
    <h2>${esc(nome)}${d.novo ? '<span class="novo">novo</span>' : ""}</h2>
    <p class="emp">${esc(emp)}</p>
    <div class="acoes">${acoes || '<span style="color:var(--suave);font-size:15px">Sem telefone cadastrado</span>'}</div>
    <details><summary>Detalhes</summary><dl>${linhas}</dl></details>
  </article>`;
}

function render(reset){
  if (reset) limite = 60;
  const lista = filtrados();
  $("contagem").textContent = lista.length === DADOS.length ? `${lista.length} empresas` : `${lista.length} de ${DADOS.length} empresas`;
  $("lista").innerHTML = lista.length
    ? lista.slice(0, limite).map(card).join("")
    : '<p class="vazio">Nenhuma empresa com esses filtros. Limpe a busca ou escolha outra cidade.</p>';
  const resto = lista.length - limite;
  $("mais").hidden = resto <= 0;
  $("mais").textContent = `Mostrar mais ${Math.min(60, resto)}`;
}

$("busca").oninput = e => { f.texto = e.target.value; render(true); };
$("cidade").onchange = e => { f.cidade = e.target.value; render(true); };
$("cel").onchange = e => { f.cel = e.target.checked; render(true); };
$("novos").onchange = e => { f.novos = e.target.checked; render(true); };
$("mais").onclick = () => { limite += 60; render(false); };
$("exportar").onclick = () => {
  const col = ["Dono / sócios","Razão social","Cultura","Cidade","Telefone 1","Telefone 2","Celular","E-mail","Endereço","CNPJ"];
  const q = v => `"${String(v ?? "").replace(/"/g,'""')}"`;
  const rows = filtrados().map(d => [d.dono, d.razao, d.culturas.join(" + "), d.cidade,
    d.tels[0]?.f, d.tels[1]?.f, d.tels.find(x => x.cel)?.f, d.email, d.endereco, d.cnpj].map(q).join(";"));
  const blob = new Blob(["\ufeff" + [col.map(q).join(";"), ...rows].join("\n")], {type:"text/csv;charset=utf-8"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = "produtores_ddd17.csv"; a.click();
};
render(true);
</script>
</body>
</html>
"""

if __name__ == "__main__":
    try:
        main()
    except Exception as erro:
        log(f"ERRO: {erro}")
        enviar_msg(f"⚠️ A coleta falhou: {html.escape(str(erro))[:500]}\nAbra a aba Actions no GitHub para ver os detalhes.")
        raise
