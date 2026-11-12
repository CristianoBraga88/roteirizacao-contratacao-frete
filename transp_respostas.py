"""
Entende as respostas das transportadoras no WhatsApp (sem tela, sem navegador).

  eh_aceite(texto)                    -> True se ela disse que COLETA a carga
  cargas_abertas(contato, dias)       -> títulos enviados pra ela nos últimos dias
  casar_cargas(texto, titulos)        -> quais cargas a mensagem cita (cidade/peso/data)
  Aceites guardados em coletas_aceitas.json:
     status "aguardando" = ela vai coletar, falta você formar a carga e mandar o OK
     status "ok"         = você já marcou OK
"""
import json
import re
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path

BASE = Path(__file__).resolve().parent
LOG = BASE / "log_contratacao.csv"
ACEITES = BASE / "coletas_aceitas.json"
VISTOS = BASE / "transp_vistos.json"


def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return " ".join(re.sub(r"[^a-z0-9/.,: ]", " ", s.lower()).split())


# ---------------------------------------------------------------- aceite
SIM = ("coleto", "coletamos", "coletar", "coleta sim", "consigo", "conseguimos", "consegue sim",
       "pode programar", "pode mandar", "pode lancar", "pode fechar", "programado", "programei",
       "aceito", "aceitamos", "fechado", "fechamos", "fecho", "vamos carregar", "vou carregar",
       "carrego", "carregamos", "tenho carro", "tenho veiculo", "tenho caminhao", "tenho carreta",
       "tenho truck", "temos carro", "temos veiculo", "temos caminhao", "temos carreta", "temos truck",
       "segue placa", "placa", "motorista", "pode ser nosso", "fico com", "ficamos com", "pego",
       "pegamos", "vou pegar", "vamos pegar", "atendo", "atendemos", "vou atender", "vamos atender",
       "confirmado", "confirmo", "ok vamos", "manda pra mim", "e nosso", "eh nosso")
NAO = ("nao consigo", "nao conseguimos", "nao coleto", "nao coletamos", "nao temos", "nao tenho",
       "sem carro", "sem veiculo", "sem caminhao", "sem carreta", "sem truck", "nao atendo",
       "nao atendemos", "nao vou", "nao vamos", "nao da", "nao rola", "infelizmente", "indisponivel",
       "nao aceito", "nao fecha", "recuso", "declino", "nao pego", "nao pegamos")


def eh_aceite(texto):
    t = " " + norm(texto) + " "
    if any(f" {n}" in t for n in NAO):
        return False
    return any(f" {s}" in t for s in SIM)


# ---------------------------------------------------------------- cargas enviadas
def cargas_abertas(contato, dias=10):
    """Títulos enviados para o contato nos últimos dias (pelo log do robô)."""
    titulos = []
    limite = datetime.now() - timedelta(days=dias)
    try:
        linhas = LOG.read_text(encoding="utf-8-sig").splitlines()[1:]
    except OSError:
        return []
    for ln in linhas:
        p = ln.split(";", 3)
        if len(p) < 4 or p[1] != "ENVIADO" or norm(p[2]) != norm(contato):
            continue
        try:
            if datetime.strptime(p[0], "%d/%m/%Y %H:%M:%S") < limite:
                continue
        except ValueError:
            continue
        det = p[3].split(" | ", 1)
        if len(det) == 2:
            for t in det[1].split(" / "):
                t = t.strip()
                if t and t not in titulos:
                    titulos.append(t)
    return titulos


def _partes(titulo):
    """'SP - TAUBATE - 19.256 kg - CARRETA *AGENDADO 08/10/2026 10:00*'
       -> cidade 'taubate', peso '19.256', data '08/10'"""
    t = re.sub(r"^AL X - ", "", titulo.strip())
    pedacos = [p.strip() for p in t.split(" - ")]
    cidade = norm(pedacos[1]) if len(pedacos) > 1 else ""
    peso = re.search(r"(\d{1,3}(?:\.\d{3})+|\d+)\s*kg", t, re.I)
    data = re.search(r"(\d{2}/\d{2})(?:/\d{4})?", t)
    return cidade, (peso.group(1) if peso else ""), (data.group(1) if data else "")


def _dias_citados(t):
    """'dia 08', 'dia 8', '08/10' -> {'08'}"""
    dias = {f"{int(d):02d}" for d in re.findall(r"\bdia\s+(\d{1,2})\b", t)}
    dias |= {f"{int(d):02d}" for d, _ in re.findall(r"\b(\d{1,2})/(\d{1,2})\b", t)}
    return dias


def _cita_cidade(t, cidade):
    if not cidade:
        return False
    if f" {cidade} " in t:
        return True
    # abreviação: "itaqua" -> itaquaquecetuba, "ribeirao" sozinho não (ambíguo, fica p/ desempate)
    palavras = t.split()
    primeira = cidade.split()[0]
    return any(len(p) >= 5 and primeira.startswith(p) for p in palavras)


def casar_cargas(texto, titulos):
    """Quais títulos a mensagem cita. Pela cidade; se a cidade repete, desempata
    pela data ('dia 08', '08/10') ou pelo peso. Sem citação e só 1 carga aberta -> essa."""
    t = " " + norm(texto) + " "
    dias = _dias_citados(t)
    sem_ponto = t.replace(".", "")
    achados = []
    por_cidade = {}
    for exato in (True, False):  # 1º cidade escrita inteira; só se não achar, abreviação
        for tit in titulos:
            cid, peso, data = _partes(tit)
            if (cid and f" {cid} " in t) if exato else _cita_cidade(t, cid):
                por_cidade.setdefault(cid, []).append((tit, peso, data))
        if por_cidade:
            break
    for cid, lista in por_cidade.items():
        if len(lista) == 1:
            achados.append(lista[0][0])
            continue
        filtrados = [tit for tit, peso, data in lista
                     if (data and data[:2] in dias) or (peso and peso.replace(".", "") in sem_ponto)]
        achados += filtrados or [tit for tit, _, _ in lista]
    if not achados:
        for tit in titulos:  # citou só o peso ("a de 19.256")
            _, peso, _ = _partes(tit)
            if peso and len(peso) >= 4 and peso.replace(".", "") in sem_ponto:
                achados.append(tit)
    if not achados and len(titulos) == 1:
        achados = list(titulos)
    return achados


# ---------------------------------------------------------------- aceites guardados
def ler_aceites():
    try:
        return json.loads(ACEITES.read_text(encoding="utf-8"))
    except Exception:
        return []


def gravar_aceites(lista):
    try:
        ACEITES.write_text(json.dumps(lista[-400:], ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError:
        pass


def titulos_ja_aceitos():
    return {a["titulo"] for a in ler_aceites() if a.get("titulo")}


def ler_vistos():
    try:
        return set(json.loads(VISTOS.read_text(encoding="utf-8")))
    except Exception:
        return set()


def gravar_vistos(v):
    try:
        VISTOS.write_text(json.dumps(sorted(v)[-3000:], ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def hora_da_mensagem(pre):
    """'[14:32, 07/10/2026] Contato Alfa: ' -> datetime (ou None)."""
    m = re.search(r"\[(\d{1,2}):(\d{2}),\s*(\d{1,2})/(\d{1,2})/(\d{4})\]", pre or "")
    if not m:
        return None
    h, mi, d, mo, a = map(int, m.groups())
    try:
        return datetime(a, mo, d, h, mi)
    except ValueError:
        return None


def resumo(titulo):
    """'SP - TAUBATE - 19.256 kg - CARRETA *AGENDADO 08/10/2026 10:00*' -> 'TAUBATE 19.256 kg (08/10)'"""
    cid, peso, data = _partes(titulo)
    t = re.sub(r"^AL X - ", "", titulo.strip()).split(" - ")
    nome = t[1] if len(t) > 1 else titulo
    return f"{nome} {peso} kg" + (f" ({data})" if data else "")
