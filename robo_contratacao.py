"""
Robô de Contratação — fica aberto o dia todo, vigiando a pasta.

Quando você salva uma Base Roteirizacao.xls NOVA nesta pasta:
  1) gera o relatório de roteirização (o mesmo HTML de sempre);
  2) separa os títulos de CONTRATAÇÃO;
     - NÃO envia "AGENDADO" sem data (o PCP ainda vai colocar a data);
     - NÃO envia veículo "Sem Tipo de Entrega" (precisa do Customer Service);
     - NÃO repete título já enviado nos últimos dias;
  3) mostra uma janela para você conferir e clicar em ENVIAR;
  4) você marca, carga por carga, para qual transportadora vai (Contato Alfa /
     Contato Beta). O robô memoriza o ESTADO -> transportadora e já traz
     marcado da próxima vez;
  5) manda UMA mensagem por contato, só com as cargas dele.

USO: dois cliques em ROBO_CONTRATACAO.bat
"""
import configparser
import glob
import json
import os
import queue
import re
import subprocess
import threading
import time
import tkinter as tk
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import ttk

import transp_respostas as TR

BASE = Path(__file__).resolve().parent
CONFIG = BASE / "contatos_contratacao.ini"
ENVIADOS = BASE / "enviados_contratacao.json"
LOG = BASE / "log_contratacao.csv"
INTERVALO = 5  # segundos entre cada olhada na pasta

FUNDO, CARTAO, TEXTO, FRACO = "#0e1b10", "#152416", "#e9f3e6", "#a7c2a3"
VERDE, VERMELHO, AMBAR = "#5fb894", "#ff8a7a", "#d9b34a"
FONTE = "Segoe UI"

RE_AGENDADO_SEM_DATA = re.compile(r"\*\s*AGENDADO\s*\*", re.IGNORECASE)
ESTADOS = BASE / "estados_contratacao.json"
CONTRATADAS = BASE / "contratadas_contratacao.json"  # carga -> transportadora contratada (sem JCLOG)
ROTEIROS = BASE / "roteiros_contratacao.json"  # título de contratação -> cliente (pro histórico)  # memória: qual estado vai para qual contato
RE_UF = re.compile(r"^(?:AL X - )?([A-Z]{2}) - ")


def uf_do_titulo(titulo):
    """'MG - CONTAGEM - ...' -> 'MG' ; 'AL X - PE - PAULISTA - ...' -> 'PE'."""
    m = RE_UF.match(titulo.strip())
    return m.group(1) if m else "??"


def ler_estados():
    try:
        return json.loads(ESTADOS.read_text(encoding="utf-8"))
    except Exception:
        return {}


def gravar_estados(mapa):
    """mapa: {UF: [nomes de contato]} — atualiza só os estados informados."""
    atual = ler_estados()
    atual.update(mapa)
    ESTADOS.write_text(json.dumps(atual, ensure_ascii=False, indent=1, sort_keys=True),
                       encoding="utf-8")


# ------------------------------------------------------------------ regras
def ler_config():
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG, encoding="utf-8")
    contatos = []
    for sec in cfg.sections():
        if sec.lower().startswith("contato"):
            c = cfg[sec]
            if c.get("ativo", "sim").strip().lower() in ("sim", "s", "1", "true"):
                contatos.append({"nome": c.get("nome", "").strip(),
                                 "empresa": c.get("empresa", "").strip(),
                                 "telefone": re.sub(r"\D", "", c.get("telefone", ""))})
    msg = cfg["mensagem"] if cfg.has_section("mensagem") else {}
    rob = cfg["robo"] if cfg.has_section("robo") else {}
    return {
        "contatos": [c for c in contatos if c["nome"]],
        "cabecalho": msg.get("cabecalho", "{saudacao}! Seguem as cargas para contratação ({data}):"),
        "rodape": msg.get("rodape", ""),
        "abrir_relatorio": str(rob.get("abrir_relatorio", "sim")).lower() in ("sim", "s", "1", "true"),
        "dias": int(rob.get("dias_para_nao_repetir", 7) or 7),
        "vigiar_respostas": str(rob.get("vigiar_respostas", "sim")).lower() in ("sim", "s", "1", "true"),
        "lembrar_min": int(rob.get("lembrar_ok_minutos", 30) or 30),
        "regioes": {k.upper(): [x.strip().upper() for x in v.split(",") if x.strip()]
                    for k, v in (cfg["regioes"].items() if cfg.has_section("regioes") else [])},
        "transportadoras": [t.strip() for t in (cfg["transportadoras"].get("lista", "") if cfg.has_section("transportadoras") else "").split(",") if t.strip()]
                           or ["ALFA CARGAS", "GAMA LOG", "DELTA", "EPSILON", "ZETA LOGISTICA", "ETA & CIA",
                               "BETA TRANSP", "THETA", "IOTA", "KAPPA"],
        "msg_ok": (cfg["mensagem"].get("resposta_ok", "") if cfg.has_section("mensagem") else "")
                  or "Ok! Carga formada: {cargas}. Obrigado!",
    }


def ler_enviados(dias):
    try:
        dados = json.loads(ENVIADOS.read_text(encoding="utf-8"))
    except Exception:
        dados = {}
    limite = datetime.now() - timedelta(days=dias)
    return {t: d for t, d in dados.items()
            if datetime.strptime(d, "%d/%m/%Y %H:%M") >= limite}


CHAVES = {}  # título que aparece na tela -> chave gravada (título + nºs únicos da base)
DADOS = {}   # título que aparece na tela -> dados do veículo (paletes, peso, valor, agenda, coleta...)


def _br(n, dec=2):
    try:
        return f"{float(n):,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return str(n)


def _coleta(v):
    c = str(v.get("saida_uberaba") or "")
    return "sem data (sem agenda ou sem km)" if (not c or c == "sem agenda" or c.startswith("?")) else c


def texto_dados(v, recuo=""):
    """Linhas com os dados de uma carga (as mesmas do botão Dados do relatório)."""
    if not v:
        return ""
    obs = str(v.get("observacao") or "").strip()
    agenda = v.get("agenda_completa") or (v.get("agenda") or "sem agenda")
    linhas = [f"Cliente: {v.get('nome', '')}",
              f"Entrega agendada: {agenda} · Coletar até: {_coleta(v)}",
              f"Paletes: {_br(v.get('pallets') or 0)} · Peso: {v.get('peso', '')} · Valor NF: {v.get('valor', '')}"]
    if obs:
        linhas.append(f"Obs./senha: {obs}")
    return "\n".join(recuo + l for l in linhas)


REGIOES = {
    "SUL": ["PR", "SC", "RS"],
    "SUDESTE": ["SP", "MG", "RJ", "ES"],
    "CENTRO-OESTE": ["MT", "MS", "GO", "DF"],
    "NORDESTE": ["MA", "PI", "CE", "RN", "PB", "PE", "AL", "SE", "BA"],
    "NORTE": ["AC", "AM", "AP", "PA", "RO", "RR", "TO"],
}
OPCOES_REGIAO = ["SUL", "NORDESTE", "NORTE", "SP-CAPITAL", "SP-INTERIOR", "MG", "GO", "DF", "MT", "MS", "RJ", "ES"]
GRANDE_SP = {  # cidades consideradas "SP-CAPITAL" (Grande São Paulo); o resto de SP é "SP-INTERIOR"
    "SAO PAULO", "OSASCO", "GUARULHOS", "BARUERI", "CARAPICUIBA", "COTIA", "ITAPEVI", "JANDIRA", "SANTANA DE PARNAIBA",
    "TABOAO DA SERRA", "EMBU", "EMBU DAS ARTES", "EMBU-GUACU", "ITAPECERICA DA SERRA", "DIADEMA", "SAO BERNARDO DO CAMPO",
    "SANTO ANDRE", "SAO CAETANO DO SUL", "MAUA", "RIBEIRAO PIRES", "RIO GRANDE DA SERRA", "MOGI DAS CRUZES", "SUZANO",
    "ITAQUAQUECETUBA", "FERRAZ DE VASCONCELOS", "POA", "ARUJA", "SANTA ISABEL", "FRANCISCO MORATO", "FRANCO DA ROCHA",
    "CAIEIRAS", "CAJAMAR", "MAIRIPORA", "VARGEM GRANDE PAULISTA", "PIRAPORA DO BOM JESUS", "JUQUITIBA",
    "SAO LOURENCO DA SERRA", "BIRITIBA-MIRIM", "SALESOPOLIS", "GUARAREMA"}


def uf_cidade(v, titulo=""):
    v = v or {}
    cid_uf = str(v.get("cidade_uf") or "")
    if " - " in cid_uf:
        uf, cidade = cid_uf.split(" - ", 1)
    else:
        partes = [p.strip() for p in (titulo or "").split(" - ")]
        uf, cidade = (partes[0], partes[1]) if len(partes) >= 2 else ("", cid_uf)
    return uf.strip().upper(), _sem_acento(cidade).strip()


def atende(regras, uf, cidade):
    """regras: ['SUL','MT','SP-INTERIOR',...]. True se a transportadora atende esse destino."""
    for r in regras:
        r = r.strip().upper()
        if r in REGIOES and uf in REGIOES[r]:
            return True
        if r == uf:
            return True
        if uf == "SP" and r == "SP-CAPITAL" and cidade in GRANDE_SP:
            return True
        if uf == "SP" and r == "SP-INTERIOR" and cidade not in GRANDE_SP:
            return True
    return False


def transportadoras_sugeridas(v, titulo, cfg):
    uf, cidade = uf_cidade(v, titulo)
    return [emp for emp, regras in cfg.get("regioes", {}).items() if atende(regras, uf, cidade)]


def _mesma_empresa(a, b):
    a, b = _sem_acento(a).split(), _sem_acento(b).split()
    return bool(a and b and (a[0].startswith(b[0][:4]) or b[0].startswith(a[0][:4])))


def gravar_regioes(mapa):
    """mapa: {EMPRESA: [regras]} -> reescreve a seção [regioes] do INI (mantém o resto)."""
    txt = CONFIG.read_text(encoding="utf-8") if CONFIG.exists() else ""
    novo = ["[regioes]",
            "; Para onde cada transportadora atende (separe por vírgula). Opções: SUL, NORDESTE, NORTE,",
            "; SUDESTE, CENTRO-OESTE, SP (todo), SP-CAPITAL (Grande SP), SP-INTERIOR, ou a sigla do estado (MG, GO...).",
            "; O robô usa isso para já indicar para qual transportadora mandar cada carga."]
    novo += [f"{emp} = {', '.join(r)}" for emp, r in mapa.items() if r]
    bloco = "\n".join(novo) + "\n"
    m = re.search(r"^\[regioes\][^\[]*", txt, re.M)
    txt = (txt[:m.start()] + bloco + "\n" + txt[m.end():]) if m else (txt.rstrip("\n") + "\n\n" + bloco)
    CONFIG.write_text(txt, encoding="utf-8")


def linha_contratacao(v, titulo=""):
    """Linha curta: MG-UBERABA 22 TON CARRETA *AGENDADO* 13/11/2026  (ou *ORDEM DE CHEGADA*)."""
    v = v or {}
    cid_uf = str(v.get("cidade_uf") or "")
    if " - " in cid_uf:
        uf, cidade = cid_uf.split(" - ", 1)
    else:
        partes = [p.strip() for p in (titulo or "").split(" - ")]
        uf, cidade = (partes[0], partes[1]) if len(partes) >= 2 else ("", cid_uf)
    if uf.strip().upper().startswith("AL X"):          # título "AL X - PE - FLORESTA - ..."
        partes = [p.strip() for p in (titulo or "").split(" - ")]
        uf, cidade = (partes[1], partes[2]) if len(partes) >= 3 else (uf[4:], cidade)
    destino = f"{uf.strip().upper()}-{cidade.strip().upper()}"
    if str(v.get("origem") or "").startswith("Macei") or (titulo or "").upper().startswith("AL X"):
        destino = "AL X " + destino      # sai de Maceió/AL
    try:
        p = v.get("peso", "")
        kg = float(str(p).replace(".", "").replace(",", ".")) if isinstance(p, str) else float(p)
    except (TypeError, ValueError):
        m = re.search(r"([\d.]+)\s*kg", titulo or "")
        kg = float(m.group(1).replace(".", "")) if m else 0
    if kg >= 1000:
        t = kg / 1000
        peso = f"{round(t):.0f} TON" if t >= 3 else f"{t:.1f} TON".replace(".", ",")
    else:
        peso = f"{kg:.0f} KG"
    veic = str(v.get("tipo_veiculo") or "").upper()
    if not veic:
        veic = next((x for x in ("RODOTREM", "CARRETA", "TRUCK", "TOCO", "VAN") if x in (titulo or "").upper()), "")
    if "CHEGADA" in (titulo or "").upper() or "CHEGADA" in str(v.get("tipo_entrega") or "").upper():
        tipo = "*ORDEM DE CHEGADA*"
    else:
        ag = str(v.get("agenda_completa") or v.get("agenda") or "")
        m = re.search(r"\d{2}/\d{2}/\d{4}", ag) or re.search(r"\d{2}/\d{2}/\d{4}", titulo or "")
        tipo = "*AGENDADO*" + (f" {m.group(0)}" if m else "")
    return " ".join(x for x in (destino, peso, veic, tipo) if x)


def dados_caminhao(v, titulo=""):
    """Texto curto para mandar à transportadora: cidade, estado, peso, veículo e entrega."""
    v = v or {}
    cid_uf = str(v.get("cidade_uf") or "")
    if " - " in cid_uf:
        uf, cidade = cid_uf.split(" - ", 1)
    else:
        m = RE_UF.match(titulo or "")
        uf, cidade = (m.group(1) if m else ""), cid_uf
    peso = v.get("peso", "")
    try:
        peso = f"{float(str(peso).replace('.', '').replace(',', '.')) if isinstance(peso, str) else float(peso):,.0f} kg".replace(",", ".")
    except (TypeError, ValueError):
        peso = f"{peso}"
    veic = str(v.get("tipo_veiculo") or "").upper() or ("CARRETA" if "CARRETA" in (titulo or "").upper() else
                                                          "TRUCK" if "TRUCK" in (titulo or "").upper() else "")
    entrega = v.get("agenda_completa") or v.get("agenda") or ""
    if not entrega or entrega == "sem agenda":
        entrega = "ORDEM DE CHEGADA" if "CHEGADA" in (titulo or "").upper() else "sem agenda"
    linhas = [f"Cidade: {cidade.strip().title()}",
              f"Estado: {uf.strip().upper()}",
              f"Peso: {peso}",
              f"Veículo: {veic}",
              f"Entrega: {entrega}"]
    col = _coleta(v)
    if v and not col.startswith("sem data"):
        linhas.append(f"Coletar até: {col}")
    if v.get("pallets"):
        linhas.append(f"Paletes: {_br(v.get('pallets'), 0)}")
    return "\n".join(linhas)


PALAVRAS_TRANSP = {  # palavra no nome do contato -> empresa sugerida
    "GAMA LOG": "GAMA LOG", "GAMA": "GAMA LOG", "BETA TRANSP": "BETA TRANSP", "BETA": "BETA TRANSP",
    "DELTA": "DELTA", "IOTA": "IOTA", "EPSILON": "EPSILON", "ZETA": "ZETA LOGISTICA",
    "ETA": "ETA & CIA", "THETA": "THETA", "KAPPA": "KAPPA", "LAMBDA": "LAMBDA",
    "MU": "MU AGRO", "ALFA": "ALFA CARGAS", "CARGAS": "", "CARGA": "", "TRANSP": "", "LOG": "",
    "FRETE": "", "EXPEDI": "", "FROTA": "", "TRANSPORTE": "", "CAMINH": ""}


def _sem_acento(t):
    import unicodedata
    return unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode().upper()


def parece_transportadora(nome):
    """Devolve (sim/não, empresa sugerida) pelo nome do contato/grupo."""
    n = _sem_acento(nome)
    for p, emp in PALAVRAS_TRANSP.items():
        fim = r"($|[^A-Z])" if p in ("EPSILON", "ZETA", "MU", "ALFA", "LOG") else ""
        if re.search(r"(^|[^A-Z])" + p + fim, n) or (p == "LOG" and re.search(r"(^|[^A-Z])LOGIST", n)):
            return True, emp
    return False, ""


def adicionar_contatos_ini(novos):
    """novos: [(nome_whats, empresa)] -> acrescenta blocos [contatoN] no INI e a empresa na lista."""
    txt = CONFIG.read_text(encoding="utf-8") if CONFIG.exists() else ""
    nums = [int(x) for x in re.findall(r"^\[contato(\d+)\]", txt, re.M)] or [0]
    n = max(nums)
    existentes = {_sem_acento(x).strip() for x in re.findall(r"^nome\s*=\s*(.+)$", txt, re.M)}
    blocos = []
    for nome, emp in novos:
        if not nome.strip() or _sem_acento(nome).strip() in existentes:
            continue
        n += 1
        blocos.append(f"\n[contato{n}]\nnome = {nome}\nempresa = {emp}\ntelefone =\nativo = sim\n")
    if blocos:
        txt = txt.rstrip("\n") + "\n" + "".join(blocos)
    m = re.search(r"^lista\s*=\s*(.*)$", txt, re.M)
    atuais = [x.strip() for x in (m.group(1).split(",") if m else []) if x.strip()]
    for _, emp in novos:
        if emp and emp.upper() not in [a.upper() for a in atuais]:
            atuais.append(emp.upper())
    linha = "lista = " + ", ".join(atuais)
    if m:
        txt = txt[:m.start()] + linha + txt[m.end():]
    else:
        txt += "\n[transportadoras]\n" + linha + "\n"
    CONFIG.write_text(txt, encoding="utf-8")
    return len(blocos)


def ler_contratadas():
    try:
        return json.loads(CONTRATADAS.read_text(encoding="utf-8"))
    except Exception:
        return {}


def gravar_contratada(chave, transportadora, titulo):
    d = ler_contratadas()
    if transportadora:
        d[chave] = {"transportadora": transportadora, "quando": datetime.now().strftime("%d/%m/%Y %H:%M"), "titulo": titulo}
    else:
        d.pop(chave, None)
    try:
        CONTRATADAS.write_text(json.dumps(dict(list(d.items())[-3000:]), ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError:
        pass


def chave_carga(titulo, nros):
    """A mesma carga = mesmo título E mesmos nºs únicos. Duas cargas iguais (mesma cidade, peso,
    veículo e agenda) mas com nºs únicos diferentes são cargas DIFERENTES."""
    nros = ",".join(sorted(n.strip() for n in str(nros or "").replace("—", "").split(",") if n.strip()))
    return f"{titulo} #{nros}" if nros else titulo


def gravar_enviados(titulos):
    dados = ler_enviados(365)
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    for t in titulos:
        dados[CHAVES.get(t, t)] = agora
    ENVIADOS.write_text(json.dumps(dados, ensure_ascii=False, indent=1), encoding="utf-8")


def separar_titulos(veiculos, enviados):
    """Retorna (para_enviar, ja_enviados, bloqueados[(titulo, motivo)])."""
    enviar, repetidos, bloqueados = [], [], []
    vistos = {}
    for v in veiculos:
        t = (v.get("titulo_contratacao") or "").strip()
        if not t:
            continue
        chave = chave_carga(t, v.get("numeros_unicos"))
        n = vistos[t] = vistos.get(t, 0) + 1
        mostrar = t if n == 1 else f"{t} ({n}ª carga igual)"   # cargas iguais com nºs únicos diferentes
        CHAVES[mostrar] = chave
        DADOS[mostrar] = v
        if RE_AGENDADO_SEM_DATA.search(t):
            bloqueados.append((mostrar, "AGENDADO sem data — aguardando o PCP"))
        elif v.get("precisa_customer_service"):
            bloqueados.append((mostrar, "Sem Tipo de Entrega — acionar Customer Service"))
        elif chave in enviados:
            repetidos.append((mostrar, enviados[chave]))
        elif t in enviados and chave != t:   # registro antigo (antes de guardar os nºs únicos)
            repetidos.append((mostrar, enviados[t] + " — título igual; confira se é a mesma carga"))
        elif mostrar not in enviar:
            enviar.append(mostrar)
    return enviar, repetidos, bloqueados


def montar_mensagem(titulos, cfg, com_dados=False):
    h = datetime.now().hour
    saud = "Bom dia" if h < 12 else ("Boa tarde" if h < 18 else "Boa noite")
    cab = cfg["cabecalho"].format(saudacao=saud, data=datetime.now().strftime("%d/%m"))
    if com_dados:
        partes = [cab]
        for t in titulos:
            partes += ["", linha_contratacao(DADOS.get(t), t)]
            d = texto_dados(DADOS.get(t), "   ")
            if d:
                partes.append(d)
    else:
        partes = [cab, ""] + [linha_contratacao(DADOS.get(t), t) for t in titulos]
    if cfg["rodape"]:
        partes += ["", cfg["rodape"]]
    return "\n".join(partes)


def base_mais_recente():
    arqs = [p for ext in ("*.xls", "*.xlsx") for p in glob.glob(str(BASE / ext))
            if "roteirizacao" in os.path.basename(p).lower().replace("ç", "c").replace("ã", "a")
            and not os.path.basename(p).startswith("~$")]
    return max(arqs, key=os.path.getmtime) if arqs else None


def processar_base(abrir_relatorio):
    """Roda o gerador de relatório e devolve os veículos de lotação."""
    import gerar_relatorio_roteirizacao as g
    g.ABRIR_NO_NAVEGADOR = abrir_relatorio
    g.main()  # gera o HTML exatamente como antes
    caminho = g.localizar_base_mais_recente(g.PASTA_SAIDA)
    dados = g.montar_dados(g.carregar_base(caminho))
    return os.path.basename(caminho), dados["lotacao"]["veiculos"]


def gravar_roteiros(veiculos):
    """Guarda o nome do cliente de cada carga (aparece no Histórico)."""
    try:
        d = json.loads(ROTEIROS.read_text(encoding="utf-8"))
    except Exception:
        d = {}
    for v in veiculos:
        tc, tr = (v.get("titulo_contratacao") or "").strip(), (v.get("titulo_roteirizacao") or "").strip()
        if tc and tr:
            d[tc] = tr
    try:
        ROTEIROS.write_text(json.dumps(dict(list(d.items())[-3000:]), ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError:
        pass


REGISTRO = BASE / "REGISTRO_ENVIOS_CONTRATACAO.csv"   # 1 linha por carga enviada (abre no Excel)


def registrar_cargas(contato, empresa, titulos, ok, detalhe=""):
    """Registro fácil de ler: data, hora, para quem, empresa, carga, cliente, resultado."""
    try:
        clientes = json.loads(ROTEIROS.read_text(encoding="utf-8"))
    except Exception:
        clientes = {}
    novo = not REGISTRO.exists()
    try:
        with open(REGISTRO, "a", encoding="utf-8-sig") as f:
            if novo:
                f.write("data;hora;enviado_para;transportadora;carga;cliente_roteiro;resultado;detalhe\n")
            agora = datetime.now()
            for tt in titulos:
                base_t = re.sub(r"\s*\(\d+ª carga igual\)\s*$", "", tt)
                cli = str(clientes.get(base_t, "")).replace(";", ",")
                f.write(f"{agora:%d/%m/%Y};{agora:%H:%M:%S};{contato};{empresa};{tt.replace(';', ',')};{cli};"
                        f"{'ENVIADO' if ok else 'FALHOU'};{str(detalhe).replace(';', ',')[:120]}\n")
    except Exception:
        pass


def registrar(evento, contato, detalhe):
    novo = not LOG.exists()
    try:
        with open(LOG, "a", encoding="utf-8-sig") as f:
            if novo:
                f.write("data_hora;evento;contato;detalhe\n")
            f.write(f"{datetime.now():%d/%m/%Y %H:%M:%S};{evento};{contato};{detalhe.replace(';', ',')}\n")
    except Exception:
        pass


# ------------------------------------------------------------------ interface
class Robo(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Robô de Contratação")
        self.configure(bg=FUNDO)
        self.geometry("1200x460")
        self.fila_ui = queue.Queue()
        self.fila_whats = queue.Queue()
        self.ultimo_mtime = None
        self.processando = False

        tk.Label(self, text="Robô de Contratação", bg=FUNDO, fg=TEXTO,
                 font=(FONTE, 17, "bold")).pack(anchor="w", padx=16, pady=(14, 0))
        self.lbl_status = tk.Label(self, text="● vigiando a pasta", bg=FUNDO, fg=VERDE,
                                   font=(FONTE, 10, "bold"), anchor="w")
        self.lbl_status.pack(fill="x", padx=16)
        tk.Label(self, text=f"Pasta: {BASE}", bg=FUNDO, fg=FRACO, font=(FONTE, 9),
                 anchor="w").pack(fill="x", padx=16)

        self.txt = tk.Text(self, bg=CARTAO, fg=TEXTO, font=("Consolas", 9), relief="flat",
                           height=14, wrap="word", state="disabled")
        self.txt.pack(fill="both", expand=True, padx=16, pady=10)

        rod = tk.Frame(self, bg=FUNDO)
        rod.pack(fill="x", padx=16, pady=(0, 12))
        self.ultimos_veiculos = []
        for texto, cmd in (("🚚 Contratar transportadora", self.janela_contratar),
                           ("🔎 Transportadoras do WhatsApp", self.buscar_transportadoras),
                           ("🗺 Regiões", self.janela_regioes),
                           ("📋 Dados das cargas", self.janela_dados),
                           ("📦 Coletas aceitas", self.janela_coletas),
                           ("📜 Histórico", self.abrir_historico),
                           ("Processar base agora", lambda: self.disparar(forcar=True)),
                           ("Abrir WhatsApp / QR Code", lambda: self.fila_whats.put(("abrir", None))),
                           ("Editar contatos", self.editar_contatos)):
            b = tk.Label(rod, text=texto, bg="#2a3f2c", fg=TEXTO, font=(FONTE, 10),
                         padx=10, pady=6, cursor="hand2")
            b.bind("<Button-1>", lambda e, c=cmd: c())
            b.pack(side="left", padx=(0, 6))

        self.protocol("WM_DELETE_WINDOW", self.fechar)
        base = base_mais_recente()
        self.ultimo_mtime = os.path.getmtime(base) if base else None
        self.escrever("Robô iniciado. Salve uma Base Roteirizacao nova nesta pasta "
                      "(ou clique em 'Processar base agora').")
        threading.Thread(target=self._vigiar, daemon=True).start()
        threading.Thread(target=self._trabalhador_whats, daemon=True).start()
        self.after(300, self._ui_loop)
        self.janela_lembrete = None
        if ler_config()["vigiar_respostas"]:
            self.fila_whats.put(("abrir", None))  # já abre o WhatsApp pra vigiar as respostas
            self.escrever("Vigiando as respostas das transportadoras: quando disserem que COLETAM, "
                          "eu aviso e te lembro de mandar o OK depois de formar a carga.")
        self.after(60000, self._lembrete_ok)

    # ---- utilidades de tela
    def escrever(self, msg):
        self.txt.configure(state="normal")
        self.txt.insert("end", f"[{datetime.now():%H:%M:%S}] {msg}\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def log_thread(self, msg):
        self.fila_ui.put(("log", msg))

    def _ui_loop(self):
        try:
            while True:
                tipo, dado = self.fila_ui.get_nowait()
                if tipo == "log":
                    self.escrever(dado)
                elif tipo == "confirmar":
                    self.janela_confirmacao(*dado)
                elif tipo == "status":
                    self.lbl_status.configure(text=dado[0], fg=dado[1])
                elif tipo == "aceite":
                    self.janela_aceite(*dado)
                elif tipo == "comprovante":
                    self.janela_comprovante(dado)
                elif tipo == "conversas":
                    self.janela_buscar_transportadoras(dado)
        except queue.Empty:
            pass
        self.after(300, self._ui_loop)

    def janela_comprovante(self, resultados):
        """Depois de enviar: mostra PARA QUEM foi cada carga (fica registrado em REGISTRO_ENVIOS_CONTRATACAO.csv)."""
        try:
            import historico_contratacao as H
            H.gerar()  # atualiza o Historico_Contratacao.xlsx (se não estiver aberto no Excel)
        except Exception:
            pass
        w = tk.Toplevel(self)
        w.title("Comprovante de envio")
        w.configure(bg=FUNDO)
        w.attributes("-topmost", True)
        w.geometry("760x" + str(min(640, 200 + 22 * sum(len(r[2]) + 2 for r in resultados))))
        falhou = [r for r in resultados if not r[3]]
        tk.Label(w, text=("✖ ATENÇÃO: algum envio FALHOU" if falhou else "✔ Enviado no WhatsApp") +
                 f"  ·  {datetime.now():%d/%m/%Y %H:%M}", bg=(VERMELHO if falhou else VERDE), fg=FUNDO,
                 font=(FONTE, 12, "bold")).pack(fill="x", ipady=6)
        txt = tk.Text(w, bg=CARTAO, fg=TEXTO, font=(FONTE, 10), relief="flat", wrap="word")
        txt.pack(fill="both", expand=True, padx=12, pady=8)
        linhas = []
        for nome, empresa, titulos, ok, det in resultados:
            linhas.append(f"{'✔' if ok else '✖ NÃO ENVIADO —'} {nome} ({empresa}) — {len(titulos)} carga(s)")
            linhas += [f"     • {tt}" for tt in titulos]
            if not ok:
                linhas.append(f"     motivo: {det}")
            linhas.append("")
        txt.insert("1.0", "\n".join(linhas))
        txt.configure(state="disabled")
        b = tk.Frame(w, bg=FUNDO)
        b.pack(fill="x", padx=12, pady=(0, 10))

        def copiar():
            self.clipboard_clear()
            self.clipboard_append("\n".join(linhas))

        tk.Button(b, text="📋 Copiar", command=copiar, bg="#2a3f2c", fg=TEXTO, relief="flat", padx=12, pady=5).pack(side="left")
        tk.Button(b, text="📜 Abrir registro (Excel)", command=lambda: os.startfile(str(REGISTRO)) if REGISTRO.exists() else None,
                  bg="#2a3f2c", fg=TEXTO, relief="flat", padx=12, pady=5).pack(side="left", padx=6)
        tk.Button(b, text="OK, conferi", command=w.destroy, bg=VERDE, fg=FUNDO, relief="flat",
                  font=(FONTE, 10, "bold"), padx=14, pady=5).pack(side="right")
        for nome, empresa, titulos, ok, det in resultados:
            self.escrever(f"📨 {'ENVIADO' if ok else 'FALHOU'} para {nome} ({empresa}): {len(titulos)} carga(s) — "
                          "registrado em REGISTRO_ENVIOS_CONTRATACAO.csv")

    def abrir_historico(self):
        try:
            import historico_contratacao as H
            arq, n = H.gerar()
            self.escrever(f"Histórico gerado: {arq.name} ({n} carga(s)). Abrindo no Excel…")
            os.startfile(str(arq))
        except PermissionError:
            self.escrever("✖ O Historico_Contratacao.xlsx está aberto no Excel. Feche e clique de novo.")
        except Exception as e:
            self.escrever(f"✖ Erro ao gerar o histórico: {e}")

    def editar_contatos(self):
        try:
            subprocess.Popen(["notepad.exe", str(CONFIG)])
        except Exception:
            os.startfile(str(CONFIG))

    # ---- vigia da pasta
    def _vigiar(self):
        while True:
            time.sleep(INTERVALO)
            base = base_mais_recente()
            if not base:
                continue
            mt = os.path.getmtime(base)
            if self.ultimo_mtime is None or mt > self.ultimo_mtime:
                time.sleep(3)  # espera terminar de salvar
                self.ultimo_mtime = os.path.getmtime(base)
                self.disparar()

    def disparar(self, forcar=False):
        if self.processando:
            return
        self.processando = True
        threading.Thread(target=self._processar, args=(forcar,), daemon=True).start()

    def _processar(self, forcar):
        try:
            cfg = ler_config()
            self.fila_ui.put(("status", ("● gerando relatório…", AMBAR)))
            nome_base, veiculos = processar_base(cfg["abrir_relatorio"])
            self.ultimos_veiculos = veiculos
            gravar_roteiros(veiculos)
            enviar, repetidos, bloqueados = separar_titulos(veiculos, ler_enviados(cfg["dias"]))
            self.log_thread(f"Base lida: {nome_base} — {len(veiculos)} veículo(s) de lotação; "
                            f"{len(enviar)} novo(s) para contratação, {len(bloqueados)} bloqueado(s), "
                            f"{len(repetidos)} já enviado(s).")
            for t, motivo in bloqueados:
                self.log_thread(f"   não enviado: {t}  ({motivo})")
            if enviar or forcar:
                self.fila_ui.put(("confirmar", (enviar, repetidos, bloqueados, cfg)))
            else:
                self.log_thread("Nada novo para enviar.")
        except BaseException as e:  # o gerador usa SystemExit para erros de base
            self.log_thread(f"ERRO ao processar a base: {e}")
        finally:
            self.processando = False
            self.fila_ui.put(("status", ("● vigiando a pasta", VERDE)))

    # ---- dados de todas as cargas (paletes, peso, valor, agenda, coleta) para copiar e mandar
    def janela_dados(self, titulos=None):
        if titulos is None:
            vs = [v for v in self.ultimos_veiculos if (v.get("titulo_contratacao") or "").strip()]
            itens = [((v.get("titulo_contratacao") or "").strip(), v) for v in vs]
        else:
            itens = [(t, DADOS.get(t)) for t in titulos]
        if not itens:
            self.escrever("Ainda não tenho cargas: salve uma Base Roteirizacao ou clique em 'Processar base agora'.")
            return
        blocos = [f"• {t}\n{texto_dados(v, '   ')}" for t, v in itens]
        texto = f"DADOS DAS CARGAS · {len(itens)} carga(s)\n\n" + "\n\n".join(blocos)
        w = tk.Toplevel(self)
        w.title("Dados das cargas")
        w.configure(bg=FUNDO)
        w.geometry("860x560")
        w.attributes("-topmost", True)
        tk.Label(w, text=f"Dados das cargas ({len(itens)})", bg=FUNDO, fg=TEXTO,
                 font=(FONTE, 14, "bold")).pack(anchor="w", padx=16, pady=(12, 0))
        tk.Label(w, text="Paletes, peso, valor da nota, entrega agendada e data para coletar de cada carga. "
                         "Clique em Copiar e cole no WhatsApp, Teams ou e-mail.",
                 bg=FUNDO, fg=FRACO, font=(FONTE, 9), anchor="w").pack(fill="x", padx=16)
        caixa = tk.Text(w, bg=CARTAO, fg=TEXTO, font=("Consolas", 10), relief="flat", wrap="word")
        caixa.insert("1.0", texto)
        caixa.pack(fill="both", expand=True, padx=16, pady=10)
        bot = tk.Frame(w, bg=FUNDO)
        bot.pack(fill="x", padx=16, pady=(0, 12))
        msg = tk.Label(bot, text="", bg=FUNDO, fg=VERDE, font=(FONTE, 10, "bold"))

        def copiar():
            self.clipboard_clear()
            self.clipboard_append(caixa.get("1.0", "end").strip())
            msg.configure(text="✔ Copiado!")
        tk.Button(bot, text="📋 COPIAR TUDO", command=copiar, bg=VERDE, fg=FUNDO, font=(FONTE, 11, "bold"),
                  relief="flat", padx=16, pady=6).pack(side="left")
        msg.pack(side="left", padx=10)
        tk.Button(bot, text="Fechar", command=w.destroy, bg="#2a3f2c", fg=TEXTO, relief="flat",
                  padx=12, pady=6).pack(side="right")

    # ---- REGIÕES QUE CADA TRANSPORTADORA ATENDE
    def janela_regioes(self):
        cfg = ler_config()
        regioes = cfg["regioes"]
        empresas = list(dict.fromkeys([e.upper() for e in cfg["transportadoras"]] + list(regioes)))
        w = tk.Toplevel(self)
        w.title("Regiões das transportadoras")
        w.configure(bg=FUNDO)
        w.geometry("1150x560")
        w.attributes("-topmost", True)
        tk.Label(w, text="🗺 Para onde cada transportadora atende", bg=FUNDO, fg=TEXTO,
                 font=(FONTE, 14, "bold")).pack(anchor="w", padx=16, pady=(12, 0))
        tk.Label(w, text="Marque as regiões. O robô usa isso para já indicar (e marcar) a transportadora certa para cada carga. "
                         "SP-CAPITAL = Grande São Paulo; SP-INTERIOR = resto do estado. Em 'Outros' digite siglas (ex.: RJ, ES, TO).",
                 bg=FUNDO, fg=FRACO, font=(FONTE, 9), anchor="w", wraplength=1100, justify="left").pack(fill="x", padx=16)
        quadro = tk.Frame(w, bg=CARTAO)
        quadro.pack(fill="both", expand=True, padx=16, pady=8)
        canvas = tk.Canvas(quadro, bg=CARTAO, highlightthickness=0)
        barra = ttk.Scrollbar(quadro, orient="vertical", command=canvas.yview)
        corpo = tk.Frame(canvas, bg=CARTAO)
        corpo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=corpo, anchor="nw")
        canvas.configure(yscrollcommand=barra.set)
        canvas.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        tk.Label(corpo, text="Transportadora", bg=CARTAO, fg=FRACO, font=(FONTE, 9, "bold")).grid(row=0, column=0, sticky="w", padx=6, pady=4)
        for j, op in enumerate(OPCOES_REGIAO):
            tk.Label(corpo, text=op, bg=CARTAO, fg=VERDE, font=(FONTE, 8, "bold")).grid(row=0, column=1 + j, padx=3)
        tk.Label(corpo, text="Outros", bg=CARTAO, fg=FRACO, font=(FONTE, 9, "bold")).grid(row=0, column=1 + len(OPCOES_REGIAO), padx=4)
        linhas = {}

        def linha(emp, r):
            atual = [x.upper() for x in regioes.get(emp, [])]
            if "SP" in atual:
                atual = [x for x in atual if x != "SP"] + ["SP-CAPITAL", "SP-INTERIOR"]
            tk.Label(corpo, text=emp, bg=CARTAO, fg=TEXTO, font=(FONTE, 10, "bold"), anchor="w").grid(row=r, column=0, sticky="w", padx=6, pady=2)
            vs = {}
            for j, op in enumerate(OPCOES_REGIAO):
                bv = tk.BooleanVar(value=op in atual)
                tk.Checkbutton(corpo, variable=bv, bg=CARTAO, selectcolor=FUNDO, activebackground=CARTAO,
                               highlightthickness=0).grid(row=r, column=1 + j)
                vs[op] = bv
            outros = tk.StringVar(value=", ".join(x for x in atual if x not in OPCOES_REGIAO))
            tk.Entry(corpo, textvariable=outros, width=14).grid(row=r, column=1 + len(OPCOES_REGIAO), padx=4)
            linhas[emp] = (vs, outros)
        for i, emp in enumerate(empresas, 1):
            linha(emp, i)
        bot = tk.Frame(w, bg=FUNDO)
        bot.pack(fill="x", padx=16, pady=(0, 12))
        nova = tk.StringVar()
        tk.Label(bot, text="Nova transportadora:", bg=FUNDO, fg=TEXTO, font=(FONTE, 9)).pack(side="left")
        tk.Entry(bot, textvariable=nova, width=22).pack(side="left", padx=4)

        def incluir():
            n = nova.get().strip().upper()
            if n and n not in linhas:
                linha(n, len(linhas) + 1)
                nova.set("")
        tk.Button(bot, text="+ incluir", command=incluir, bg="#2a3f2c", fg=TEXTO, relief="flat", padx=8, pady=3).pack(side="left")
        aviso = tk.Label(bot, text="", bg=FUNDO, fg=VERDE, font=(FONTE, 10, "bold"))
        aviso.pack(side="left", padx=10)

        def salvar():
            mapa = {}
            for emp, (vs, outros) in linhas.items():
                r = [op for op, bv in vs.items() if bv.get()]
                if "SP-CAPITAL" in r and "SP-INTERIOR" in r:
                    r = [x for x in r if x not in ("SP-CAPITAL", "SP-INTERIOR")] + ["SP"]
                r += [x.strip().upper() for x in outros.get().split(",") if x.strip()]
                mapa[emp] = list(dict.fromkeys(r))
            gravar_regioes(mapa)
            novas = [e for e in mapa if mapa[e] and e not in [t.upper() for t in ler_config()["transportadoras"]]]
            if novas:
                adicionar_contatos_ini([("", e) for e in novas])
            aviso.configure(text="✔ Regiões salvas. Já valem no envio e no Contratar transportadora.")
            self.escrever("🗺 Regiões das transportadoras salvas: " + "; ".join(f"{e}: {', '.join(r)}" for e, r in mapa.items() if r))
        tk.Button(bot, text="Fechar", command=w.destroy, bg="#2a3f2c", fg=TEXTO, relief="flat", padx=12, pady=6).pack(side="right", padx=(6, 0))
        tk.Button(bot, text="SALVAR", command=salvar, bg=VERDE, fg=FUNDO, font=(FONTE, 10, "bold"), relief="flat",
                  padx=16, pady=6).pack(side="right")

    # ---- BUSCAR TRANSPORTADORAS NO WHATSAPP (só nomes das conversas)
    def buscar_transportadoras(self):
        self.escrever("🔎 Lendo os NOMES das conversas do WhatsApp (sem abrir mensagens)… pode levar 1 minuto.")
        self.fila_whats.put(("listar", None))

    def janela_buscar_transportadoras(self, nomes):
        cfg = ler_config()
        ja = {_sem_acento(c["nome"]).strip() for c in cfg["contatos"]}
        empresas = cfg["transportadoras"]
        w = tk.Toplevel(self)
        w.title("Transportadoras no WhatsApp")
        w.configure(bg=FUNDO)
        w.geometry("820x600")
        w.attributes("-topmost", True)
        tk.Label(w, text="🔎 Transportadoras encontradas no seu WhatsApp", bg=FUNDO, fg=TEXTO,
                 font=(FONTE, 14, "bold")).pack(anchor="w", padx=16, pady=(12, 0))
        tk.Label(w, text="Mostro só conversas com nome de transportadora/logística (cargas, transp, log, Alfa, Cooper, Rodo…). "
                         "Marque as que quer usar, confira a empresa e clique em ADICIONAR.",
                 bg=FUNDO, fg=FRACO, font=(FONTE, 9), anchor="w", wraplength=780, justify="left").pack(fill="x", padx=16)
        topo = tk.Frame(w, bg=FUNDO)
        topo.pack(fill="x", padx=16, pady=6)
        tk.Label(topo, text="Procurar outro nome:", bg=FUNDO, fg=TEXTO, font=(FONTE, 9)).pack(side="left")
        busca = tk.StringVar()
        tk.Entry(topo, textvariable=busca, width=28).pack(side="left", padx=6)
        quadro = tk.Frame(w, bg=CARTAO)
        quadro.pack(fill="both", expand=True, padx=16, pady=4)
        canvas = tk.Canvas(quadro, bg=CARTAO, highlightthickness=0)
        barra = ttk.Scrollbar(quadro, orient="vertical", command=canvas.yview)
        corpo = tk.Frame(canvas, bg=CARTAO)
        corpo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=corpo, anchor="nw")
        canvas.configure(yscrollcommand=barra.set)
        canvas.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        escolhas = {}  # nome -> (BooleanVar, StringVar)

        def montar(*_):
            for f in corpo.winfo_children():
                f.destroy()
            termo = _sem_acento(busca.get()).strip()
            lin = 0
            for nome in nomes:
                eh, emp = parece_transportadora(nome)
                if termo and len(termo) >= 3:
                    if termo not in _sem_acento(nome):
                        continue
                elif not eh:
                    continue
                if nome not in escolhas:
                    escolhas[nome] = (tk.BooleanVar(value=False), tk.StringVar(value=emp))
                bv, ev = escolhas[nome]
                cadastrado = _sem_acento(nome).strip() in ja
                tk.Checkbutton(corpo, text=nome + ("   (já cadastrado)" if cadastrado else ""), variable=bv,
                               state=("disabled" if cadastrado else "normal"), bg=CARTAO, fg=TEXTO,
                               selectcolor=FUNDO, activebackground=CARTAO, activeforeground=TEXTO, anchor="w",
                               font=(FONTE, 10), highlightthickness=0, width=48).grid(row=lin, column=0, sticky="w", padx=6, pady=1)
                ttk.Combobox(corpo, textvariable=ev, values=empresas, width=22).grid(row=lin, column=1, padx=6)
                lin += 1
            if lin == 0:
                tk.Label(corpo, text="Nenhuma conversa com nome de transportadora. Use 'Procurar outro nome'.",
                         bg=CARTAO, fg=FRACO, font=(FONTE, 10)).grid(row=0, column=0, padx=10, pady=10)
        busca.trace_add("write", montar)
        montar()
        bot = tk.Frame(w, bg=FUNDO)
        bot.pack(fill="x", padx=16, pady=(4, 12))
        aviso = tk.Label(bot, text="", bg=FUNDO, fg=VERDE, font=(FONTE, 10, "bold"))
        aviso.pack(side="left")

        def adicionar():
            novos = [(n, ev.get().strip().upper() or n) for n, (bv, ev) in escolhas.items() if bv.get()]
            if not novos:
                aviso.configure(text="Marque pelo menos um contato.", fg=AMBAR)
                return
            q = adicionar_contatos_ini(novos)
            self.escrever(f"✔ {q} contato(s) de transportadora adicionado(s): " + ", ".join(f"{n} ({e})" for n, e in novos))
            aviso.configure(text=f"✔ {q} adicionado(s). Já aparecem no envio e no 'Contratar transportadora'.", fg=VERDE)
        tk.Button(bot, text="Fechar", command=w.destroy, bg="#2a3f2c", fg=TEXTO, relief="flat",
                  padx=12, pady=6).pack(side="right", padx=(6, 0))
        tk.Button(bot, text="ADICIONAR SELECIONADOS", command=adicionar, bg=VERDE, fg=FUNDO,
                  font=(FONTE, 10, "bold"), relief="flat", padx=14, pady=6).pack(side="right")

    # ---- CONTRATAR TRANSPORTADORA (substitui o app JCLOG)
    def janela_contratar(self, titulos=None):
        """Lista as cargas de contratação com: transportadora (escolher), botão COPIAR com os dados do
        caminhão (cidade, estado, peso, carreta/truck, entrega) e botão CONTRATADA para registrar."""
        if titulos is None:
            enviar, repetidos, bloqueados = separar_titulos(self.ultimos_veiculos, {})
            titulos = enviar + [t for t, _ in bloqueados if "AGENDADO sem data" not in _]
        if not titulos:
            self.escrever("Ainda não tenho cargas: salve uma Base Roteirizacao ou clique em 'Processar base agora'.")
            return
        cfg = ler_config()
        transp = cfg["transportadoras"]
        feitas = ler_contratadas()
        w = tk.Toplevel(self)
        w.title("Contratar transportadora")
        w.configure(bg=FUNDO)
        w.geometry("1150x640")
        w.attributes("-topmost", True)
        tk.Label(w, text=f"🚚 Contratar transportadora · {len(titulos)} carga(s)", bg=FUNDO, fg=TEXTO,
                 font=(FONTE, 15, "bold")).pack(anchor="w", padx=16, pady=(12, 0))
        tk.Label(w, text="1) Clique em COPIAR e cole no WhatsApp/Teams/e-mail da transportadora.  "
                         "2) Quando ela aceitar, escolha a transportadora e clique em CONTRATADA (fica registrado).",
                 bg=FUNDO, fg=FRACO, font=(FONTE, 9), anchor="w").pack(fill="x", padx=16)
        filtro = tk.StringVar(value="pendentes")
        topo = tk.Frame(w, bg=FUNDO)
        topo.pack(fill="x", padx=16, pady=(6, 0))
        quadro = tk.Frame(w, bg=CARTAO)
        quadro.pack(fill="both", expand=True, padx=16, pady=8)
        canvas = tk.Canvas(quadro, bg=CARTAO, highlightthickness=0)
        barra = ttk.Scrollbar(quadro, orient="vertical", command=canvas.yview)
        corpo = tk.Frame(canvas, bg=CARTAO)
        corpo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=corpo, anchor="nw")
        canvas.configure(yscrollcommand=barra.set)
        canvas.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))
        aviso = tk.Label(w, text="", bg=FUNDO, fg=VERDE, font=(FONTE, 10, "bold"), anchor="w")

        def copiar(texto, o_que):
            self.clipboard_clear()
            self.clipboard_append(texto)
            aviso.configure(text=f"✔ Copiado: {o_que}")

        def montar():
            for f in corpo.winfo_children():
                f.destroy()
            feitas.clear()
            feitas.update(ler_contratadas())
            cab = ["Carga", "Dados do caminhão", "", "Transportadora", "", "Situação"]
            for j, c in enumerate(cab):
                tk.Label(corpo, text=c, bg=CARTAO, fg=FRACO, font=(FONTE, 9, "bold")).grid(row=0, column=j, sticky="w", padx=6, pady=6)
            lin = 1
            for t in titulos:
                v = DADOS.get(t) or {}
                chave = CHAVES.get(t, t)
                ja = feitas.get(chave)
                if filtro.get() == "pendentes" and ja:
                    continue
                if filtro.get() == "contratadas" and not ja:
                    continue
                dados = linha_contratacao(v, t)
                tk.Label(corpo, text=t + (f"\n{v.get('nome', '')}" if v.get("nome") else ""), bg=CARTAO,
                         fg=(FRACO if ja else TEXTO), font=(FONTE, 9, "bold"), anchor="w", justify="left",
                         wraplength=300).grid(row=lin, column=0, sticky="nw", padx=6, pady=4)
                tk.Label(corpo, text=dados, bg=CARTAO, fg=TEXTO, font=("Consolas", 9), anchor="w",
                         justify="left").grid(row=lin, column=1, sticky="nw", padx=6, pady=4)
                tk.Button(corpo, text="📋 COPIAR", command=lambda d=dados, tt=t: copiar(d, tt),
                          bg=VERDE, fg=FUNDO, font=(FONTE, 9, "bold"), relief="flat", padx=10, pady=3).grid(row=lin, column=2, padx=6, sticky="n", pady=4)
                sug = transportadoras_sugeridas(v, t, cfg)
                var = tk.StringVar(value=(ja or {}).get("transportadora", "") or (sug[0] if sug else ""))
                if sug and not ja:
                    tk.Label(corpo, text="sugerido: " + " / ".join(sug), bg=CARTAO, fg=AMBAR, font=(FONTE, 8)).grid(
                        row=lin, column=3, sticky="s", padx=6)
                cb = ttk.Combobox(corpo, textvariable=var, values=transp, width=20)
                cb.grid(row=lin, column=3, padx=6, sticky="n", pady=6)
                _lbl_carga = corpo.grid_slaves(row=lin, column=0)[0]

                def _pinta(*_, lb=_lbl_carga, vv=var, jj=ja):
                    if not jj:
                        lb.configure(bg=("#7a1f1f" if vv.get().strip() else CARTAO), fg=("#ffffff" if vv.get().strip() else TEXTO))
                var.trace_add("write", _pinta)
                _pinta()

                def marcar(tt=t, ch=chave, vv=var):
                    nome = vv.get().strip()
                    if not nome:
                        aviso.configure(text="Escolha a transportadora antes de marcar como contratada.", fg=AMBAR)
                        return
                    gravar_contratada(ch, nome, tt)
                    registrar_cargas(nome, nome, [tt], True, "contratada (sem JCLOG)")
                    registrar("contratada", nome, tt)
                    self.escrever(f"🚚 Contratada com {nome}: {tt}")
                    aviso.configure(text=f"✔ {tt} → {nome}", fg=VERDE)
                    montar()

                def desfazer(tt=t, ch=chave):
                    gravar_contratada(ch, "", tt)
                    self.escrever(f"↩ Contratação desfeita: {tt}")
                    montar()
                if ja:
                    tk.Button(corpo, text="↩ Desfazer", command=desfazer, bg="#2a3f2c", fg=TEXTO, relief="flat",
                              padx=8, pady=3).grid(row=lin, column=4, padx=6, sticky="n", pady=4)
                    sit = f"✔ {ja['transportadora']}\n{ja['quando']}"
                    cor = VERDE
                else:
                    tk.Button(corpo, text="✔ CONTRATADA", command=marcar, bg="#2a3f2c", fg=TEXTO, relief="flat",
                              font=(FONTE, 9, "bold"), padx=8, pady=3).grid(row=lin, column=4, padx=6, sticky="n", pady=4)
                    sit, cor = "pendente", AMBAR
                tk.Label(corpo, text=sit, bg=CARTAO, fg=cor, font=(FONTE, 9, "bold"), justify="left").grid(row=lin, column=5, sticky="nw", padx=6, pady=4)
                lin += 1
            if lin == 1:
                tk.Label(corpo, text="Nada neste filtro.", bg=CARTAO, fg=FRACO, font=(FONTE, 10)).grid(row=1, column=0, columnspan=6, padx=10, pady=10)

        for txt, val in (("Pendentes", "pendentes"), ("Contratadas", "contratadas"), ("Todas", "todas")):
            tk.Radiobutton(topo, text=txt, variable=filtro, value=val, command=montar, bg=FUNDO, fg=TEXTO,
                           selectcolor=CARTAO, activebackground=FUNDO, activeforeground=TEXTO,
                           font=(FONTE, 10), highlightthickness=0).pack(side="left", padx=(0, 12))

        def copiar_pendentes():
            blocos = [linha_contratacao(DADOS.get(t), t) for t in titulos if not ler_contratadas().get(CHAVES.get(t, t))]
            if blocos:
                copiar("\n".join(blocos), f"{len(blocos)} carga(s) pendente(s)")
        tk.Button(topo, text="📋 Copiar TODAS as pendentes", command=copiar_pendentes, bg=VERDE, fg=FUNDO,
                  font=(FONTE, 10, "bold"), relief="flat", padx=12, pady=4).pack(side="right")
        montar()
        bot = tk.Frame(w, bg=FUNDO)
        bot.pack(fill="x", padx=16, pady=(0, 12))
        aviso.pack(in_=bot, side="left")

        def fechar():
            canvas.unbind_all("<MouseWheel>")
            w.destroy()
        w.protocol("WM_DELETE_WINDOW", fechar)
        tk.Button(bot, text="✏ Editar lista de transportadoras", command=self.editar_contatos, bg="#2a3f2c", fg=TEXTO,
                  relief="flat", padx=10, pady=5).pack(side="right", padx=(6, 0))
        tk.Button(bot, text="Fechar", command=fechar, bg="#2a3f2c", fg=TEXTO, relief="flat",
                  padx=12, pady=5).pack(side="right")

    # ---- confirmação: cada título com uma caixinha por contato
    def janela_confirmacao(self, enviar, repetidos, bloqueados, cfg):
        contatos = cfg["contatos"]
        memoria = ler_estados()
        w = tk.Toplevel(self)
        w.title("Confirmar envio para contratação")
        w.configure(bg=FUNDO)
        w.geometry("900x620")
        w.attributes("-topmost", True)
        try:
            import winsound
            winsound.MessageBeep()
        except Exception:
            self.bell()
        tk.Label(w, text="Conferir antes de enviar", bg=FUNDO, fg=TEXTO,
                 font=(FONTE, 15, "bold")).pack(anchor="w", padx=16, pady=(12, 0))
        tk.Label(w, text="Em cada carga, clique em ▼ e escolha para qual transportadora vai (pode marcar mais de uma). "
                         "Já vem escolhida pela região / pelo que você usou da última vez para aquele estado.",
                 bg=FUNDO, fg=FRACO, font=(FONTE, 9), anchor="w").pack(fill="x", padx=16)

        # área rolável (pode ter muitas cargas)
        quadro = tk.Frame(w, bg=CARTAO)
        quadro.pack(fill="both", expand=True, padx=16, pady=10)
        canvas = tk.Canvas(quadro, bg=CARTAO, highlightthickness=0)
        barra = ttk.Scrollbar(quadro, orient="vertical", command=canvas.yview)
        corpo = tk.Frame(canvas, bg=CARTAO)
        corpo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=corpo, anchor="nw")
        canvas.configure(yscrollcommand=barra.set)
        canvas.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))

        # cabeçalho da tabela
        tk.Label(corpo, text="UF", bg=CARTAO, fg=FRACO, font=(FONTE, 9, "bold")).grid(row=0, column=0, padx=(10, 4), pady=6)
        tk.Label(corpo, text="Carga (título de contratação)", bg=CARTAO, fg=FRACO,
                 font=(FONTE, 9, "bold"), anchor="w").grid(row=0, column=1, sticky="w")
        tk.Label(corpo, text="Enviar para", bg=CARTAO, fg=VERDE,
                 font=(FONTE, 9, "bold")).grid(row=0, column=2, padx=10, sticky="w")

        linhas = []  # (titulo, uf, {nome: BooleanVar})
        lin = 1
        for titulo, quando in [(t, None) for t in enviar] + list(repetidos):
            uf = uf_do_titulo(titulo)
            lembrado = memoria.get(uf)
            lbl_uf = tk.Label(corpo, text=uf, bg=CARTAO, fg=(AMBAR if lembrado is None else TEXTO),
                              font=(FONTE, 10, "bold"))
            lbl_uf.grid(row=lin, column=0, padx=(10, 4), sticky="nsew")
            rotulo = titulo if quando is None else f"{titulo}\n(já enviado em {quando})"
            dd = DADOS.get(titulo) or {}
            if dd:
                rotulo += (f"\n   paletes {_br(dd.get('pallets') or 0)} · {dd.get('valor', '')} · "
                           f"entrega {dd.get('agenda_completa') or dd.get('agenda') or 'sem agenda'} · coletar até {_coleta(dd)}")
            _sg = transportadoras_sugeridas(DADOS.get(titulo), titulo, cfg)
            if _sg:
                rotulo += "\n   → atende: " + " / ".join(_sg)
            lbl_rot = tk.Label(corpo, text=rotulo, bg=CARTAO, fg=(TEXTO if quando is None else FRACO),
                               font=(FONTE, 10), anchor="w", justify="left")
            lbl_rot.grid(row=lin, column=1, sticky="nsew", pady=2)
            marcas = {}
            sug_reg = transportadoras_sugeridas(DADOS.get(titulo), titulo, cfg)
            tem_contato = any(_mesma_empresa(cc.get("empresa", ""), e) for cc in contatos for e in sug_reg)
            for c in contatos:
                por_regiao = any(_mesma_empresa(c.get("empresa", ""), e) for e in sug_reg)
                if sug_reg and tem_contato:
                    padrao = quando is None and por_regiao
                else:
                    padrao = quando is None and lembrado is not None and c["nome"] in lembrado
                marcas[c["nome"]] = tk.BooleanVar(value=padrao)

            # filtro (lista suspensa com caixinhas, igual ao filtro do Excel) — em vez de uma coluna por contato
            mb = tk.Menubutton(corpo, text="", bg="#2a3f2c", fg=TEXTO, activebackground="#2a3f2c",
                               activeforeground=TEXTO, relief="flat", font=(FONTE, 9, "bold"), padx=10, pady=4,
                               anchor="w", width=30, cursor="hand2", indicatoron=True)
            menu = tk.Menu(mb, tearoff=0, bg=CARTAO, fg=TEXTO, activebackground="#2a3f2c", activeforeground=TEXTO,
                           selectcolor=VERDE, font=(FONTE, 10))
            for c in sorted(contatos, key=lambda c: not any(_mesma_empresa(c.get("empresa", ""), e) for e in sug_reg)):
                ok_reg = any(_mesma_empresa(c.get("empresa", ""), e) for e in sug_reg)
                menu.add_checkbutton(label=("★ " if ok_reg else "   ") + f"{c.get('empresa') or c['nome']}  ({c['nome']})",
                                     variable=marcas[c["nome"]])
            mb.configure(menu=menu)
            mb.grid(row=lin, column=2, sticky="w", padx=10, pady=2)

            def mostrar(*_, b=mb, m=marcas):
                """Botão mostra quem vai receber: verde = escolhida; cinza = ninguém."""
                quem = [next((c.get("empresa") or c["nome"] for c in contatos if c["nome"] == n), n)
                        for n, v in m.items() if v.get()]
                if quem:
                    b.configure(text="✔ " + ", ".join(dict.fromkeys(quem)), bg=VERDE, fg=FUNDO,
                                activebackground=VERDE, activeforeground=FUNDO)
                else:
                    b.configure(text="▼ escolher transportadora", bg="#2a3f2c", fg=FRACO,
                                activebackground="#2a3f2c", activeforeground=TEXTO)
            for var in marcas.values():
                var.trace_add("write", mostrar)
            mostrar()
            linhas.append((titulo, uf, marcas))
            lin += 1

        if any(memoria.get(uf_do_titulo(t)) is None for t in enviar):
            tk.Label(corpo, text="UF em amarelo = estado novo, ainda sem transportadora definida. Escolha no ▼ quem atende (★ = atende a região).",
                     bg=CARTAO, fg=AMBAR, font=(FONTE, 9), anchor="w").grid(
                row=lin, column=0, columnspan=3, sticky="w", padx=10, pady=(8, 0))
            lin += 1
        if bloqueados:
            tk.Label(corpo, text="Não serão enviados:", bg=CARTAO, fg=VERMELHO,
                     font=(FONTE, 9, "bold"), anchor="w").grid(row=lin, column=0, columnspan=4, sticky="w", padx=10, pady=(10, 0))
            lin += 1
            for t, motivo in bloqueados:
                tk.Label(corpo, text=f"✖ {t} — {motivo}", bg=CARTAO, fg=VERMELHO, anchor="w",
                         font=(FONTE, 9)).grid(row=lin, column=0, columnspan=3, sticky="w", padx=10)
                lin += 1

        bot = tk.Frame(w, bg=FUNDO)
        bot.pack(fill="x", padx=16, pady=(0, 14))
        com_dados = tk.BooleanVar(value=False)
        tk.Checkbutton(bot, text="Mandar com os dados (paletes, agenda, coleta, peso, valor)", variable=com_dados,
                       bg=FUNDO, fg=VERDE, selectcolor=CARTAO, activebackground=FUNDO,
                       highlightthickness=0, font=(FONTE, 9, "bold")).pack(side="left", padx=(0, 10))
        lembrar = tk.BooleanVar(value=True)
        tk.Checkbutton(bot, text="Lembrar estes estados para as próximas vezes", variable=lembrar,
                       bg=FUNDO, fg=TEXTO, selectcolor=CARTAO, activebackground=FUNDO,
                       highlightthickness=0, font=(FONTE, 9)).pack(side="left")

        def enviar_agora():
            por_contato = {c["nome"]: [] for c in contatos}
            mapa_uf = {}
            for titulo, uf, marcas in linhas:
                quem = [n for n, v in marcas.items() if v.get()]
                for n in quem:
                    por_contato[n].append(titulo)
                if quem and uf != "??":
                    mapa_uf[uf] = quem
            pacotes = [(c, por_contato[c["nome"]]) for c in contatos if por_contato[c["nome"]]]
            if not pacotes:
                self.escrever("Nenhuma carga marcada para ninguém — envio cancelado.")
                w.destroy()
                return
            # confirmação final: mostra exatamente o que vai para CADA transportadora
            # (as caixas vêm marcadas pela memória do estado — aqui dá pra ver se alguma foi junto sem querer)
            from tkinter import messagebox
            resumo = []
            for c, ts in pacotes:
                resumo.append(f"➜ {c['nome']} ({c.get('empresa', '')}) — {len(ts)} carga(s):")
                resumo += [f"     • {t}" for t in ts]
            if not messagebox.askyesno("Confirmar envio no WhatsApp",
                                       "Vou enviar AGORA no WhatsApp:\n\n" + "\n".join(resumo) +
                                       "\n\nEstá certo? (Não = volta para a tela e você desmarca o que não quer)",
                                       parent=w):
                self.escrever("Envio NÃO confirmado — nada foi enviado. Ajuste as marcações e clique de novo.")
                return
            if lembrar.get() and mapa_uf:
                gravar_estados(mapa_uf)
            for c, ts in pacotes:
                self.escrever(f"Enviando {len(ts)} carga(s) para {c['nome']}…")
            self.fila_whats.put(("enviar", [(c, ts, montar_mensagem(ts, cfg, com_dados.get())) for c, ts in pacotes]))
            canvas.unbind_all("<MouseWheel>")
            w.destroy()

        def cancelar():
            self.escrever("Envio cancelado por você.")
            canvas.unbind_all("<MouseWheel>")
            w.destroy()

        tk.Button(bot, text="🚚 Contratar (sem WhatsApp)", command=lambda: self.janela_contratar([t for t, _, _ in linhas]),
                  bg="#2a3f2c", fg=TEXTO, font=(FONTE, 10), relief="flat", padx=12, pady=6).pack(side="right", padx=(8, 0))
        tk.Button(bot, text="📋 Dados", command=lambda: self.janela_dados([t for t, _, _ in linhas]),
                  bg="#2a3f2c", fg=TEXTO, font=(FONTE, 10), relief="flat", padx=12, pady=6).pack(side="right", padx=(8, 0))
        tk.Button(bot, text="ENVIAR NO WHATSAPP", command=enviar_agora, bg=VERDE, fg=FUNDO,
                  font=(FONTE, 11, "bold"), relief="flat", padx=16, pady=6).pack(side="right")
        tk.Button(bot, text="Agora não", command=cancelar, bg="#2a3f2c", fg=TEXTO,
                  font=(FONTE, 10), relief="flat", padx=12, pady=6).pack(side="right", padx=8)

    # ---- WhatsApp (tudo numa thread só, que mantém o Chrome aberto)
    def _trabalhador_whats(self):
        from whats_envio import WhatsApp
        wa = WhatsApp(log=self.log_thread)
        vistos = TR.ler_vistos()
        primeira_vez = not vistos
        ultima_completa = 0
        while True:
            try:
                acao, dado = self.fila_whats.get(timeout=30)
            except queue.Empty:
                acao, dado = "vigiar", None
            try:
                if acao == "abrir":
                    ok = wa.abrir()
                    self.log_thread("WhatsApp Web conectado." if ok else
                                    "WhatsApp Web ainda não está logado — leia o QR Code.")
                elif acao == "listar":  # só os NOMES das conversas (não lê mensagens)
                    if wa.abrir(espera_login=60):
                        nomes = wa.listar_conversas()
                        self.log_thread(f"Li {len(nomes)} nome(s) de conversa no WhatsApp.")
                        self.fila_ui.put(("conversas", nomes))
                    else:
                        self.log_thread("WhatsApp Web não está logado — leia o QR Code e clique de novo.")
                elif acao == "enviar":
                    resultados = []
                    for c, titulos, texto in dado:
                        ok, det = wa.enviar(c["nome"], texto, c["telefone"])
                        self.log_thread(("✔ " if ok else "✖ ") + f"{c['nome']}: {det}")
                        registrar("ENVIADO" if ok else "FALHOU", c["nome"],
                                  det + " | " + " / ".join(titulos))
                        registrar_cargas(c["nome"], c.get("empresa", ""), titulos, ok, det)
                        resultados.append((c["nome"], c.get("empresa", ""), titulos, ok, det))
                        if ok:
                            gravar_enviados(titulos)
                        else:
                            self.log_thread(f"   As cargas de {c['nome']} NÃO foram marcadas como enviadas; "
                                            "use 'Processar base agora' para tentar de novo.")
                    self.fila_ui.put(("comprovante", resultados))
                elif acao == "ok":  # mandar o OK pra transportadora
                    contato, texto, ids = dado
                    c = next((x for x in ler_config()["contatos"] if x["nome"] == contato), {"telefone": ""})
                    ok, det = wa.enviar(contato, texto, c.get("telefone", ""))
                    self.log_thread(("✔ OK enviado — " if ok else "✖ OK NÃO enviado — ") + f"{contato}: {det}")
                    registrar("OK_ENVIADO" if ok else "OK_FALHOU", contato, texto)
                    lista = TR.ler_aceites()
                    for a in lista:
                        if a["id"] in ids:
                            a["status"] = "ok" if ok else "aguardando"
                            if ok:
                                a["ok_em"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                    TR.gravar_aceites(lista)
                elif acao == "vigiar":
                    cfg = ler_config()
                    self._fila_extra(wa, cfg)  # envios pedidos pelo robô "Agendado com NF"
                    if not cfg["vigiar_respostas"] or not wa.aberto():
                        continue
                    # só quem tem carga aberta (enviada nos últimos dias)
                    abertos = {c["nome"]: c for c in cfg["contatos"] if TR.cargas_abertas(c["nome"])}
                    if not abertos:
                        continue
                    completa = time.time() - ultima_completa > 180  # a cada 3 min lê todas
                    alvo = list(abertos) if completa else wa.com_nao_lidas(list(abertos))
                    if completa:
                        ultima_completa = time.time()
                    for nome in alvo:
                        self._ler_respostas(wa, abertos[nome], vistos, primeira_vez)
                    primeira_vez = False
                    TR.gravar_vistos(vistos)
            except Exception as e:
                self.log_thread(f"ERRO no WhatsApp: {e.__class__.__name__}: {str(e)[:150]}")

    def _fila_extra(self, wa, cfg):
        """Envia o que o robô 'Agendado com NF' pediu. Erro de um item não derruba os outros
        e fica registrado no log (o Agendado com NF mostra o resultado)."""
        fila_arq = BASE / "fila_contratacao_extra.json"
        if not fila_arq.exists():
            return
        try:
            fila = json.loads(fila_arq.read_text(encoding="utf-8"))
        except Exception as e:
            self.log_thread(f"✖ Fila do Agendado com NF ilegível: {e}")
            fila = []
        try:
            fila_arq.unlink()
        except OSError:
            pass
        for item in fila:
            nome = item.get("contato", "")
            titulos = item.get("titulos", [])
            try:
                c = next((x for x in cfg["contatos"] if x["nome"] == nome), {"telefone": ""})
                tel = c.get("telefone", "") or item.get("telefone", "")  # a Agenda manda o telefone do contato
                if titulos:
                    self.log_thread(f"{item.get('origem', 'Agendado com NF')}: enviando {len(titulos)} carga(s) para {nome}…")
                else:
                    self.log_thread(f"{item.get('origem', 'Agenda')}: enviando mensagem para {nome}…")
                ok, det = wa.enviar(nome, item.get("texto", ""), tel)
            except Exception as e:
                ok, det = False, f"erro no WhatsApp: {e.__class__.__name__}: {str(e)[:120]}"
            self.log_thread(("✔ " if ok else "✖ ") + f"{item.get('origem', 'Envio')} → {nome}: {det}")
            registrar("ENVIADO" if ok else "FALHOU", nome, det + " | " + " / ".join(titulos))
            if ok:
                gravar_enviados(titulos)

    def _ler_respostas(self, wa, contato, vistos, primeira_vez):
        nome = contato["nome"]
        msgs = wa.mensagens_recebidas(nome, contato.get("telefone", ""))
        if msgs is None:
            return
        novos = []
        limite = datetime.now() - timedelta(hours=3)
        for m in msgs:
            if m["id"] in vistos:
                continue
            vistos.add(m["id"])
            quando = TR.hora_da_mensagem(m["pre"])
            if primeira_vez and (quando is None or quando < limite):
                continue  # na 1ª vez só olha as das últimas 3 horas
            if m["texto"].strip():
                novos.append(m["texto"].strip())
        if not novos:
            return
        texto = "\n".join(novos)
        self.log_thread(f"{nome} respondeu: {texto[:160]}")
        if not TR.eh_aceite(texto):
            return
        ja = TR.titulos_ja_aceitos()
        abertas = [t for t in TR.cargas_abertas(nome) if t not in ja]
        cargas = TR.casar_cargas(texto, abertas)
        self.fila_ui.put(("aceite", (nome, texto, cargas, abertas)))

    # ---- janela: transportadora disse que COLETA
    def janela_aceite(self, nome, texto, cargas, abertas):
        try:
            import winsound
            winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
        except Exception:
            self.bell()
        w = tk.Toplevel(self)
        w.title(f"{nome} vai coletar")
        w.configure(bg=FUNDO)
        w.attributes("-topmost", True)
        w.geometry("760x560")
        bot = tk.Frame(w, bg=FUNDO)
        bot.pack(side="bottom", fill="x", padx=16, pady=(0, 14))
        tk.Label(w, text=f"🚚 {nome} disse que COLETA", bg=FUNDO, fg=VERDE,
                 font=(FONTE, 15, "bold")).pack(anchor="w", padx=16, pady=(12, 2))
        q = tk.Text(w, height=4, bg=CARTAO, fg=TEXTO, font=(FONTE, 10), relief="flat", wrap="word")
        q.insert("1.0", texto)
        q.configure(state="disabled")
        q.pack(fill="x", padx=16, pady=(2, 8))
        tk.Label(w, text=("Marque qual(is) carga(s) ela pegou:" if cargas else
                          "⚠ Não consegui identificar a carga pela mensagem — marque qual(is) é(são):"),
                 bg=FUNDO, fg=(TEXTO if cargas else AMBAR), font=(FONTE, 10, "bold")).pack(anchor="w", padx=16)
        quadro = tk.Frame(w, bg=CARTAO)
        quadro.pack(fill="both", expand=True, padx=16, pady=6)
        marcas = []
        for t in abertas:
            v = tk.BooleanVar(value=t in cargas)
            tk.Checkbutton(quadro, text=t, variable=v, bg=CARTAO, fg=TEXTO, selectcolor=FUNDO,
                           activebackground=CARTAO, activeforeground=TEXTO, anchor="w",
                           font=(FONTE, 10), highlightthickness=0).pack(fill="x", padx=8, pady=1)
            marcas.append((t, v))
        if not abertas:
            tk.Label(quadro, text="(nenhuma carga aberta enviada pra ela nos últimos 10 dias)", bg=CARTAO,
                     fg=FRACO).pack(anchor="w", padx=8, pady=6)
        tk.Label(w, text="Agora: forme a carga (O.C.) no Sankhya pra ela. Depois volte aqui (📦 Coletas aceitas) "
                         "e clique em 'Carga formada → mandar OK'. Eu te lembro até você marcar.",
                 bg=FUNDO, fg=AMBAR, font=(FONTE, 9, "bold"), wraplength=720, justify="left").pack(anchor="w", padx=16)

        def guardar():
            escolhidas = [t for t, v in marcas if v.get()] or [""]
            lista = TR.ler_aceites()
            agora = datetime.now()
            for t in escolhidas:
                lista.append({"id": f"{agora:%Y%m%d%H%M%S}-{len(lista)}", "contato": nome, "titulo": t,
                              "mensagem": texto, "recebido": agora.strftime("%d/%m/%Y %H:%M"),
                              "status": "aguardando"})
            TR.gravar_aceites(lista)
            self._ultimo_lembrete = time.time()  # primeiro lembrete só daqui a X min
            registrar("ACEITE", nome, " / ".join(t for t in escolhidas if t) or texto[:120])
            self.escrever(f"📦 {nome} vai coletar: " + (", ".join(TR.resumo(t) for t in escolhidas if t) or "(carga não identificada)")
                          + " — falta formar a carga e mandar o OK.")
            w.destroy()

        tk.Button(bot, text="Não é coleta", command=w.destroy, bg="#2a3f2c", fg=TEXTO, relief="flat",
                  padx=12, pady=6).pack(side="right", padx=(6, 0))
        tk.Button(bot, text="GUARDAR — vou formar a carga", command=guardar, bg=VERDE, fg=FUNDO, relief="flat",
                  font=(FONTE, 10, "bold"), padx=14, pady=6).pack(side="right")

    # ---- lista de coletas aceitas / mandar OK
    def janela_coletas(self, lembrete=False):
        lista = TR.ler_aceites()
        pend = [a for a in lista if a["status"] == "aguardando"]
        feitos = [a for a in lista if a["status"] == "ok"][-15:]
        w = tk.Toplevel(self)
        w.title("Coletas aceitas")
        w.configure(bg=FUNDO)
        w.attributes("-topmost", True)
        w.geometry("860x560")
        if lembrete:
            self.janela_lembrete = w
            w.protocol("WM_DELETE_WINDOW", lambda: (setattr(self, "janela_lembrete", None), w.destroy()))
        bot = tk.Frame(w, bg=FUNDO)
        bot.pack(side="bottom", fill="x", padx=16, pady=(0, 14))
        titulo = (f"⏰ Lembrete: falta mandar o OK de {len(pend)} coleta(s)" if lembrete
                  else f"📦 Coletas aceitas — {len(pend)} aguardando o seu OK")
        tk.Label(w, text=titulo, bg=FUNDO, fg=(AMBAR if pend else VERDE),
                 font=(FONTE, 14, "bold")).pack(anchor="w", padx=16, pady=(12, 2))
        tk.Label(w, text="Depois de FORMAR A CARGA no Sankhya, marque e clique em 'Carga formada → mandar OK'.",
                 bg=FUNDO, fg=FRACO, font=(FONTE, 9)).pack(anchor="w", padx=16)
        quadro = tk.Frame(w, bg=CARTAO)
        quadro.pack(fill="both", expand=True, padx=16, pady=8)
        marcas = []
        for a in pend:
            v = tk.BooleanVar(value=False)
            txt = f"{a['recebido'][:16]}  ·  {a['contato']}  ·  {a['titulo'] or '(carga não identificada: ' + a['mensagem'][:60] + ')'}"
            tk.Checkbutton(quadro, text=txt, variable=v, bg=CARTAO, fg=TEXTO, selectcolor=FUNDO,
                           activebackground=CARTAO, activeforeground=TEXTO, anchor="w", justify="left",
                           font=(FONTE, 10), highlightthickness=0, wraplength=790).pack(fill="x", padx=8, pady=1)
            marcas.append((a, v))
        if not pend:
            tk.Label(quadro, text="Nada pendente. ✔", bg=CARTAO, fg=VERDE, font=(FONTE, 11)).pack(anchor="w", padx=8, pady=8)
        if feitos:
            tk.Label(quadro, text="\nÚltimos OKs:", bg=CARTAO, fg=FRACO, font=(FONTE, 9, "bold")).pack(anchor="w", padx=8)
            for a in reversed(feitos):
                tk.Label(quadro, text=f"✔ {a.get('ok_em', '')}  {a['contato']}  {TR.resumo(a['titulo']) if a['titulo'] else ''}",
                         bg=CARTAO, fg=FRACO, font=(FONTE, 9), anchor="w").pack(fill="x", padx=8)

        def selecionados():
            return [a for a, v in marcas if v.get()]

        def mandar_ok():
            sel = selecionados()
            if not sel:
                return
            por_contato = {}
            for a in sel:
                por_contato.setdefault(a["contato"], []).append(a)
            cfg = ler_config()
            for contato, itens in por_contato.items():
                cargas = ", ".join(TR.resumo(a["titulo"]) for a in itens if a["titulo"]) or "conforme combinado"
                self._confirmar_ok(contato, cfg["msg_ok"].format(cargas=cargas), [a["id"] for a in itens])
            w.destroy()
            if lembrete:
                self.janela_lembrete = None

        def so_marcar():
            sel = {a["id"] for a in selecionados()}
            if not sel:
                return
            lista2 = TR.ler_aceites()
            for a in lista2:
                if a["id"] in sel:
                    a["status"] = "ok"
                    a["ok_em"] = datetime.now().strftime("%d/%m/%Y %H:%M")
            TR.gravar_aceites(lista2)
            self.escrever(f"Marquei {len(sel)} coleta(s) como OK (sem mandar mensagem).")
            w.destroy()
            if lembrete:
                self.janela_lembrete = None

        def descartar():
            sel = {a["id"] for a in selecionados()}
            if not sel:
                return
            lista2 = [a for a in TR.ler_aceites() if a["id"] not in sel]
            TR.gravar_aceites(lista2)
            w.destroy()
            if lembrete:
                self.janela_lembrete = None

        tk.Button(bot, text="Fechar", command=lambda: (setattr(self, "janela_lembrete", None), w.destroy()),
                  bg="#2a3f2c", fg=TEXTO, relief="flat", padx=12, pady=6).pack(side="right", padx=(6, 0))
        tk.Button(bot, text="Apagar (não era coleta)", command=descartar, bg="#2a3f2c", fg=TEXTO, relief="flat",
                  padx=12, pady=6).pack(side="right", padx=(6, 0))
        tk.Button(bot, text="Só marcar OK (sem mensagem)", command=so_marcar, bg="#2a3f2c", fg=TEXTO,
                  relief="flat", padx=12, pady=6).pack(side="right", padx=(6, 0))
        tk.Button(bot, text="Carga formada → mandar OK", command=mandar_ok, bg=VERDE, fg=FUNDO, relief="flat",
                  font=(FONTE, 10, "bold"), padx=14, pady=6).pack(side="right")

    def _confirmar_ok(self, contato, texto, ids):
        """Mostra o OK antes de mandar (pode editar)."""
        w = tk.Toplevel(self)
        w.title(f"OK para {contato}")
        w.configure(bg=FUNDO)
        w.attributes("-topmost", True)
        w.geometry("600x260")
        tk.Label(w, text=f"Mensagem de OK para {contato} (pode editar):", bg=FUNDO, fg=TEXTO,
                 font=(FONTE, 11, "bold")).pack(anchor="w", padx=14, pady=(12, 4))
        a = tk.Text(w, height=5, bg=CARTAO, fg=TEXTO, font=(FONTE, 10), relief="flat", wrap="word",
                    insertbackground=TEXTO)
        a.insert("1.0", texto)
        a.pack(fill="both", expand=True, padx=14)
        bot = tk.Frame(w, bg=FUNDO)
        bot.pack(fill="x", padx=14, pady=10)

        def enviar():
            t = a.get("1.0", "end").strip()
            if t:
                self.fila_whats.put(("ok", (contato, t, ids)))
                self.escrever(f"Enviando OK para {contato}…")
                w.destroy()
        tk.Button(bot, text="Cancelar", command=w.destroy, bg="#2a3f2c", fg=TEXTO, relief="flat",
                  padx=12, pady=6).pack(side="right", padx=(6, 0))
        tk.Button(bot, text=f"ENVIAR OK PARA {contato.upper()}", command=enviar, bg=VERDE, fg=FUNDO, relief="flat",
                  font=(FONTE, 10, "bold"), padx=14, pady=6).pack(side="right")

    def _lembrete_ok(self):
        """De X em X minutos, se tiver coleta aceita sem OK, abre o lembrete."""
        try:
            cfg = ler_config()
            pend = [a for a in TR.ler_aceites() if a["status"] == "aguardando"]
            agora = time.time()
            ultimo = getattr(self, "_ultimo_lembrete", 0)
            if pend and self.janela_lembrete is None and agora - ultimo >= cfg["lembrar_min"] * 60:
                self._ultimo_lembrete = agora
                self.janela_coletas(lembrete=True)
        except Exception as e:
            self.escrever(f"Lembrete: {e}")
        self.after(60000, self._lembrete_ok)

    def fechar(self):
        self.destroy()


if __name__ == "__main__":
    Robo().mainloop()
