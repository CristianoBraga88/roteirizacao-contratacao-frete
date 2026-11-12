"""
Gera o HISTÓRICO DE CARGAS ENVIADAS para cada transportadora
(a partir do log do Robô de Contratação).

Saída: Historico_Contratacao.xlsx (nesta pasta)
  - aba "Todas"       -> tudo, da mais nova para a mais antiga
  - uma aba por transportadora (Contato Alfa, Contato Beta, ...)
  - aba "Buscar"      -> dica de como procurar (Ctrl+F) por cliente, cidade, ticket

USO: dois cliques em HISTORICO_CONTRATACAO.bat (ou botão no Robô de Contratação).
"""
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
LOG = BASE / "log_contratacao.csv"
ROTEIROS = BASE / "roteiros_contratacao.json"
ACEITES = BASE / "coletas_aceitas.json"
SAIDA = BASE / "Historico_Contratacao.xlsx"

RE_TIT = re.compile(r"^(?:(AL) X - )?([A-Z]{2}) - (.+?) - ([\d.]+) kg - ([A-Z ]+?)\s*\*(.*)\*\s*$")


def _ler_json(p, padrao):
    try:
        return json.loads(Path(p).read_text(encoding="utf-8"))
    except Exception:
        return padrao


RE_TIT_AG = re.compile(r"^(?:(AL) X - )?([A-Z]{2}) - (.+?) - ([\d.,]+) (?:kg|TON)(?: - (.*))?$")


def partes(titulo):
    m = RE_TIT.match(titulo.strip())
    if not m:
        m2 = RE_TIT_AG.match(titulo.strip())  # formato do robô "Agendado com NF"
        if m2:
            origem, uf, cidade, peso, resto = m2.groups()
            resto = (resto or "").strip()
            veic = ""
            mv = re.match(r"^(CARRETA|TRUCK|TOCO)\s*-\s*(.*)$", resto)
            if mv:
                veic, resto = mv.groups()
            if "TON" in titulo:
                peso = peso + " t"
            return {"origem": "MACEIÓ (AL)" if origem else "UBERABA", "uf": uf, "cidade": cidade.strip(),
                    "peso": peso, "veiculo": veic, "agenda": resto}
        return {"origem": "", "uf": "", "cidade": titulo, "peso": "", "veiculo": "", "agenda": ""}
    origem, uf, cidade, peso, veic, agenda = m.groups()
    return {"origem": "MACEIÓ (AL)" if origem else "UBERABA", "uf": uf, "cidade": cidade.strip(),
            "peso": peso, "veiculo": veic.strip(), "agenda": agenda.strip()}


def ler_envios():
    linhas = []
    try:
        brutas = LOG.read_text(encoding="utf-8-sig").splitlines()[1:]
    except OSError:
        return []
    roteiros = _ler_json(ROTEIROS, {})
    aceites = {}
    for a in _ler_json(ACEITES, []):
        if a.get("titulo"):
            aceites[(a["contato"], a["titulo"])] = a
    for ln in brutas:
        p = ln.split(";", 3)
        if len(p) < 4 or p[1] != "ENVIADO" or " | " not in p[3]:
            continue
        try:
            quando = datetime.strptime(p[0], "%d/%m/%Y %H:%M:%S")
        except ValueError:
            continue
        for t in p[3].split(" | ", 1)[1].split(" / "):
            t = t.strip()
            if not t:
                continue
            x = partes(t)
            a = aceites.get((p[2], t))
            linhas.append({
                "quando": quando, "transportadora": p[2], "titulo": t, **x,
                "cliente": roteiros.get(t, ""),
                "coleta": ("OK enviado" if a and a.get("status") == "ok" else
                           "aceitou (falta OK)" if a else ""),
            })
    linhas.sort(key=lambda r: r["quando"], reverse=True)
    return linhas


def gerar():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    envios = ler_envios()
    wb = Workbook()
    cab = ["Enviado em", "Transportadora", "Origem", "UF", "Cidade", "Peso (kg)", "Veículo",
           "Agendamento", "Cliente / Roteiro", "Coleta", "Título enviado"]
    larg = [17, 18, 12, 5, 26, 10, 10, 26, 60, 18, 70]
    verde = PatternFill("solid", fgColor="1F5E3B")

    def aba(ws, linhas, titulo):
        ws.append([titulo])
        ws["A1"].font = Font(bold=True, size=13)
        ws.append([f"Gerado em {datetime.now():%d/%m/%Y %H:%M}  ·  {len(linhas)} carga(s) enviada(s)  ·  "
                   "Ctrl+F para procurar cliente, cidade ou ticket"])
        ws.append(cab)
        for c in range(1, len(cab) + 1):
            cel = ws.cell(3, c)
            cel.font = Font(bold=True, color="FFFFFF")
            cel.fill = verde
        for r in linhas:
            ws.append([r["quando"], r["transportadora"], r["origem"], r["uf"], r["cidade"],
                       r["peso"], r["veiculo"], r["agenda"], r["cliente"], r["coleta"], r["titulo"]])
            ws.cell(ws.max_row, 1).number_format = "dd/mm/yyyy hh:mm"
        for i, w in enumerate(larg, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A4"
        if linhas:
            ws.auto_filter.ref = f"A3:{get_column_letter(len(cab))}{ws.max_row}"
        for row in ws.iter_rows(min_row=4):
            for cel in row:
                cel.alignment = Alignment(vertical="top")

    ws = wb.active
    ws.title = "Todas"
    aba(ws, envios, "HISTÓRICO DE CARGAS ENVIADAS PARA CONTRATAÇÃO — todas as transportadoras")
    for transp in sorted({r["transportadora"] for r in envios}):
        nome = re.sub(r"[\\/*?:\[\]]", "", transp)[:31]
        aba(wb.create_sheet(nome), [r for r in envios if r["transportadora"] == transp],
            f"HISTÓRICO — {transp}")
    # resumo por transportadora
    rs = wb.create_sheet("Resumo", 1)
    rs.append(["Transportadora", "Cargas enviadas", "Primeiro envio", "Último envio"])
    for c in range(1, 5):
        rs.cell(1, c).font = Font(bold=True, color="FFFFFF")
        rs.cell(1, c).fill = verde
    for transp in sorted({r["transportadora"] for r in envios}):
        lst = [r for r in envios if r["transportadora"] == transp]
        rs.append([transp, len(lst), min(r["quando"] for r in lst), max(r["quando"] for r in lst)])
        rs.cell(rs.max_row, 3).number_format = rs.cell(rs.max_row, 4).number_format = "dd/mm/yyyy hh:mm"
    for col, w in zip("ABCD", (22, 16, 18, 18)):
        rs.column_dimensions[col].width = w
    wb.save(SAIDA)
    return SAIDA, len(envios)


if __name__ == "__main__":
    try:
        arq, n = gerar()
        print(f"Pronto: {arq.name} com {n} carga(s).")
        if "--nao-abrir" not in sys.argv:
            os.startfile(str(arq))
    except PermissionError:
        print("O Historico_Contratacao.xlsx está aberto no Excel. Feche e rode de novo.")
        input("Enter para sair")
