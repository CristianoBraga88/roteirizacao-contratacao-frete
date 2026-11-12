"""
Gera o relatório de Roteirização: Lotação (com títulos pra copiar, prontos
pra Contratação e pra Roteirização) e Fracionado por operador logístico.

Lê:   a base de Roteirização MAIS RECENTE que estiver na mesma pasta do
      script (qualquer .xls/.xlsx com "roteirizacao" no nome — tanto faz se
      vier "Base Roteirizacao.xls", "Base_Roteirizacao.xls" ou
      "Base Roteirizacao (1).xls")
Gera: relatorio_roteirizacao.html (mesma pasta do script)

Como usar:
- Exporte a base do Sankhya e salve na pasta do script.
- Rode o script. No terminal ele mostra QUAL arquivo leu e quando ele foi
  salvo. O mesmo aparece no topo do relatório, junto com o "gerado em".
- Toda vez que rodar, o script abre no navegador o HTML que ele ACABOU de
  gerar (o da pasta do script). Assim você nunca fica olhando uma cópia
  antiga do relatório que esteja em outra pasta.
- Pode fechar as abas antigas do relatório — cada execução abre uma nova.
- Se o "gerado em" do topo do relatório não for o horário de agora, o
  script deu erro: olhe o terminal do VS Code.

Regras aplicadas:
- Universo: O.C. PAI vazio, Ordem de carga vazia, Cód. Parceiro Orig. preenchido
- Lotação x Fracionado: peso total do parceiro (só linhas sem agendamento)
  < 7.000kg e sem agendamento -> Fracionado; senão -> Lotação (igual ao VBA
  RoteirizarCargas)
- Lotação: agrupa por Cód. Parceiro + Tipo de Entrega + Cidade/Estado +
  Tipo de Carga + Data Agendamento + Observação, e distribui em veículos:
    - qualquer LINHA INDIVIDUAL com peso >= 27.000kg vira um RODOTREM
      sozinho (nunca por soma de linhas menores)
    - o restante segue TRUCK (até 13.000kg) / CARRETA (até 23.000kg)
  Cada veículo fechado gera um título "Para Contratação" e um "Para
  Roteirização", conforme o Tipo de Entrega (Agendado, Agendado com NF,
  Ordem de Chegada)
- PRAZO DE SAÍDA DE UBERABA (cargas de Lotação com data de agendamento):
    km estimado Uberaba -> cidade (linha reta x FATOR_ESTRADA, ou o km
    que você puser em distancias_uberaba.ini) / KM_POR_DIA = dias de
    viagem. Regra do CD (e-mail do Gerente, 03/11/2026), igual p/ Uberaba e Maceió:
      D-0 liberação da carga -> D+1 preparação e carregamento -> D+2 trânsito (500 km/dia)
    "Coletar até" = agenda - dias de viagem (início do trânsito).
    "Carregar até" = 1 dia útil antes da coleta (preparação/carregamento).
    "Liberar até" = 1 dia útil antes do carregamento (liberação da O.C.).
    Capacidade: Uberaba 23 O.C.s/dia, Maceió 6 O.C.s/dia — o relatório avisa
    quando um dia de carregamento passa do limite.
- Fracionado por operador: replica as regras que o time já usa no Sankhya
  pra puxar a lista de cada operador (Nu, Theta, Kappa,
  Epsilon, Iota) — filtros por UF/cidade de destino, tipo de operação,
  parceiro e perfil de parceiro
"""

import importlib
import importlib.metadata
import base64
import glob
import html as html_lib
import json
import math
import os
import re
import subprocess
import sys
import unicodedata
import webbrowser
from datetime import datetime, timedelta
from pathlib import Path

# ----------------------------------------------------------------------
# ATUALIZADOR DE DEPENDÊNCIAS
# ----------------------------------------------------------------------
DEPENDENCIAS = {
    "pandas": "2.0.0",
    "openpyxl": "3.1.0",
    "xlrd": "2.0.1",
}


def _versao_instalada(pacote: str):
    try:
        return importlib.metadata.version(pacote)
    except importlib.metadata.PackageNotFoundError:
        return None


def _versao_menor(atual: str, minima: str) -> bool:
    def normaliza(v):
        return [int(p) for p in v.split(".") if p.isdigit()]
    return normaliza(atual) < normaliza(minima)


def garantir_dependencias():
    pendentes = []
    for pacote, versao_minima in DEPENDENCIAS.items():
        instalada = _versao_instalada(pacote)
        if instalada is None:
            pendentes.append((pacote, f"não instalado -> {versao_minima}"))
        elif _versao_menor(instalada, versao_minima):
            pendentes.append((pacote, f"{instalada} -> {versao_minima}"))

    if not pendentes:
        return

    print("Atualizando dependências:")
    for pacote, resumo in pendentes:
        print(f"  - {pacote}: {resumo}")

    pacotes_para_instalar = [f"{pacote}>={DEPENDENCIAS[pacote]}" for pacote, _ in pendentes]
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", *pacotes_para_instalar])
    except subprocess.CalledProcessError as e:
        comando = " ".join(pacotes_para_instalar)
        print(
            "Não consegui instalar/atualizar automaticamente. Rode manualmente:\n"
            f"    {sys.executable} -m pip install --upgrade {comando}"
        )
        raise SystemExit(1) from e

    importlib.invalidate_caches()
    print("Dependências atualizadas. Prosseguindo...\n")


garantir_dependencias()

import pandas as pd  # noqa: E402


# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------
PASTA_SAIDA = os.path.dirname(os.path.abspath(__file__))

# a base é localizada automaticamente: o script procura, na MESMA pasta
# onde está, todos os arquivos .xls/.xlsx que tenham "roteirizacao" no
# nome e usa o que foi salvo por último. Assim não importa se o arquivo
# veio com espaço, underline ou "(1)" no nome.
TRECHO_NOME_BASE = "roteirizacao"
EXTENSOES_BASE = ("*.xls", "*.xlsx")

# se a base mais recente tiver sido salva há mais que isso, o script
# avisa no terminal (não bloqueia — é só um alerta de "base velha")
ALERTA_BASE_ANTIGA_HORAS = 12

ARQUIVO_SAIDA = os.path.join(PASTA_SAIDA, "relatorio_roteirizacao.html")
CAMINHO_LOGO = os.path.join(PASTA_SAIDA, "logo.png")

# se True, abre no navegador o relatório recém-gerado TODA vez que o
# script rodar (garante que você está vendo o arquivo certo, e não uma
# cópia antiga). Coloque False pra nunca abrir automaticamente.
ABRIR_NO_NAVEGADOR = True

# cabeçalho real da base de Roteirização começa na linha 2 (contando a
# partir de 0) — confirmado com a base de exemplo real
LINHA_CABECALHO = 2

COL_OC_PAI = "O.C. PAI"
COL_ORDEM_CARGA = "Ordem de carga"
COL_TIPO_OPERACAO = "Tipo Operação"
COL_COD_PARCEIRO = "Cód. Parceiro Orig."
COL_CLIENTE = "Nome Parceiro (Parceiro)"
COL_COD_CIDADE_PARCEIRO = "Cód. Cidade (Parceiro)"
COL_ESTADO = "Estado"
COL_CIDADE_ESTADO = "Cidade/Estado"
COL_TIPO_CARGA = "Tipo de Carga"
COL_PESO = "Peso bruto dos Itens"
COL_DATA_AGENDAMENTO = "Data Agendamento Entrega"
COL_OBSERVACAO = "[Observação]"
COL_VALOR = "Vlr. Nota"
COL_NRO_UNICO = "Nro. Único"
COL_TIPO_ENTREGA = "Tipo de Entrega"
COL_PENDENTE = "Pendente"
COL_EMPRESA = "Empresa"
COL_PERFIL_PARCEIRO = "Perfil Principal Parceiro"
COL_TOT_PALLET_PEDIDO = "Tot_pallet_pedido"

PESO_MAX_FRACIONADO = 7000.0
CAP_TRUCK = 13000.0
CAP_CARRETA = 23000.0
PESO_MIN_RODOTREM = 27000.0  # linha INDIVIDUAL >= isso vira rodotrem sozinha

EMPRESA_UBERABA = 2
EMPRESA_MACEIO = 209

# ---- prazo de saída de Uberaba (só cargas de Lotação agendadas, origem Uberaba)
KM_POR_DIA = 500             # quanto o caminhão roda por dia
DIAS_EXPEDICAO = 1           # dias úteis que a expedição precisa antes da saída
FATOR_ESTRADA = 1.25         # linha reta x 1,25 ≈ km de estrada (ajuste fino no .ini)
CARREGA_SABADO = True        # a carga pode sair de Uberaba no sábado?
CARREGA_DOMINGO = False      # ...e no domingo? (se não, volta pro dia anterior permitido)
DIAS_LIBERACAO = 1           # D-0: liberação da carga, 1 dia útil antes do carregamento (regra do Gerente)
CAPACIDADE_OC_DIA = {"Uberaba": 23, "Maceió": 6}   # O.C.s que o CD carrega por dia
UBERABA_LAT_LON = (-19.7472, -47.9381)
MACEIO_LAT_LON = (-9.6658, -35.7353)
ARQ_COORDENADAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cidades_coordenadas.csv")
ARQ_KM_MANUAL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "distancias_uberaba.ini")

# textos pro botão "?" de cada operador, explicando como funciona a criação
# de carga/coleta/peso — usados só pra exibir, não afetam nenhum cálculo
INFO_OPERADORES = {
    "Theta": {
        "coleta": "Coleta no CD e Entrega no cliente",
        "criacao_carga": "1 O.C. pai e 1 O.C. Filho pra tudo que estiver no filtro",
        "peso": "Sem restrição",
        "data_coleta": "Realiza coleta diária Seg a Sex, só precisa do envio do romaneio",
    },
    "Xi": {
        "coleta": "Coleta no CD e Entrega no cliente",
        "criacao_carga": "1 O.C. pai e 1 O.C. Filho para cada região (ex: PA 1 Filho, RN 1 Filho...)",
        "peso": "Sem restrições, porém recomendado peso de truck",
        "data_coleta": "Realiza coleta diária Seg a Sex, só precisa do envio do romaneio",
    },
    "Nu": {
        "coleta": "Não coleta no CD — necessário contratar para entregar",
        "criacao_carga": "1 O.C. pai e 1 O.C. Filho pra tudo que estiver no filtro",
        "peso": "Acima de 17 toneladas - CARRETA",
        "data_coleta": "Combinar com transportador de LOTAÇÃO",
    },
    "Kappa": {
        "coleta": "Coleta no CD e Entrega no cliente",
        "criacao_carga": "1 O.C. pai e 1 O.C. Filho para cada região (ex: PA 1 Filho, RN 1 Filho...)",
        "peso": "Acima de 17 toneladas - CARRETA",
        "data_coleta": "Realiza coleta diária Seg a Sex, só precisa do envio do romaneio",
    },
    "Epsilon": {
        "coleta": "Coleta no CD e Entrega no cliente",
        "criacao_carga": "1 O.C. pai e 1 O.C. Filho para cada região (ex: PA 1 Filho, RN 1 Filho...)",
        "peso": "Sem restrição",
        "data_coleta": "Realiza coleta diária Seg a Sex, só precisa do envio do romaneio",
    },
    "Iota": {
        "coleta": "Coleta no CD e Entrega no cliente",
        "criacao_carga": "1 O.C. Pai e 1 Filho para cada MACRO REGIÃO/REGIÃO (ex: Macro 2 MG = 1 O.C. Pai e 1 Filho; Macro 906 MG = 1 O.C. Pai e 1 Filho; Macro 906 SP = 1 O.C. Pai e 1 Filho)",
        "peso": "Sem restrição",
        "data_coleta": "Realiza coleta diária Seg a Sex, só precisa do envio do romaneio",
    },
}

# código de UF (CODUFDESTINO do Sankhya) -> sigla, pra traduzir os filtros
UF_POR_CODIGO = {
    1: "SP", 2: "MG", 3: "DF", 4: "GO", 5: "MT", 6: "BA", 7: "RJ", 8: "PR",
    9: "PA", 10: "PE", 11: "RO", 12: "MS", 13: "SC", 14: "TO", 15: "RS",
    16: "ES", 17: "PB", 18: "AM", 19: "AL", 20: "AC", 21: "CE", 22: "SE",
    23: "PI", 24: "RR", 26: "RN", 28: "AP", 31: "MA",
}


def _ufs(*codigos):
    return {UF_POR_CODIGO[c] for c in codigos}


# cidades que eram exclusivas da Omicron — o operador foi removido, mas
# essas cidades continuam fora da Nu (não caem em nenhum operador
# automaticamente, ficam de fora até decidir o que fazer com elas)
CIDADES_OMICRON = {
    403, 558, 628, 888, 891, 1098, 1437, 1552, 1658, 1779, 1851, 1855, 2037,
    2364, 2376, 2401, 2515, 2863, 3026, 3112, 3468, 3885, 4178, 4224, 4408,
    4502, 4537, 4599, 4606, 4858, 5107, 5114, 5439, 5619, 5735,
}

CODTIPOPER_FRACIONADO = {3112, 3113, 3114, 3166, 3168, 3144, 3130}
PARCEIRO_EXCLUIDO = "00000"
# Perfil Principal do Parceiro = Farma / Drogaria / Perfumaria. Pelos nomes
# que vêm na própria base: 8010000 REDE FARMA, 8020000 REDE PERFUMARIA,
# 5030000 PEQUENO PERFUMARIA, 5020000 (pequena farma/drogaria). Configuração LTL Empresa Exemplo de 06/10/2026:
#   THETA: RJ+ES todos os canais + Farma/Drog/Perf em MG, SP, RS, PR, SC
#   IOTA: MG, SP, PR, SC, RS todos os canais EXCETO Farma/Drog/Perf
#   EPSILON: GO, DF, MT, MS, PA, RO, RR, AC, AP (+ TO e AM)
#   KAPPA: 100% Nordeste · NU: sem mudança
PERFIL_ATIVA_ESPECIAL = {"8010000", "8020000", "5030000", "5020000"}


# ----------------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------------
def _para_numero_brl(valor) -> float:
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace("R$", "").strip()
    if not texto:
        return 0.0
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0


def _para_inteiro_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    if isinstance(valor, float):
        return str(int(valor))
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto


def _limpa_maiuscula(s) -> str:
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = " ".join(s.split())
    return s.upper()


def _perfil_prefixo(valor) -> str:
    """'6010000 - REDE ALIMENTAR' -> '6010000' (só o código antes do '-')."""
    texto = str(valor).strip()
    if not texto or texto.lower() == "nan":
        return ""
    return texto.split("-")[0].strip()


def fmt_moeda(v: float) -> str:
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_peso(v: float) -> str:
    return f"{v:,.0f}".replace(",", ".") + " kg"


def carregar_logo_base64(caminho: str) -> str:
    if not os.path.isfile(caminho):
        print(f"Aviso: logo não encontrada em {caminho} — gerando sem logo.")
        return ""
    with open(caminho, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return f"data:image/png;base64,{b64}"


def abrir_no_navegador(caminho: str) -> None:
    """Abre o HTML no navegador padrão. No Windows usa os.startfile (mais
    confiável com caminhos com acento/espaço); nos outros sistemas cai no
    módulo webbrowser."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(caminho)  # type: ignore[attr-defined]
        else:
            webbrowser.open(Path(caminho).resolve().as_uri())
        print("Abrindo o relatório no navegador...")
    except Exception as e:  # noqa: BLE001
        print(f"Não consegui abrir o navegador automaticamente ({e}).")
        print(f"Abra manualmente: {caminho}")


def salvar_html(caminho: str, conteudo: str) -> None:
    """Grava num arquivo temporário e depois troca de uma vez pelo final.
    Assim, se você der F5 bem na hora em que o script está gravando, o
    navegador nunca pega um HTML pela metade."""
    temporario = caminho + ".tmp"
    with open(temporario, "w", encoding="utf-8") as f:
        f.write(conteudo)
    os.replace(temporario, caminho)


def localizar_base_mais_recente(pasta: str) -> str:
    """Procura na pasta todos os .xls/.xlsx com 'roteirizacao' no nome
    (ignorando acento, maiúscula, espaço e underline) e devolve o que foi
    salvo por último. Ignora arquivos temporários do Excel (~$...)."""
    candidatos = []
    for padrao in EXTENSOES_BASE:
        for caminho in glob.glob(os.path.join(pasta, padrao)):
            nome = os.path.basename(caminho)
            if nome.startswith("~$"):
                continue
            nome_normalizado = _limpa_maiuscula(nome).lower()
            if TRECHO_NOME_BASE in nome_normalizado:
                candidatos.append(caminho)

    if not candidatos:
        raise SystemExit(
            f"Não encontrei nenhuma base de Roteirização em:\n    {pasta}\n\n"
            f"Salve o export do Sankhya nessa pasta com '{TRECHO_NOME_BASE}' "
            "no nome (ex: 'Base Roteirizacao.xls')."
        )

    candidatos.sort(key=os.path.getmtime, reverse=True)
    escolhido = candidatos[0]

    if len(candidatos) > 1:
        print("Encontrei mais de uma base na pasta — usando a mais recente:")
        for c in candidatos:
            data = datetime.fromtimestamp(os.path.getmtime(c)).strftime("%d/%m/%Y %H:%M")
            marcador = "  -> " if c == escolhido else "     "
            print(f"{marcador}{os.path.basename(c)}  (salvo em {data})")
        print("  Dica: apague as bases antigas da pasta pra evitar confusão.\n")

    return escolhido


def _normaliza_nome_coluna(nome: str) -> str:
    nome = " ".join(str(nome).split())
    return nome.lower().replace(".", "").replace("?", "").strip()


def renomear_colunas_esperadas(df: "pd.DataFrame", caminho: str) -> "pd.DataFrame":
    obrigatorias = [
        COL_OC_PAI, COL_ORDEM_CARGA, COL_TIPO_OPERACAO, COL_COD_PARCEIRO,
        COL_CLIENTE, COL_COD_CIDADE_PARCEIRO, COL_ESTADO, COL_CIDADE_ESTADO,
        COL_TIPO_CARGA, COL_PESO, COL_DATA_AGENDAMENTO, COL_OBSERVACAO,
        COL_VALOR, COL_TIPO_ENTREGA, COL_PENDENTE, COL_PERFIL_PARCEIRO,
    ]
    opcionais = [COL_NRO_UNICO, COL_EMPRESA, COL_TOT_PALLET_PEDIDO]

    mapa_normalizado = {_normaliza_nome_coluna(c): c for c in df.columns}
    renomeio = {}
    faltando = []

    for esperada in obrigatorias:
        chave = _normaliza_nome_coluna(esperada)
        if chave in mapa_normalizado:
            coluna_real = mapa_normalizado[chave]
            if coluna_real != esperada:
                renomeio[coluna_real] = esperada
        else:
            faltando.append(esperada)

    for esperada in opcionais:
        chave = _normaliza_nome_coluna(esperada)
        if chave in mapa_normalizado:
            coluna_real = mapa_normalizado[chave]
            if coluna_real != esperada:
                renomeio[coluna_real] = esperada
        else:
            print(f"Aviso: coluna opcional '{esperada}' não encontrada — seguindo sem ela.")

    if faltando:
        colunas_da_planilha = "\n".join(f"  - {c!r}" for c in df.columns)
        raise SystemExit(
            "Não encontrei estas colunas na planilha:\n"
            + "\n".join(f"  - {c!r}" for c in faltando)
            + f"\n\nArquivo lido: {caminho}"
            + "\n\nColunas que a planilha realmente tem:\n" + colunas_da_planilha
            + "\n\nConfere se o nome mudou ou se tem cabeçalho extra antes das colunas."
        )

    return df.rename(columns=renomeio)


def carregar_base(caminho: str) -> pd.DataFrame:
    df = pd.read_excel(caminho, header=LINHA_CABECALHO)
    df.columns = [str(c).strip() for c in df.columns]
    df = renomear_colunas_esperadas(df, caminho)

    if COL_DATA_AGENDAMENTO in df.columns:
        df[COL_DATA_AGENDAMENTO] = pd.to_datetime(df[COL_DATA_AGENDAMENTO], errors="coerce")
    if COL_VALOR in df.columns:
        df[COL_VALOR] = df[COL_VALOR].apply(_para_numero_brl)
    if COL_PESO in df.columns:
        df[COL_PESO] = df[COL_PESO].apply(_para_numero_brl)
    if COL_TOT_PALLET_PEDIDO in df.columns:
        df[COL_TOT_PALLET_PEDIDO] = df[COL_TOT_PALLET_PEDIDO].apply(_para_numero_brl)
    if COL_EMPRESA in df.columns:
        df[COL_EMPRESA] = pd.to_numeric(df[COL_EMPRESA], errors="coerce")
    if COL_TIPO_OPERACAO in df.columns:
        df[COL_TIPO_OPERACAO] = pd.to_numeric(df[COL_TIPO_OPERACAO], errors="coerce")
    if COL_COD_CIDADE_PARCEIRO in df.columns:
        df[COL_COD_CIDADE_PARCEIRO] = pd.to_numeric(df[COL_COD_CIDADE_PARCEIRO], errors="coerce")
    for col in (COL_OC_PAI, COL_ORDEM_CARGA, COL_COD_PARCEIRO, COL_NRO_UNICO):
        if col in df.columns:
            df[col] = df[col].apply(_para_inteiro_texto)
    for col in (COL_CLIENTE, COL_ESTADO, COL_CIDADE_ESTADO, COL_TIPO_CARGA,
                COL_OBSERVACAO, COL_TIPO_ENTREGA, COL_PENDENTE, COL_PERFIL_PARCEIRO):
        if col in df.columns:
            # fillna("") antes do astype(str): em pandas >= 3.0 um NaN
            # convertido direto por astype(str) permanece como float NaN
            df[col] = df[col].fillna("").astype(str).str.strip()
    return df


# ----------------------------------------------------------------------
# LOTAÇÃO: agrupamento + distribuição em veículos + títulos
# ----------------------------------------------------------------------
def _titulo_contratacao(tipo_entrega: str, cidade_uf: str, peso_total: float,
                         tipo_veiculo: str, data_agendamento, origem: str = "") -> str:
    te = _limpa_maiuscula(tipo_entrega)
    if "AGENDADO COM NF" in te and pd.notna(data_agendamento):
        te = "AGENDADO"  # tem agenda marcada: é carga agendada normal (tem Contratação)
    peso_fmt = fmt_peso(peso_total)
    prefixo_origem = "AL X - " if origem == "Maceió" else ""
    if "AGENDADO COM NF" in te:
        return ""  # não existe título de contratação pra esse tipo
    if "AGENDADO" in te:
        data_str = data_agendamento.strftime("%d/%m/%Y %H:%M") if pd.notna(data_agendamento) else ""
        return f"{prefixo_origem}{cidade_uf} - {peso_fmt} - {tipo_veiculo} *AGENDADO {data_str}*"
    if "ORDEM DE CHEGADA" in te:
        return f"{prefixo_origem}{cidade_uf} - {peso_fmt} - {tipo_veiculo} *ORDEM DE CHEGADA*"
    return f"{prefixo_origem}{cidade_uf} - {peso_fmt} - {tipo_veiculo}"


def _titulo_roteirizacao(tipo_entrega: str, cidade_uf: str, cod_parceiro: str,
                          nome_parceiro: str, data_agendamento, observacao: str) -> str:
    te = _limpa_maiuscula(tipo_entrega)
    if "AGENDADO COM NF" in te and pd.notna(data_agendamento):
        te = "AGENDADO"  # tem agenda marcada: título igual ao de carga agendada
    if "AGENDADO COM NF" in te:
        return f"({cidade_uf}) {cod_parceiro} {nome_parceiro} *AGENDAR*"
    if "AGENDADO" in te:
        data_str = data_agendamento.strftime("%d/%m/%Y %H:%M") if pd.notna(data_agendamento) else ""
        texto = f"{nome_parceiro} *AGENDADO {data_str}*"
        if observacao and observacao.strip():
            texto += f" {observacao.strip()}"
        return texto
    if "ORDEM DE CHEGADA" in te:
        return f"{nome_parceiro} *ORDEM DE CHEGADA*"
    return nome_parceiro


def _distribuir_em_veiculos(itens: list, contexto: dict) -> list:
    """itens: lista de dicts {peso, valor, pallet, nro_unico}. Linha
    individual >= 27t sempre vira um RODOTREM sozinho (nunca por soma);
    o resto segue TRUCK (13t) / CARRETA (23t) na ordem em que aparece."""
    resultado = []
    normais = []

    for it in itens:
        if it["peso"] >= PESO_MIN_RODOTREM:
            resultado.append({
                "tipo_veiculo": "RODOTREM", "peso": it["peso"], "valor": it["valor"],
                "pallets": it["pallet"], "nros_unicos": [it["nro_unico"]] if it["nro_unico"] else [],
                "qtd_linhas": 1, **contexto,
            })
        else:
            normais.append(it)

    peso_restante = sum(it["peso"] for it in normais)
    atual = None

    for it in normais:
        peso_item, valor_item, pallet_item, nro = it["peso"], it["valor"], it["pallet"], it["nro_unico"]
        if peso_item == 0 and valor_item == 0 and pallet_item == 0:
            continue
        while True:
            if atual is None:
                tipo_veiculo = "CARRETA" if peso_restante > CAP_TRUCK else "TRUCK"
                capacidade = CAP_CARRETA if tipo_veiculo == "CARRETA" else CAP_TRUCK
                atual = {
                    "tipo_veiculo": tipo_veiculo, "capacidade": capacidade,
                    "peso": 0.0, "valor": 0.0, "pallets": 0.0,
                    "nros_unicos": [], "qtd_linhas": 0, **contexto,
                }
            if atual["peso"] == 0.0 or atual["peso"] + peso_item <= atual["capacidade"]:
                atual["peso"] += peso_item
                atual["valor"] += valor_item
                atual["pallets"] += pallet_item
                atual["qtd_linhas"] += 1
                if nro:
                    atual["nros_unicos"].append(nro)
                peso_restante = max(0.0, peso_restante - peso_item)
                if atual["peso"] >= atual["capacidade"]:
                    resultado.append(atual)
                    atual = None
                break
            else:
                resultado.append(atual)
                atual = None

    if atual is not None and (atual["peso"] > 0 or atual["valor"] > 0 or atual["pallets"] > 0):
        resultado.append(atual)

    return resultado


# ----------------------------------------------------------------------
# PRAZO DE SAÍDA DE UBERABA
# ----------------------------------------------------------------------
_COORD = None
_KM_MANUAL = None
DIAS_SEMANA = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"]


def _norm_cidade(s) -> str:
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().upper()
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _carregar_coordenadas():
    global _COORD, _KM_MANUAL
    if _COORD is None:
        _COORD = {}
        try:
            with open(ARQ_COORDENADAS, encoding="utf-8") as f:
                for linha in f:
                    if linha.startswith("#") or linha.startswith("uf;"):
                        continue
                    p = linha.rstrip("\n").split(";")
                    if len(p) >= 4:
                        _COORD[(p[0], p[1])] = (float(p[2]), float(p[3]))
        except OSError:
            print(f"  AVISO: não achei {ARQ_COORDENADAS} — o prazo de saída fica em branco.")
    if _KM_MANUAL is None:
        _KM_MANUAL = {}
        try:
            with open(ARQ_KM_MANUAL, encoding="utf-8") as f:
                for linha in f:
                    linha = linha.split(";")[0].strip()
                    if "=" in linha and not linha.startswith("["):
                        k, v = linha.split("=", 1)
                        try:
                            _KM_MANUAL[_norm_cidade(k)] = float(v.strip().replace(",", "."))
                        except ValueError:
                            pass
        except OSError:
            pass


def km_de_uberaba(cidade_uf: str, origem: str = "Uberaba"):
    """'SP - TAUBATE' -> (km, 'manual'|'estimado') ou (None, motivo). origem: 'Uberaba' ou 'Maceió'."""
    _carregar_coordenadas()
    chave = _norm_cidade(cidade_uf)
    if origem != "Maceió" and chave in _KM_MANUAL:
        return _KM_MANUAL[chave], "manual"
    m = re.match(r"^([A-Z]{2}) (.+)$", chave)
    if not m:
        return None, "cidade sem UF"
    uf, cidade = m.group(1), m.group(2)
    coord = _COORD.get((uf, cidade))
    if coord is None:
        return None, "cidade não encontrada"
    la1, lo1 = map(math.radians, MACEIO_LAT_LON if origem == "Maceió" else UBERABA_LAT_LON)
    la2, lo2 = map(math.radians, coord)
    d = math.sin((la2 - la1) / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2
    reta = 2 * 6371 * math.asin(math.sqrt(d))
    return round(reta * FATOR_ESTRADA), "estimado"


def _volta_para_dia_permitido(d, permite_sab, permite_dom):
    while (d.weekday() == 5 and not permite_sab) or (d.weekday() == 6 and not permite_dom):
        d -= timedelta(days=1)
    return d


def prazo_saida_uberaba(cidade_uf: str, data_agendamento, origem: str) -> dict:
    vazio = {"km": "", "km_fonte": "", "dias_viagem": "", "saida_uberaba": "", "expedicao_ate": "", "liberar_ate": "",
             "saida_ordem": "9999-99-99", "exped_ordem": "", "liberar_ordem": "", "status_prazo": "", "classe_prazo": ""}
    origem_calc = "Maceió" if origem == "Maceió" else "Uberaba"
    if data_agendamento is None or pd.isna(data_agendamento):
        return {**vazio, "saida_uberaba": "sem agenda", "classe_prazo": "prazo-na"}
    km, fonte = km_de_uberaba(cidade_uf, origem_calc)
    if km is None:
        return {**vazio, "saida_uberaba": f"? ({fonte})", "status_prazo": "ponha o km em distancias_uberaba.ini",
                "classe_prazo": "prazo-na"}
    dias = max(1, math.ceil(km / KM_POR_DIA))
    agenda = pd.Timestamp(data_agendamento).normalize().to_pydatetime().date()
    saida = _volta_para_dia_permitido(agenda - timedelta(days=dias), CARREGA_SABADO, CARREGA_DOMINGO)
    exped = saida
    for _ in range(DIAS_EXPEDICAO):
        exped = _volta_para_dia_permitido(exped - timedelta(days=1), False, False)
    liberar = exped
    for _ in range(DIAS_LIBERACAO):
        liberar = _volta_para_dia_permitido(liberar - timedelta(days=1), False, False)
    hoje = datetime.now().date()
    if saida < hoje:
        status, classe = "ATRASADO: já devia ter saído", "prazo-atrasado"
    elif exped < hoje:
        status, classe = "URGENTE: carregar e sair hoje", "prazo-atrasado"
    elif liberar < hoje:
        status, classe = "Liberação atrasada: liberar HOJE", "prazo-atrasado"
    elif liberar == hoje:
        status, classe = "Liberar HOJE", "prazo-hoje"
    elif (liberar - hoje).days <= 1:
        status, classe = "Liberar amanhã", "prazo-perto"
    else:
        status, classe = f"liberar em {(liberar - hoje).days} dias", "prazo-ok"
    fmt = lambda d: f"{d:%d/%m} ({DIAS_SEMANA[d.weekday()]})"
    return {"km": f"{km:.0f}", "km_fonte": fonte, "dias_viagem": dias, "saida_uberaba": fmt(saida),
            "expedicao_ate": fmt(exped), "liberar_ate": fmt(liberar), "saida_ordem": saida.isoformat(),
            "exped_ordem": exped.isoformat(), "liberar_ordem": liberar.isoformat(), "status_prazo": status,
            "classe_prazo": classe}


def _nome_origem(empresa) -> str:
    if pd.isna(empresa):
        return "—"
    try:
        cod = int(empresa)
    except (ValueError, TypeError):
        return "—"
    if cod == EMPRESA_UBERABA:
        return "Uberaba"
    if cod == EMPRESA_MACEIO:
        return "Maceió"
    return str(cod)


def classificar_lotacao_fracionado(universo: pd.DataFrame) -> pd.DataFrame:
    """Adiciona _tem_agendamento, _peso_total_parceiro e _is_fracionado ao
    universo. Usado tanto pra montar a Lotação quanto pra saber, dentro do
    universo de Maceió, o que é Fracionado de verdade pro Xi (que não
    tem recorte geográfico — é ele ou nada)."""
    universo = universo.copy()
    # "Agendado com NF" só fica de fora da isenção quando NÃO tem data
    # marcada (aí sim é avaliado pra Fracionado igual qualquer outro, pelo
    # peso). Com data marcada, conta como "tem agendamento" normalmente,
    # igual "Agendado" — vira Lotação de verdade, só fica sinalizado em
    # vermelho na tabela pra avisar que ainda é um NF.
    universo["_tem_agendamento"] = universo[COL_DATA_AGENDAMENTO].notna()

    sem_agendamento = universo[~universo["_tem_agendamento"]]
    peso_por_parceiro = sem_agendamento.groupby(COL_COD_PARCEIRO)[COL_PESO].sum()
    universo["_peso_total_parceiro"] = (
        universo[COL_COD_PARCEIRO].map(peso_por_parceiro).fillna(universo[COL_PESO])
    )
    universo["_is_fracionado"] = (
        (universo["_peso_total_parceiro"] < PESO_MAX_FRACIONADO) & (~universo["_tem_agendamento"])
    )
    return universo


def calcular_lotacao(universo_classificado: pd.DataFrame) -> dict:
    lotacao_rows = universo_classificado[~universo_classificado["_is_fracionado"]].copy()

    veiculos = []
    if not lotacao_rows.empty:
        lotacao_rows["_grp_key"] = (
            lotacao_rows[COL_COD_PARCEIRO].apply(_limpa_maiuscula) + "|"
            + lotacao_rows[COL_TIPO_ENTREGA].apply(_limpa_maiuscula) + "|"
            + lotacao_rows[COL_CIDADE_ESTADO].apply(_limpa_maiuscula) + "|"
            + lotacao_rows[COL_TIPO_CARGA].apply(_limpa_maiuscula) + "|"
            + lotacao_rows[COL_DATA_AGENDAMENTO].apply(
                lambda d: d.strftime("%Y-%m-%d") if pd.notna(d) else ""
            ) + "|"
            + lotacao_rows[COL_OBSERVACAO].apply(_limpa_maiuscula)
        )

        tem_pallet = COL_TOT_PALLET_PEDIDO in lotacao_rows.columns
        tem_nro = COL_NRO_UNICO in lotacao_rows.columns

        for _, grupo in lotacao_rows.groupby("_grp_key"):
            itens = [
                {
                    "peso": float(r[COL_PESO]),
                    "valor": float(r[COL_VALOR]),
                    "pallet": float(r[COL_TOT_PALLET_PEDIDO]) if tem_pallet else 0.0,
                    "nro_unico": r[COL_NRO_UNICO] if tem_nro else "",
                }
                for _, r in grupo.iterrows()
            ]
            contexto = {
                "cod": grupo[COL_COD_PARCEIRO].iloc[0],
                "nome": grupo[COL_CLIENTE].iloc[0] if COL_CLIENTE in grupo.columns else "",
                "cidade_uf": grupo[COL_CIDADE_ESTADO].iloc[0],
                "tipo_entrega": grupo[COL_TIPO_ENTREGA].iloc[0],
                "data_agendamento": grupo[COL_DATA_AGENDAMENTO].iloc[0],
                "observacao": grupo[COL_OBSERVACAO].iloc[0] if COL_OBSERVACAO in grupo.columns else "",
                "origem": _nome_origem(grupo[COL_EMPRESA].iloc[0]) if COL_EMPRESA in grupo.columns else "—",
            }
            veiculos.extend(_distribuir_em_veiculos(itens, contexto))

    veiculos_json = []
    veiculos_agendado_nf_json = []
    titulos_contratacao = []
    titulos_roteirizacao = []
    for v in sorted(veiculos, key=lambda x: -x["valor"]):
        te_upper = _limpa_maiuscula(v["tipo_entrega"])
        tipo_entrega_vazio = not v["tipo_entrega"] or not v["tipo_entrega"].strip()

        t_contr = _titulo_contratacao(v["tipo_entrega"], v["cidade_uf"], v["peso"], v["tipo_veiculo"], v["data_agendamento"], v.get("origem", ""))
        t_rot = _titulo_roteirizacao(v["tipo_entrega"], v["cidade_uf"], v["cod"], v["nome"], v["data_agendamento"], v["observacao"])
        if t_contr:
            titulos_contratacao.append(t_contr)
        if t_rot:
            titulos_roteirizacao.append(t_rot)

        eh_agendado_com_nf = "AGENDADO COM NF" in te_upper
        tem_data_agendamento = pd.notna(v["data_agendamento"])
        # se tem NF mas já tem data de agendamento marcada, trata como
        # Lotação normal (entra na tabela principal) — só destaca em
        # vermelho o Tipo de Entrega pra avisar que ainda é um NF
        destaque_nf = eh_agendado_com_nf and tem_data_agendamento

        linha_json = {
            "tipo_veiculo": v["tipo_veiculo"],
            "cidade_uf": v["cidade_uf"],
            "tipo_entrega": "Sem Tipo de Entrega" if tipo_entrega_vazio else v["tipo_entrega"],
            "destaque_nf": bool(destaque_nf),
            "precisa_customer_service": tipo_entrega_vazio,
            "nome": v["nome"],
            "origem": v.get("origem", "—"),
            "valor": fmt_moeda(v["valor"]),
            "peso": fmt_peso(v["peso"]),
            "pallets": round(v["pallets"], 2),
            "qtd_linhas": v["qtd_linhas"],
            "numeros_unicos": ", ".join(v["nros_unicos"]) if v["nros_unicos"] else "—",
            "titulo_contratacao": t_contr,
            "titulo_roteirizacao": t_rot,
            "agenda": v["data_agendamento"].strftime("%d/%m %H:%M") if pd.notna(v["data_agendamento"]) else "",
            "agenda_completa": v["data_agendamento"].strftime("%d/%m/%Y %H:%M") if pd.notna(v["data_agendamento"]) else "sem agenda",
            "observacao": "" if v.get("observacao") is None or (isinstance(v.get("observacao"), float) and pd.isna(v.get("observacao"))) else str(v.get("observacao")).strip(),
            **prazo_saida_uberaba(v["cidade_uf"], v["data_agendamento"], v.get("origem", "")),
            "_valor_num": float(v["valor"]),
            "_peso_num": float(v["peso"]),
            "uf": (str(v["cidade_uf"]).strip()[:2].upper() if v["cidade_uf"] else "—"),
        }

        if eh_agendado_com_nf and not tem_data_agendamento:
            veiculos_agendado_nf_json.append(linha_json)
        else:
            veiculos_json.append(linha_json)

    veiculos_json.sort(key=lambda x: (x["saida_ordem"], -x["_valor_num"]))
    for n, v in enumerate(veiculos_json + veiculos_agendado_nf_json, 1):
        v["n_carga"] = n
    # capacidade do CD por dia de carregamento (regra do Gerente: Uberaba 23 O.C.s/dia, Maceió 6)
    cont_dia = {}
    for v in veiculos_json:
        if v.get("exped_ordem"):
            ori = "Maceió" if v.get("origem") == "Maceió" else "Uberaba"
            cont_dia[(ori, v["exped_ordem"])] = cont_dia.get((ori, v["exped_ordem"]), 0) + 1
    hoje_iso = datetime.now().date().isoformat()
    capacidade = []
    for (ori, dia), qtd in sorted(cont_dia.items(), key=lambda x: (x[0][1], x[0][0])):
        if dia < hoje_iso:
            continue
        lim = CAPACIDADE_OC_DIA.get(ori, 999)
        d = datetime.strptime(dia, "%Y-%m-%d").date()
        capacidade.append({"origem": ori, "dia": f"{d:%d/%m} ({DIAS_SEMANA[d.weekday()]})", "qtd": qtd, "limite": lim,
                           "estourou": qtd > lim})
    for v in veiculos_json:
        ori = "Maceió" if v.get("origem") == "Maceió" else "Uberaba"
        v["dia_lotado"] = bool(v.get("exped_ordem")) and cont_dia.get((ori, v["exped_ordem"]), 0) > CAPACIDADE_OC_DIA.get(ori, 999)
    return {
        "capacidade": capacidade,
        "qtd_veiculos": len(veiculos),
        "valor": fmt_moeda(sum(v["valor"] for v in veiculos)),
        "peso": fmt_peso(sum(v["peso"] for v in veiculos)),
        "veiculos": veiculos_json,
        "veiculos_agendado_nf": veiculos_agendado_nf_json,
        "titulos_contratacao": titulos_contratacao,
        "titulos_roteirizacao": titulos_roteirizacao,
    }


# ----------------------------------------------------------------------
# FRACIONADO POR OPERADOR (replica os filtros do Sankhya)
# ----------------------------------------------------------------------
def calcular_fracionado_por_operador(df_frac_base: pd.DataFrame, tecmar_df: pd.DataFrame) -> list:
    """df_frac_base já deve estar filtrado por: Ordem de carga vazia e
    Pendente = Sim. As regras geográficas abaixo (Nu, Theta,
    Kappa, Epsilon, Iota) valem só pra Uberaba (Empresa 2) e pra quem
    não tiver Empresa preenchida — Maceió (Empresa 209) nunca entra aqui,
    porque lá é sempre Xi, independente da região. tecmar_df já vem
    filtrado pela MESMA regra de peso que decide Lotação x Fracionado (o
    Xi não tem recorte geográfico próprio, então sem esse corte de
    peso ele pegaria toda a Empresa 209, inclusive o que já virou Lotação)."""
    if COL_EMPRESA in df_frac_base.columns:
        eh_maceio = df_frac_base[COL_EMPRESA] == EMPRESA_MACEIO
    else:
        eh_maceio = pd.Series(False, index=df_frac_base.index)

    uberaba_df = df_frac_base[~eh_maceio]  # Uberaba + Empresa não identificada

    # blindagem extra: exclui pelo Nro. Único qualquer linha que já esteja
    # no Xi, mesmo que a Empresa dela não tenha vindo preenchida em
    # df_frac_base (evita contar a mesma carga em dois operadores ao
    # mesmo tempo — ex: Kappa + Xi)
    if COL_NRO_UNICO in tecmar_df.columns and COL_NRO_UNICO in uberaba_df.columns:
        nros_tecmar = set(tecmar_df[COL_NRO_UNICO]) - {""}
        if nros_tecmar:
            uberaba_df = uberaba_df[~uberaba_df[COL_NRO_UNICO].isin(nros_tecmar)]

    perfil_num = uberaba_df[COL_PERFIL_PARCEIRO].apply(_perfil_prefixo)
    nao_excluido = uberaba_df[COL_COD_PARCEIRO] != PARCEIRO_EXCLUIDO

    operadores = {}

    # XI: só a fatia da Empresa 209 (Maceió) que a mesma regra de peso
    # já classificaria como Fracionado (sem isso, ele engoliria tudo)
    operadores["Xi"] = tecmar_df

    nao_eh_perfil_ativa = ~perfil_num.isin(PERFIL_ATIVA_ESPECIAL)

    # NU: SP, exceto as cidades que eram da Omicron, o parceiro
    # excluído e o perfil especial da Theta (Omicron não existe mais como
    # operador — essas cidades ficam de fora de todo mundo por enquanto)
    mask = (
        (uberaba_df[COL_ESTADO] == "SP")
        & (~uberaba_df[COL_COD_CIDADE_PARCEIRO].isin(CIDADES_OMICRON))
        & nao_excluido
        & nao_eh_perfil_ativa
    )
    operadores["Nu"] = uberaba_df[mask]

    # THETA: ES/RJ sempre, ou MG/SP/RS/PR/SC quando o parceiro é Farma/Drogaria/Perfumaria
    mask = (
        uberaba_df[COL_ESTADO].isin(_ufs(16, 7))
        | (uberaba_df[COL_ESTADO].isin(_ufs(2, 1, 15, 8, 13)) & perfil_num.isin(PERFIL_ATIVA_ESPECIAL))
    )
    operadores["Theta"] = uberaba_df[mask]

    # KAPPA: nordeste + MA
    mask = uberaba_df[COL_ESTADO].isin(_ufs(6, 10, 17, 19, 21, 22, 23, 26, 31))
    operadores["Kappa"] = uberaba_df[mask]

    # EPSILON: GO, DF, MT, MS, PA, RO, RR, AC, AP, TO, AM
    mask = uberaba_df[COL_ESTADO].isin(_ufs(4, 3, 5, 12, 9, 11, 24, 20, 28, 14, 18))
    operadores["Epsilon"] = uberaba_df[mask]

    # IOTA: MG, SP, PR, SC, RS, todos os canais EXCETO Farma/Drogaria/
    # Perfumaria (esses vão pra Theta) e o parceiro excluído
    mask = (
        uberaba_df[COL_ESTADO].isin(_ufs(2, 1, 8, 13, 15))
        & nao_excluido
        & nao_eh_perfil_ativa
    )
    operadores["Iota"] = uberaba_df[mask]

    resultado = []
    for nome, subdf in operadores.items():
        ufs_presentes = sorted(u for u in subdf[COL_ESTADO].unique() if u)
        qtd_linhas = int(len(subdf))
        qtd_pedidos_unicos = (
            int(subdf[COL_NRO_UNICO].nunique()) if COL_NRO_UNICO in subdf.columns else qtd_linhas
        )
        resultado.append({
            "operador": nome,
            "qtd_pedidos": qtd_pedidos_unicos,
            "qtd_linhas": qtd_linhas,
            "valor": fmt_moeda(subdf[COL_VALOR].sum()),
            "peso": fmt_peso(subdf[COL_PESO].sum()),
            "ufs": ", ".join(ufs_presentes) if ufs_presentes else "—",
            "info": INFO_OPERADORES.get(nome, {}),
        })
    return resultado


# ----------------------------------------------------------------------
# MONTAGEM GERAL
# ----------------------------------------------------------------------
def montar_dados(df: pd.DataFrame) -> dict:
    universo = df[
        (df[COL_OC_PAI].str.strip() == "")
        & (df[COL_ORDEM_CARGA].str.strip() == "")
        & (df[COL_COD_PARCEIRO].str.strip() != "")
    ].copy()

    if not universo.empty:
        universo_classificado = classificar_lotacao_fracionado(universo)
        lotacao = calcular_lotacao(universo_classificado)
    else:
        universo_classificado = universo
        lotacao = {
            "qtd_veiculos": 0, "valor": fmt_moeda(0), "peso": fmt_peso(0),
            "veiculos": [], "veiculos_agendado_nf": [], "titulos_contratacao": [], "titulos_roteirizacao": [],
        }

    # base do fracionado (Xi + os operadores geográficos): usa a MESMA
    # classificação de peso/agendamento da Lotação (_is_fracionado) — ela já
    # decide sozinha se um "Agendado com NF" é fracionado (pedido < 7t e sem
    # data de agenda) ou lotação (pesado ou com data); não precisa excluir
    # "Agendado com NF" à parte, isso duplicaria a regra e cortaria pedido
    # que devia contar
    if not universo_classificado.empty and "_is_fracionado" in universo_classificado.columns:
        fracionado_geral = universo_classificado[
            universo_classificado["_is_fracionado"]
            & (universo_classificado[COL_PENDENTE].str.strip().str.lower() == "sim")
        ].copy()
    else:
        fracionado_geral = universo_classificado.iloc[0:0]

    if COL_EMPRESA in fracionado_geral.columns:
        tecmar_df = fracionado_geral[fracionado_geral[COL_EMPRESA] == EMPRESA_MACEIO]
    else:
        tecmar_df = fracionado_geral.iloc[0:0]

    fracionado_por_operador = (
        calcular_fracionado_por_operador(fracionado_geral, tecmar_df) if not fracionado_geral.empty else []
    )

    # lista de pedidos (um por linha) para a aba "Dados das cargas"
    carga_do_nro = {}
    for v in lotacao.get("veiculos", []) + lotacao.get("veiculos_agendado_nf", []):
        for nro in str(v.get("numeros_unicos", "")).split(","):
            if nro.strip() and nro.strip() != "—":
                col = v.get("saida_uberaba", "") or ""
                col = "—" if (not col or col == "sem agenda" or col.startswith("?")) else col
                carga_do_nro[nro.strip()] = (v.get("n_carga", 0), f"{v.get('n_carga', '')} · {v['tipo_veiculo']}", col)
    pedidos = []
    if not universo_classificado.empty:
        tem_pal = COL_TOT_PALLET_PEDIDO in universo_classificado.columns
        tem_nro = COL_NRO_UNICO in universo_classificado.columns
        for _, r in universo_classificado.iterrows():
            nro = str(r[COL_NRO_UNICO]).strip() if tem_nro else ""
            fr = bool(r.get("_is_fracionado", False))
            ag = r[COL_DATA_AGENDAMENTO]
            obs = r[COL_OBSERVACAO] if COL_OBSERVACAO in r.index else ""
            cid = str(r[COL_CIDADE_ESTADO] or "")
            pedidos.append({
                "nro": nro or "—",
                "carga": carga_do_nro.get(nro, (0, "Fracionado" if fr else "—"))[1],
                "_n": carga_do_nro.get(nro, (9999, ""))[0],
                "coleta": carga_do_nro.get(nro, (0, "", "—"))[2],
                "cliente": str(r[COL_CLIENTE] or ""), "cidade_uf": cid, "uf": cid.strip()[:2].upper() if cid else "—",
                "agenda": ag.strftime("%d/%m/%Y %H:%M") if pd.notna(ag) else "sem agenda",
                "paletes": float(r[COL_TOT_PALLET_PEDIDO]) if tem_pal and pd.notna(r[COL_TOT_PALLET_PEDIDO]) else 0.0,
                "peso": float(r[COL_PESO] or 0), "valor": float(r[COL_VALOR] or 0),
                "tipo_entrega": str(r[COL_TIPO_ENTREGA] or "").strip() or "Sem Tipo de Entrega",
                "observacao": "" if obs is None or (isinstance(obs, float) and pd.isna(obs)) else str(obs).strip(),
            })
        pedidos.sort(key=lambda p: (p["_n"], p["carga"], p["cliente"]))

    return {
        "pedidos": pedidos,
        "lotacao": lotacao,
        "fracionado_por_operador": fracionado_por_operador,
        "total_universo_valor": fmt_moeda(universo[COL_VALOR].sum()) if not universo.empty else fmt_moeda(0),
        "total_universo_peso": fmt_peso(universo[COL_PESO].sum()) if not universo.empty else fmt_peso(0),
        "qtd_linhas_universo": int(len(universo)),
    }


TEMPLATE_HTML = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<!-- sem cache: garante que o F5 sempre busque a versão mais nova do arquivo -->
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<title>Relatório de Roteirização</title>
<style>
  :root{{
    --bg:#0e1b10; --card:#152416; --border:#2a3f2c; --border-strong:#3d5a3f;
    --text:#e9f3e6; --text-2:#a7c2a3; --text-muted:#7c9a7b;
    --red-800:#ffb3ab; --red-600:#ff8a7a; --gray-50:#1b2e1c;
    --amber:#d9b34a; --blue:#5fb894;
  }}
  *{{box-sizing:border-box;}}
  body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif;
    background:var(--bg); color:var(--text); margin:0; padding:24px;}}
  .wrap{{max-width:none; width:calc(100vw - 40px); margin:0 auto;}}
  .logo-header{{display:flex; justify-content:center; margin-bottom:20px;}}
  .logo-header img{{max-width:220px; max-height:96px; width:auto; height:auto;}}
  h1{{font-size:20px; font-weight:500; margin:0; color:var(--text);}}
  .sub{{font-size:13px; color:var(--text-2); margin:4px 0 0;}}
  .topbar{{margin-bottom:20px;}}
  .fonte-base{{display:inline-block; margin-top:8px; font-size:12px; padding:4px 10px; border-radius:6px;
    background:var(--gray-50); border:0.5px solid var(--border-strong); color:var(--text-2);}}
  .fonte-base strong{{color:var(--text); font-weight:500;}}
  .fonte-base.antiga{{border-color:var(--amber); color:var(--amber);}}
  .cards{{display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:12px; margin-bottom:24px;}}
  .card{{background:var(--card); border:0.5px solid var(--border); border-radius:10px; padding:14px 16px;}}
  .card .lbl{{font-size:12px; color:var(--text-2); margin:0;}}
  .card .val{{font-size:22px; font-weight:500; margin:4px 0 0; color:var(--text);}}
  .section{{margin-bottom:22px;}}
  .section-head{{display:flex; align-items:center; gap:8px; margin-bottom:10px; flex-wrap:wrap;}}
  .section-head .title{{font-weight:500; font-size:15px; margin:0; color:var(--text);}}
  .section-head .note{{font-size:12px; color:var(--text-muted);}}
  table{{width:100%; border-collapse:collapse; font-size:13px;}}
  thead th{{text-align:left; padding:8px 12px; font-weight:500; color:var(--text-2); background:var(--gray-50); border-bottom:0.5px solid var(--border);}}
  tbody td{{padding:8px 12px; border-top:0.5px solid var(--border); vertical-align:middle; color:var(--text);}}
  td.num{{text-align:right;}}
  .tbl-card{{background:var(--card); border:0.5px solid var(--border); border-radius:12px; overflow:hidden; overflow-x:auto; margin-bottom:14px;}}
  .subhead{{font-size:13px; font-weight:500; margin:0; padding:12px 14px 4px; color:var(--text);}}
  footer{{font-size:12px; color:var(--text-muted); margin-top:24px; text-align:center;}}

  .btn-acao{{background:var(--blue); color:#0e1b10; border:none; border-radius:8px; padding:8px 14px;
    font-size:13px; font-weight:500; cursor:pointer;}}
  .btn-acao:hover{{opacity:0.9;}}
  .topo-dados{{display:flex; gap:12px; align-items:center; flex-wrap:wrap; margin-top:12px;}}
  .btn-grande{{font-size:15px !important; padding:11px 20px !important; font-weight:700;}}
  .abas{{display:flex; gap:6px; flex-wrap:wrap; margin:4px 0 16px; border-bottom:1px solid var(--border);}}
  .aba{{background:transparent; color:var(--text-2); border:1px solid transparent; border-bottom:none; border-radius:10px 10px 0 0;
    padding:10px 16px; font-size:14px; font-weight:600; cursor:pointer; margin-bottom:-1px;}}
  .aba.ativa{{background:var(--card); color:var(--blue); border-color:var(--border);}}
  .dados-box{{margin-bottom:14px;}} .dados-head{{display:flex; justify-content:space-between; align-items:center; gap:10px; padding:10px 12px; flex-wrap:wrap;}}
  #dadosCargasBody .btn-acao{{padding:4px 10px; font-size:12px;}} #dadosCargasBody td.num, #dadosCargasBody td:nth-child(4){{white-space:nowrap;}}
  .msg-copiado{{font-size:12px; color:var(--blue); margin-left:8px; display:none;}}

  .badge-cs{{display:inline-block; font-size:11px; padding:2px 8px; border-radius:6px; font-weight:500;
    background:rgba(255,138,122,0.15); color:var(--red-600); margin-left:6px;}}
  .destaque-nf{{color:var(--red-600); font-weight:600;}}
  .btn-info{{background:none; border:1px solid var(--border-strong); color:var(--text-2); border-radius:50%;
    width:22px; height:22px; font-size:12px; cursor:pointer; line-height:1;}}
  .btn-info:hover{{border-color:var(--blue); color:var(--blue);}}
  .linha-info td{{background:var(--gray-50); font-size:12px; color:var(--text-2); padding:10px 14px;}}
  .linha-info dl{{margin:0; display:grid; grid-template-columns:140px 1fr; gap:4px 10px;}}
  .linha-info dt{{color:var(--text-muted);}}
  .linha-info dd{{margin:0; color:var(--text);}}
  .hide{{display:none;}}
  .prazo-atrasado{{color:var(--red-600); font-weight:600;}}
  .prazo-hoje{{color:var(--amber); font-weight:600;}}
  .prazo-perto{{color:var(--amber);}}
  .prazo-ok{{color:var(--text-2);}}
  .cap-box{{margin:6px 0 10px; font-size:12px; color:var(--text-2); line-height:2;}}
  .cap-chip{{display:inline-block; padding:1px 8px; border-radius:10px; border:1px solid rgba(127,127,127,.35); margin-right:4px; white-space:nowrap;}}
  .cap-chip.cap-est{{background:rgba(255,138,122,.18); border-color:var(--red-600); color:var(--red-600); font-weight:600;}}
  .cap-alerta{{color:var(--red-600); font-weight:600; margin-top:2px;}}
  .badge-cap{{font-size:10px; padding:0 6px; border-radius:8px; background:rgba(255,138,122,.2); color:var(--red-600); white-space:nowrap;}}
  .prazo-na{{color:var(--text-muted);}}
  td.prazo b{{font-weight:600;}}
  #lotacaoBody td{{white-space:nowrap; padding:6px 8px; font-size:12.5px;}}
  #lotacaoBody td.cliente{{max-width:260px; overflow:hidden; text-overflow:ellipsis;}}
  #lotacaoBody td.nros{{white-space:normal; font-size:11.5px; color:var(--text-2); max-width:200px;}}
  #lotacaoBody td.situacao{{font-size:12px;}}
  #lotacaoBody td.botoes .btn-acao{{padding:5px 10px; font-size:12px; margin-right:4px;}}
  table thead th{{white-space:nowrap; padding:8px 8px; font-size:12px;}}
  tr.linha-prazo-atrasado td{{background:rgba(255,138,122,0.07);}}
  .filtro-uf{{display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin:0 0 10px;}}
  .filtro-lbl{{font-size:12px; color:var(--text-2);}}
  .chips{{display:flex; gap:6px; flex-wrap:wrap;}}
  .chip{{background:var(--card); border:0.5px solid var(--border-strong); color:var(--text-2); border-radius:14px;
    padding:4px 10px; font-size:12px; cursor:pointer;}}
  .chip:hover{{border-color:var(--blue); color:var(--text);}}
  .chip.ativo{{background:var(--blue); color:#0e1b10; border-color:var(--blue); font-weight:600;}}
  .chip .qtd{{opacity:0.7; margin-left:4px;}}
  .chip.tem-urgente{{border-color:var(--red-600);}}
  .filtro-resumo{{font-size:12px; color:var(--text-muted);}}
  .marc-info{{font-size:12px; color:var(--text-2); margin-left:auto; display:flex; align-items:center; gap:8px;}}
  #lotacaoBody tr, #agendadoNfBody tr{{cursor:pointer;}}
  #lotacaoBody tr:hover td, #agendadoNfBody tr:hover td{{background:rgba(127,178,229,0.08);}}
  tr.marcada td, #lotacaoBody tr.marcada td, #agendadoNfBody tr.marcada td{{background:#7a1f17 !important; color:#ffffff !important;}}
  tr.marcada td *{{color:#ffffff !important;}}
  tr.marcada .btn-acao{{background:#ffffff; color:#7a1f17 !important;}}
  tr.linha-prazo-hoje td{{background:rgba(217,179,74,0.07);}}
</style>
</head>
<body>
<div class="wrap">
  <div class="logo-header">
    {logo_tag}
  </div>
  <div class="topbar">
    <h1>Relatório de Roteirização</h1>
    <p class="sub">Universo: O.C. PAI vazio, Ordem de carga vazia &middot; gerado em {gerado_em}</p>
    <span class="fonte-base{classe_base_antiga}">Base lida: <strong>{arquivo_base}</strong> &middot; salva em {base_modificada} &middot; {qtd_linhas_base} linhas{aviso_base_antiga}</span>
    <div class="topo-dados">
      <button type="button" class="btn-acao btn-grande" onclick="copiarTodasTexto()">📋 Copiar dados de TODAS as cargas</button>
      <span class="sub" id="msgTodas">cliente, destino, agendamento da entrega, data da coleta, veículo, paletes, peso, valor e senha de cada carga, de uma vez só</span>
    </div>
  </div>

  <div class="cards">
    <div class="card"><p class="lbl">Total no universo (Lotação + Fracionado)</p><p class="val">{total_universo_valor}</p><p class="sub">{total_universo_peso}</p></div>
    <div class="card"><p class="lbl">Lotação</p><p class="val" id="lotacaoValor"></p><p class="sub" id="lotacaoPeso"></p></div>
  </div>

  <div class="abas">
    <button type="button" class="aba ativa" id="abaBtnRot" onclick="mostrarAba('rot')">Roteirização</button>
    <button type="button" class="aba" id="abaBtnDados" onclick="mostrarAba('dados')">Dados das cargas (paletes, peso, valor, cidade, agenda)</button>
  </div>

  <div class="section" id="secDados" style="display:none;">
    <div class="section-head"><p class="title">Dados das cargas &mdash; os mesmos veículos da aba Roteirização</p>
      <span class="note">uma linha por carga · respeita o filtro de estado da aba Roteirização</span></div>
    <div class="tbl-card dados-box">
      <div class="dados-head"><b id="dadosResumo"></b>
        <span><span class="msg-copiado" id="msgCopiado2">Copiado!</span>
        <button type="button" class="btn-acao" onclick="copiarDadosTodas()">Copiar cargas (cola no Excel)</button></span></div>
      <table><thead><tr><th>#</th><th>Veículo</th><th>Cliente</th><th>Cidade/UF</th><th>Agenda da entrega</th><th>Coletar até</th><th class="num">Paletes</th>
        <th class="num">Peso</th><th class="num">Valor da nota</th><th class="num">Pedidos</th><th>Entrega</th><th>Obs.</th><th></th></tr></thead>
        <tbody id="dadosCargasBody"></tbody></table>
    </div>
    <div class="section-head" style="margin-top:18px;"><p class="title">Pedidos (um por linha)</p>
      <span class="note">todos os pedidos do universo da base, com a carga em que cada um entrou</span>
      <button type="button" class="btn-acao" onclick="copiarPedidos()">Copiar pedidos (cola no Excel)</button></div>
    <div class="tbl-card">
      <table><thead><tr><th>Nro. Único</th><th>Carga</th><th>Cliente</th><th>Cidade/UF</th><th>Agenda da entrega</th><th>Coletar até</th><th class="num">Paletes</th>
        <th class="num">Peso</th><th class="num">Valor da nota</th><th>Entrega</th><th>Obs.</th></tr></thead>
        <tbody id="pedidosBody"></tbody></table>
    </div>
  </div>

  <div class="section">
    <div class="section-head">
      <p class="title">Lotação &mdash; veículos formados (mais urgente primeiro)</p>
      <span class="note">TRUCK até 13t &middot; CARRETA até 23t &middot; RODOTREM &ge;27t &middot; Regra do CD (Gerente): <b>Liberar até</b> (D-0) &rarr; <b>Carregar até</b> (D+1, preparação e carregamento) &rarr; <b>Coletar até</b> (D+2, início do trânsito a {km_por_dia} km/dia). Vale para Uberaba e Maceió &middot; km estimado x {fator_estrada} ou do distancias_uberaba.ini &middot; capacidade: Uberaba 23 O.C.s/dia, Maceió 6</span>
    </div>
    <div id="capacidadeBox" class="cap-box"></div>
    <div class="filtro-uf">
      <span class="filtro-lbl">Filtrar por estado:</span>
      <div id="chipsUf" class="chips"></div>
      <span id="resumoFiltro" class="filtro-resumo"></span>
      <span class="marc-info"><span id="qtdMarcadas">0</span> linha(s) marcada(s) &middot; clique na linha para marcar/desmarcar
        <button type="button" class="chip" onclick="limparMarcadas()">Limpar marcações</button></span>
    </div>
    <span class="msg-copiado" id="msgCopiado" style="margin-bottom:8px; display:none;">Copiado!</span>
    <div class="tbl-card">
      <table>
        <thead>
          <tr>
            <th>Liberar até</th><th>Carregar até</th><th>Coletar até</th><th>Agenda</th><th>Situação</th>
            <th>Veículo</th><th>Origem</th><th>Cliente</th><th>Cidade/UF</th><th>Entrega</th>
            <th class="num">Km · dias</th><th class="num">Peso</th><th class="num">Valor</th><th class="num">Lin.</th>
            <th>Números únicos</th><th>Títulos</th>
          </tr>
        </thead>
        <tbody id="lotacaoBody"></tbody>
      </table>
    </div>
  </div>

  <div class="section">
    <div class="section-head">
      <p class="title">Agendado com NF SEM data de agenda &mdash; separado da Lotação</p>
      <span class="note">ainda falta agendar (título *AGENDAR*). Os "Agendado com NF" que JÁ têm data estão lá em cima, como carga agendada normal, com Contratação</span>
    </div>
    <div class="tbl-card">
      <table>
        <thead>
          <tr>
            <th>Tipo</th><th>Origem</th><th>Cliente</th><th>Cidade/UF</th>
            <th>Peso</th><th>Valor</th><th>Linhas somadas</th><th>Números únicos</th><th>Título Roteirização</th>
          </tr>
        </thead>
        <tbody id="agendadoNfBody"></tbody>
      </table>
    </div>
  </div>

  <div class="section">
    <div class="section-head">
      <p class="title">Fracionado por operador</p>
      <span class="note">replica os filtros que o time já usa no Sankhya pra puxar a lista de cada operador &middot; Empresa 209 (Maceió) é sempre Xi</span>
    </div>
    <div class="tbl-card">
      <table>
        <thead><tr><th></th><th>Operador</th><th>Qtde. pedidos (únicos)</th><th>Qtde. linhas</th><th>Valor</th><th>Peso</th><th>UFs</th></tr></thead>
        <tbody id="fracionadoBody"></tbody>
      </table>
    </div>
    <p class="sub">Um pedido pode aparecer em mais de um operador quando os critérios se sobrepõem (ex: SP entre Nu/Theta/Iota) — é intencional, pra mostrar todas as opções possíveis.</p>
  </div>

  <footer>Gerado em {gerado_em} a partir de {arquivo_base}. Se este horário não for o de agora, você está olhando um relatório antigo — rode o script de novo (ele abre o arquivo certo sozinho).</footer>
</div>

<script>
const lotacao = {lotacao_json};
const pedidosBase = {pedidos_json};
const fracionadoPorOperador = {fracionado_por_operador_json};

document.getElementById('lotacaoValor').textContent = lotacao.valor;
document.getElementById('lotacaoPeso').textContent = lotacao.qtd_veiculos + ' veículo(s) · ' + lotacao.peso;

let ufsSelecionadas = new Set();   // vazio = todos
try {{ const salvo = JSON.parse(localStorage.getItem('rotFiltroUf') || '[]'); ufsSelecionadas = new Set(salvo); }} catch(e) {{}}
function passaFiltro(v){{ return ufsSelecionadas.size === 0 || ufsSelecionadas.has(v.uf); }}
function fmtBR(n, dec){{ return n.toLocaleString('pt-BR', {{minimumFractionDigits: dec, maximumFractionDigits: dec}}); }}

function montarChips(){{
  const todos = lotacao.veiculos.concat(lotacao.veiculos_agendado_nf);
  const cont = {{}}, urg = {{}};
  todos.forEach(v => {{ cont[v.uf] = (cont[v.uf] || 0) + 1; if (v.classe_prazo === 'prazo-atrasado') urg[v.uf] = true; }});
  const ufs = Object.keys(cont).sort();
  // tira do filtro salvo o estado que não existe mais na base de hoje
  ufsSelecionadas = new Set([...ufsSelecionadas].filter(u => cont[u]));
  // tira marcações de cargas que não estão mais na base
  const chavesHoje = new Set(todos.map(chaveLinha));
  marcadas = new Set([...marcadas].filter(k => chavesHoje.has(k)));
  let h = `<button type="button" class="chip ${{ufsSelecionadas.size === 0 ? 'ativo' : ''}}" onclick="filtrarUf('')">Todos<span class="qtd">${{todos.length}}</span></button>`;
  h += ufs.map(u => `<button type="button" class="chip ${{ufsSelecionadas.has(u) ? 'ativo' : ''}} ${{urg[u] ? 'tem-urgente' : ''}}"
      title="${{urg[u] ? 'tem carga atrasada / urgente' : ''}}" onclick="filtrarUf('${{u}}', event)">${{u}}<span class="qtd">${{cont[u]}}</span></button>`).join('');
  document.getElementById('chipsUf').innerHTML = h;
}}

function filtrarUf(uf, ev){{
  if (!uf) ufsSelecionadas.clear();
  else if (ev && (ev.ctrlKey || ev.metaKey)) {{ ufsSelecionadas.has(uf) ? ufsSelecionadas.delete(uf) : ufsSelecionadas.add(uf); }}
  else if (ufsSelecionadas.size === 1 && ufsSelecionadas.has(uf)) ufsSelecionadas.clear();
  else {{ ufsSelecionadas = new Set([uf]); }}
  try {{ localStorage.setItem('rotFiltroUf', JSON.stringify([...ufsSelecionadas])); }} catch(e) {{}}
  renderTabelas();
}}

// ---- marcar linhas (fica salvo no navegador, vale mesmo gerando o relatório de novo)
let marcadas = new Set();
try {{ marcadas = new Set(JSON.parse(localStorage.getItem('rotLinhasMarcadas') || '[]')); }} catch(e) {{}}
function chaveLinha(v){{ return (v.numeros_unicos && v.numeros_unicos !== '—') ? v.numeros_unicos : (v.cidade_uf + '|' + v.nome + '|' + v.peso); }}
function estaMarcada(v){{ return marcadas.has(chaveLinha(v)); }}
function salvarMarcadas(){{
  try {{ localStorage.setItem('rotLinhasMarcadas', JSON.stringify([...marcadas])); }} catch(e) {{}}
  const el = document.getElementById('qtdMarcadas');
  if (el) el.textContent = document.querySelectorAll('tr.marcada').length;
}}
function clicarLinha(tr, ev){{
  const k = tr.dataset.chave;
  const noBotao = ev && ev.target && ev.target.closest('button');
  if (noBotao) {{ marcadas.add(k); tr.classList.add('marcada'); }}      // copiou o título = marca como feita
  else if (marcadas.has(k)) {{ marcadas.delete(k); tr.classList.remove('marcada'); }}
  else {{ marcadas.add(k); tr.classList.add('marcada'); }}
  salvarMarcadas();
}}
function limparMarcadas(){{
  marcadas.clear();
  document.querySelectorAll('tr.marcada').forEach(t => t.classList.remove('marcada'));
  salvarMarcadas();
}}

function renderTabelas(){{
  montarChips();
  const vis = lotacao.veiculos.filter(passaFiltro);
  const peso = vis.reduce((a, v) => a + v._peso_num, 0), valor = vis.reduce((a, v) => a + v._valor_num, 0);
  document.getElementById('resumoFiltro').textContent = ufsSelecionadas.size === 0 ? 'Ctrl+clique para escolher mais de um estado'
    : `${{vis.length}} veículo(s) · ${{fmtBR(peso, 0)}} kg · R$ ${{fmtBR(valor, 2)}}  (Ctrl+clique soma estados)`;
document.getElementById('lotacaoBody').innerHTML = lotacao.veiculos.map((v, i) => !passaFiltro(v) ? '' : `
  <tr class="linha-${{v.classe_prazo}} ${{estaMarcada(v) ? 'marcada' : ''}}" data-chave="${{chaveLinha(v)}}" onclick="clicarLinha(this, event)">
    <td class="${{v.classe_prazo}}"><b>${{v.liberar_ate || '—'}}</b></td>
    <td class="${{v.classe_prazo}}">${{v.expedicao_ate || '—'}}${{v.dia_lotado ? ' <span class="badge-cap" title="Neste dia de carregamento o CD passa do limite de O.C.s">dia lotado</span>' : ''}}</td>
    <td class="${{v.classe_prazo}}">${{v.saida_uberaba}}</td>
    <td>${{v.agenda || '—'}}</td>
    <td class="situacao ${{v.classe_prazo}}">${{v.status_prazo || ''}}</td>
    <td>${{v.tipo_veiculo}}</td>
    <td>${{v.origem}}</td>
    <td class="cliente" title="${{v.nome}}">${{v.nome}}</td>
    <td>${{v.cidade_uf}}</td>
    <td>${{v.destaque_nf ? `<span class="destaque-nf" title="Agendado com NF, mas já tem agenda: tratado como carga agendada">${{v.tipo_entrega}}</span>` : v.tipo_entrega}}${{v.precisa_customer_service ? '<span class="badge-cs">Acionar Customer Service</span>' : ''}}</td>
    <td class="num" title="${{v.km_fonte}}">${{v.km ? v.km + ' km · ' + v.dias_viagem + 'd' : '—'}}${{v.km_fonte === 'manual' ? ' ✓' : ''}}</td>
    <td class="num">${{v.peso}}</td>
    <td class="num">${{v.valor}}</td>
    <td class="num">${{v.qtd_linhas}}</td>
    <td class="nros">${{v.numeros_unicos}}</td>
    <td class="botoes">
      ${{v.titulo_contratacao ? `<button type="button" class="btn-acao" onclick="copiarTituloLinha(${{i}}, 'contratacao')" title="${{v.titulo_contratacao}}">Contratação</button>` : ''}}
      ${{v.titulo_roteirizacao ? `<button type="button" class="btn-acao" onclick="copiarTituloLinha(${{i}}, 'roteirizacao')" title="${{v.titulo_roteirizacao}}">Roteirização</button>` : ''}}
    </td>
  </tr>`).join('') || '<tr><td colspan="16" class="sub">Nenhum veículo de lotação para este filtro.</td></tr>';
  renderDadosCargas();
  const cap = lotacao.capacidade || [];
  const est = cap.filter(c => c.estourou);
  document.getElementById('capacidadeBox').innerHTML = !cap.length ? '' :
    `<b>Carregamentos por dia (todas as lotações agendadas):</b> ` + cap.map(c => `<span class="cap-chip ${{c.estourou ? 'cap-est' : ''}}">${{c.origem}} ${{c.dia}}: ${{c.qtd}}/${{c.limite}}</span>`).join(' ')
    + (est.length ? `<div class="cap-alerta">⚠ Passou da capacidade: ${{est.map(c => `${{c.origem}} ${{c.dia}} (${{c.qtd}} O.C.s, limite ${{c.limite}} — antecipar ${{c.qtd - c.limite}})`).join(' · ')}}</div>` : '');



document.getElementById('agendadoNfBody').innerHTML = lotacao.veiculos_agendado_nf.map((v, i) => !passaFiltro(v) ? '' : `
  <tr class="${{estaMarcada(v) ? 'marcada' : ''}}" data-chave="${{chaveLinha(v)}}" onclick="clicarLinha(this, event)">
    <td>${{v.tipo_veiculo}}</td>
    <td>${{v.origem}}</td>
    <td>${{v.nome}}</td>
    <td>${{v.cidade_uf}}</td>
    <td class="num">${{v.peso}}</td>
    <td class="num">${{v.valor}}</td>
    <td class="num">${{v.qtd_linhas}}</td>
    <td>${{v.numeros_unicos}}</td>
    <td style="white-space:nowrap;">
      ${{v.titulo_roteirizacao ? `<button type="button" class="btn-acao" onclick="copiarTituloLinhaNf(${{i}})">Roteirização</button>` : ''}}
    </td>
  </tr>`).join('') || '<tr><td colspan="9" class="sub">Nenhum Agendado com NF (sem data) para este filtro.</td></tr>';
}}
renderTabelas();
salvarMarcadas();

document.getElementById('fracionadoBody').innerHTML = fracionadoPorOperador.map((o, i) => `
  <tr>
    <td><button type="button" class="btn-info" onclick="toggleInfoOperador(${{i}})" title="Como funciona ${{o.operador}}">?</button></td>
    <td>${{o.operador}}</td>
    <td class="num">${{o.qtd_pedidos}}</td>
    <td class="num">${{o.qtd_linhas}}</td>
    <td class="num">${{o.valor}}</td>
    <td class="num">${{o.peso}}</td>
    <td>${{o.ufs}}</td>
  </tr>
  <tr class="linha-info hide" id="infoOperador-${{i}}">
    <td colspan="7">
      <dl>
        <dt>Coleta</dt><dd>${{o.info.coleta || '—'}}</dd>
        <dt>Criação de carga</dt><dd>${{o.info.criacao_carga || '—'}}</dd>
        <dt>Peso</dt><dd>${{o.info.peso || '—'}}</dd>
        <dt>Data de coleta</dt><dd>${{o.info.data_coleta || '—'}}</dd>
      </dl>
    </td>
  </tr>`).join('') || '<tr><td colspan="7" class="sub">Nada de fracionado disponível agora.</td></tr>';

function toggleInfoOperador(i){{
  document.getElementById('infoOperador-' + i).classList.toggle('hide');
}}

function _copiarTexto(texto, callback){{
  if(navigator.clipboard && navigator.clipboard.writeText){{
    navigator.clipboard.writeText(texto).then(callback).catch(() => _copiarFallback(texto, callback));
  }} else {{
    _copiarFallback(texto, callback);
  }}
}}

function _copiarFallback(texto, callback){{
  const area = document.createElement('textarea');
  area.value = texto;
  area.style.position = 'fixed';
  area.style.opacity = '0';
  document.body.appendChild(area);
  area.select();
  try {{ document.execCommand('copy'); }} catch(e) {{}}
  document.body.removeChild(area);
  if(callback) callback();
}}

function _mostrarCopiado(){{
  const msg = document.getElementById('msgCopiado');
  msg.style.display = 'inline';
  setTimeout(() => {{ msg.style.display = 'none'; }}, 2000);
}}

function copiarTituloLinha(indice, tipo){{
  const v = lotacao.veiculos[indice];
  if(!v) return;
  const texto = tipo === 'contratacao' ? _linhaContratacao(v) : v.titulo_roteirizacao;
  if(!texto) return;
  _copiarTexto(texto, _mostrarCopiado);
}}

function _linhaContratacao(v){{
  // MG-UBERABA 22 TON CARRETA *AGENDADO* 13/11/2026  (ou *ORDEM DE CHEGADA*)
  const t = v.titulo_contratacao || '';
  const cu = (v.cidade_uf || '').split(' - ');
  const destino = cu.length > 1 ? (cu[0].trim() + '-' + cu.slice(1).join(' - ').trim()).toUpperCase() : (v.cidade_uf || '').toUpperCase();
  let kg = typeof v._peso_num === 'number' ? v._peso_num : parseFloat(String(v.peso || '').replace(/[^0-9,]/g, '').replace(',', '.')) || 0;
  const peso = kg >= 1000 ? (kg >= 3000 ? Math.round(kg / 1000) + ' TON' : (kg / 1000).toFixed(1).replace('.', ',') + ' TON') : Math.round(kg) + ' KG';
  const veic = (v.tipo_veiculo || '').toUpperCase();
  let tipo;
  if(/CHEGADA/i.test(t) || /CHEGADA/i.test(v.tipo_entrega || '')) tipo = '*ORDEM DE CHEGADA*';
  else {{ const m = String(v.agenda_completa || '').match(/\\d{{2}}\\/\\d{{2}}\\/\\d{{4}}/) || t.match(/\\d{{2}}\\/\\d{{2}}\\/\\d{{4}}/); tipo = '*AGENDADO*' + (m ? ' ' + m[0] : ''); }}
  const al = (v.origem === 'Maceió' || /^AL X/i.test(t)) ? 'AL X ' : '';
  return [al + destino, peso, veic, tipo].filter(x => x).join(' ');
}}
function _textoDados(v){{
  const l = [`Cliente: ${{v.nome}}`, `Destino: ${{v.cidade_uf}}`, `Agendamento da entrega: ${{v.agenda_completa}}`,
    `Coletar até: ${{_coleta(v)}}`, `Veículo: ${{v.tipo_veiculo}}`,
    `Paletes: ${{fmtBR(v.pallets || 0, 2)}}`, `Peso: ${{v.peso}}`, `Valor da nota: ${{v.valor}}`, `Entrega: ${{v.tipo_entrega}}`];
  if(v.observacao) l.push(`Obs.: ${{v.observacao}}`);
  if(v.numeros_unicos && v.numeros_unicos !== '—') l.push(`Números únicos: ${{v.numeros_unicos}}`);
  return l.join('\\n');
}}
function copiarTodasTexto(){{
  const t = _todasCargas();
  if(!t.length){{ document.getElementById('msgTodas').textContent = 'Nenhuma carga para copiar (confira o filtro de estado).'; return; }}
  const filtro = ufsSelecionadas.size ? ` (estados: ${{[...ufsSelecionadas].join(', ')}})` : '';
  const texto = `DADOS DAS CARGAS · ${{t.length}} carga(s)${{filtro}}\\n\\n` + t.map(v => `CARGA ${{v.n_carga}}\\n` + _textoDados(v)).join('\\n\\n');
  _copiarTexto(texto, () => {{
    const m = document.getElementById('msgTodas');
    m.textContent = `✔ Copiado: ${{t.length}} carga(s). É só colar (Ctrl+V) no e-mail, Teams ou WhatsApp.`;
    m.style.color = 'var(--blue)';
  }});
}}
function _coleta(v){{ const c = v.saida_uberaba || ''; return (!c || c === 'sem agenda' || c.startsWith('?')) ? 'sem data (sem agenda ou sem km)' : c; }}
function copiarDadosLinha(v){{ if(v) _copiarTexto(_textoDados(v), _mostrarCopiado); }}
function _todasCargas(){{ return lotacao.veiculos.concat(lotacao.veiculos_agendado_nf).filter(passaFiltro); }}
function renderDadosCargas(){{
  const box = document.getElementById('dadosCargasBody'); if(!box) return;
  const t = _todasCargas();
  const pal = t.reduce((a, v) => a + (v.pallets || 0), 0), peso = t.reduce((a, v) => a + v._peso_num, 0), val = t.reduce((a, v) => a + v._valor_num, 0);
  document.getElementById('dadosResumo').textContent = `${{t.length}} carga(s) · ${{fmtBR(pal, 2)}} paletes · ${{fmtBR(peso, 0)}} kg · R$ ${{fmtBR(val, 2)}}`;
  box.innerHTML = t.map((v, i) => `<tr>
    <td>${{v.n_carga || ''}}</td><td>${{v.tipo_veiculo}}</td><td class="cliente" title="${{v.nome}}">${{v.nome}}</td><td>${{v.cidade_uf}}</td><td>${{v.agenda_completa}}</td><td><b>${{_coleta(v)}}</b></td>
    <td class="num">${{fmtBR(v.pallets || 0, 2)}}</td><td class="num">${{v.peso}}</td><td class="num">${{v.valor}}</td><td class="num">${{v.qtd_linhas}}</td>
    <td>${{v.tipo_entrega}}</td><td>${{v.observacao || ''}}</td>
    <td><button type="button" class="btn-acao" onclick="copiarDadosLinha(_todasCargas()[${{i}}])">Copiar</button></td></tr>`).join('')
    || '<tr><td colspan="13" class="sub">Nenhuma carga para este filtro.</td></tr>';
  const ped = pedidosBase.filter(p => ufsSelecionadas.size === 0 || ufsSelecionadas.has(p.uf));
  document.getElementById('pedidosBody').innerHTML = ped.map(p => `<tr>
    <td>${{p.nro}}</td><td>${{p.carga}}</td><td class="cliente" title="${{p.cliente}}">${{p.cliente}}</td><td>${{p.cidade_uf}}</td><td>${{p.agenda}}</td><td>${{p.coleta}}</td>
    <td class="num">${{fmtBR(p.paletes, 2)}}</td><td class="num">${{fmtBR(p.peso, 0)}} kg</td><td class="num">R$ ${{fmtBR(p.valor, 2)}}</td>
    <td>${{p.tipo_entrega}}</td><td>${{p.observacao}}</td></tr>`).join('')
    || '<tr><td colspan="11" class="sub">Nenhum pedido para este filtro.</td></tr>';
}}
function mostrarAba(nome){{
  document.querySelectorAll('.section').forEach(sec => {{ sec.style.display = (nome === 'dados') === (sec.id === 'secDados') ? '' : 'none'; }});
  document.getElementById('abaBtnRot').classList.toggle('ativa', nome !== 'dados');
  document.getElementById('abaBtnDados').classList.toggle('ativa', nome === 'dados');
  if(nome === 'dados') renderDadosCargas();
  try {{ history.replaceState(null, '', nome === 'dados' ? '#dados' : '#'); }} catch(e) {{}}
}}
function _copiado2(){{ _mostrarCopiado(); const m = document.getElementById('msgCopiado2'); if(m){{ m.style.display='inline'; setTimeout(() => m.style.display='none', 2000); }} }}
function copiarPedidos(){{
  const cab = ['Nro. Único','Carga','Cliente','Cidade/UF','Agendamento da entrega','Coletar até','Paletes','Peso (kg)','Valor da nota (R$)','Entrega','Obs.'];
  const lin = pedidosBase.filter(p => ufsSelecionadas.size === 0 || ufsSelecionadas.has(p.uf)).map(p => [p.nro, p.carga, p.cliente, p.cidade_uf, p.agenda, p.coleta,
    fmtBR(p.paletes, 2), fmtBR(p.peso, 0), fmtBR(p.valor, 2), p.tipo_entrega, p.observacao].join('\\t'));
  _copiarTexto([cab.join('\\t')].concat(lin).join('\\n'), _copiado2);
}}
function toggleDadosCargas(){{
  const box = document.getElementById('dadosCargasBox');
  box.style.display = box.style.display === 'none' ? 'block' : 'none';
  renderDadosCargas();
}}
function copiarDadosTodas(){{
  const cab = ['Veículo','Cliente','Cidade/UF','Agendamento da entrega','Coletar até','Paletes','Peso (kg)','Valor da nota (R$)','Entrega','Obs.','Números únicos'];
  const lin = _todasCargas().map(v => [v.tipo_veiculo, v.nome, v.cidade_uf, v.agenda_completa, _coleta(v), fmtBR(v.pallets || 0, 2),
    fmtBR(v._peso_num, 0), fmtBR(v._valor_num, 2), v.tipo_entrega, v.observacao || '', v.numeros_unicos].join('\\t'));
  _copiarTexto([cab.join('\\t')].concat(lin).join('\\n'), _copiado2);
}}

if(location.hash === '#dados') setTimeout(() => mostrarAba('dados'), 0);

function copiarTituloLinhaNf(indice){{
  const v = lotacao.veiculos_agendado_nf[indice];
  if(!v || !v.titulo_roteirizacao) return;
  _copiarTexto(v.titulo_roteirizacao, _mostrarCopiado);
}}
</script>
</body>
</html>
"""


def gerar_html(dados: dict, logo_data_uri: str, info_base: dict) -> str:
    logo_tag = f'<img src="{logo_data_uri}" alt="Logo">' if logo_data_uri else ""
    base_antiga = info_base["horas_desde_salva"] > ALERTA_BASE_ANTIGA_HORAS
    return TEMPLATE_HTML.format(
        gerado_em=datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        logo_tag=logo_tag,
        arquivo_base=html_lib.escape(info_base["nome"]),
        base_modificada=info_base["modificada_em"],
        qtd_linhas_base=info_base["qtd_linhas"],
        classe_base_antiga=" antiga" if base_antiga else "",
        aviso_base_antiga=(
            f" &middot; atenção: base salva há mais de {ALERTA_BASE_ANTIGA_HORAS}h, exporte de novo do Sankhya"
            if base_antiga else ""
        ),
        total_universo_valor=dados["total_universo_valor"],
        total_universo_peso=dados["total_universo_peso"],
        lotacao_json=json.dumps(dados["lotacao"], ensure_ascii=False),
        pedidos_json=json.dumps(dados.get("pedidos", []), ensure_ascii=False),
        fracionado_por_operador_json=json.dumps(dados["fracionado_por_operador"], ensure_ascii=False),
        km_por_dia=KM_POR_DIA,
        dias_expedicao=DIAS_EXPEDICAO,
        fator_estrada=str(FATOR_ESTRADA).replace(".", ","),
    )


def main():
    caminho_base = localizar_base_mais_recente(PASTA_SAIDA)
    mtime = os.path.getmtime(caminho_base)
    modificada_em = datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M")
    horas_desde_salva = (datetime.now().timestamp() - mtime) / 3600

    print(f"Lendo base: {caminho_base}")
    print(f"  Salva em:  {modificada_em}")
    if horas_desde_salva > ALERTA_BASE_ANTIGA_HORAS:
        print(
            f"  ATENÇÃO: essa base foi salva há {horas_desde_salva:.0f}h. "
            "Se você acabou de exportar do Sankhya, o arquivo novo não está nesta pasta:\n"
            f"    {PASTA_SAIDA}"
        )

    df = carregar_base(caminho_base)
    dados = montar_dados(df)
    logo_data_uri = carregar_logo_base64(CAMINHO_LOGO)
    info_base = {
        "nome": os.path.basename(caminho_base),
        "modificada_em": modificada_em,
        "horas_desde_salva": horas_desde_salva,
        "qtd_linhas": int(df[COL_COD_PARCEIRO].str.strip().ne("").sum()),
    }
    html = gerar_html(dados, logo_data_uri, info_base)

    salvar_html(ARQUIVO_SAIDA, html)

    lot = dados["lotacao"]
    print(f"\nRelatório gerado em: {ARQUIVO_SAIDA}")
    print(f"  Linhas com parceiro na base: {info_base['qtd_linhas']}")
    print(f"  Linhas no universo:          {dados['qtd_linhas_universo']}")
    print(f"  Veículos de lotação:         {lot['qtd_veiculos']}")
    print(f"  Agendado com NF (sem data):  {len(lot['veiculos_agendado_nf'])}")
    print(f"  Operadores no fracionado:    {len(dados['fracionado_por_operador'])}")
    com_prazo = sorted([v for v in lot["veiculos"] + lot["veiculos_agendado_nf"] if v.get("km")],
                       key=lambda v: v["saida_ordem"])
    if com_prazo:
        print("\n  PRAZOS (liberar até | carregar até | coletar até | agenda | cidade | km):")
        for v in com_prazo:
            print(f"    {v.get('liberar_ate',''):>11} | {v['expedicao_ate']:>11} | {v['saida_uberaba']:>11} | {v['agenda']:>11} | "
                  f"{v['cidade_uf'][:28]:28} | {v['km']:>5} km | {v['status_prazo']}")

    # abre SEMPRE o HTML que acabou de ser gravado — assim você nunca fica
    # olhando uma cópia antiga do relatório que esteja em outra pasta
    if ABRIR_NO_NAVEGADOR:
        abrir_no_navegador(ARQUIVO_SAIDA)
    else:
        print(f"\nPronto! Abra o relatório em: {ARQUIVO_SAIDA}")


if __name__ == "__main__":
    main()