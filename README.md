# Roteirização e Contratação de Frete

> ⚠️ **Dados fictícios.** Este repositório é uma versão de portfólio de uma ferramenta usada no dia a dia
> de um CD de logística. Nomes de empresas, transportadoras, clientes e pessoas foram trocados por
> exemplos (ALFA, BETA, CLIENTE FICTICIO…). Nenhuma planilha real está incluída.

## O que faz
Lê a base de pedidos exportada do ERP (Sankhya), separa **lotação x fracionado**, monta os veículos (truck / carreta / rodotrem), calcula **até quando a carga precisa sair** para chegar na data agendada e gera um relatório HTML. Um robô com janela (Tkinter) vigia a pasta e, com um clique de confirmação, envia as cargas para cada transportadora pelo **WhatsApp Web** (Selenium).

## Problema que resolve
Antes: montar as cargas e digitar mensagem por mensagem para cada transportadora, à mão, todo dia. Agora: exportar a base → o relatório e as mensagens saem prontos; a pessoa só confere e clica em ENVIAR.

**Uso real:** Em uso diário desde 05/10/2026: **286 envios registrados** no log até 12/11/2026.

## Destaques técnicos
- **Prazo de saída automático**: km = linha reta × 1,25 (coordenadas IBGE) ÷ 500 km/dia → `coletar até`, `carregar até`, `liberar até`, pulando domingo.
- **Regras de negócio configuráveis**: peso de fracionado, capacidade de truck/carreta, capacidade de O.C. por dia.
- **Humano no controle**: nada é enviado sem clique; antes de enviar confere se o nome no topo da conversa é o contato certo.
- **Memória**: aprende UF → transportadora e não repete título já enviado.

## Tecnologias
Python · pandas · openpyxl · Selenium (WhatsApp Web) · Tkinter · HTML/CSS/JS gerado · geodistância (haversine)

## Arquivos
| Arquivo | Função |
|---|---|
| `gerar_relatorio_roteirizacao.py` | lê a base e gera `relatorio_roteirizacao.html` |
| `robo_contratacao.py` | janela que vigia a pasta e envia as cargas |
| `whats_envio.py` | envio pelo WhatsApp Web num Chrome separado |
| `historico_contratacao.py` | gera planilha de histórico dos envios |
| `transp_respostas.py` | lê as respostas das transportadoras |
| `cidades_coordenadas.csv` | coordenadas dos municípios (IBGE, dado público) |
| `exemplos/` | base fictícia + relatório gerado a partir dela |

## Como rodar
```bash
pip install -r requirements.txt
copy "exemplos\Base Roteirizacao (exemplo).xlsx" .
python gerar_relatorio_roteirizacao.py      # gera relatorio_roteirizacao.html
python robo_contratacao.py                  # (opcional) robô com janela + WhatsApp
```

## Autor
Cristiano Braga — [github.com/CristianoBraga88](https://github.com/CristianoBraga88)
