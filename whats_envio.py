"""
Envio de mensagens pelo WhatsApp Web, num Chrome SEPARADO só do robô
(perfil em C:\\chrome-whats). Na 1ª vez você lê o QR Code; depois ele entra sozinho.

Segurança:
  - antes de enviar, confere se o NOME no topo da conversa é o contato certo;
    se não bater, NÃO envia e avisa.
"""
import re
import time
import unicodedata

PERFIL_WHATS = r"C:\chrome-whats"

# Seletores do WhatsApp Web (com alternativas, porque o WhatsApp muda o site às vezes)
SEL_LOGADO = ["#pane-side", "div[aria-label='Lista de conversas']", "div[aria-label='Chat list']"]
SEL_BUSCA = ["div[contenteditable='true'][data-tab='3']",
             "div[role='textbox'][aria-label*='esquis']",
             "input[aria-label*='esquis']", "input[placeholder*='esquis']",
             "#side div[contenteditable='true']", "#side input[type='text']",
             "div[role='textbox'][aria-label*='earch']", "input[aria-label*='earch']"]
SEL_BOTAO_BUSCA = ["#side button[aria-label*='esquis']", "#side span[data-icon='search']",
                   "button[aria-label*='Search']"]
# telas que travam o WhatsApp Web depois de muito tempo aberto (o motivo de ter que fechar e abrir)
TXT_USAR_AQUI = ("usar aqui", "use here")
TXT_TRAVADO = ("whatsapp esta aberto em outra janela", "whatsapp is open in another window",
               "telefone sem conexao", "computador nao conectado", "phone not connected",
               "tentando conectar", "trying to reach", "carregando suas conversas")
SEL_CAIXA = ["footer div[contenteditable='true'][data-tab='10']",
             "footer div[contenteditable='true'][role='textbox']",
             "footer div[contenteditable='true']"]
SEL_CABECALHO = ["#main header", "header[data-testid='conversation-header']"]


def _norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return " ".join(s.lower().split())


class WhatsApp:
    def __init__(self, log=print):
        self.driver = None
        self.log = log

    # ---------------------------------------------------------------- abrir
    def _achar(self, seletores, timeout=0):
        from selenium.webdriver.common.by import By
        fim = time.time() + timeout
        while True:
            for css in seletores:
                for e in self.driver.find_elements(By.CSS_SELECTOR, css):
                    try:
                        if e.is_displayed():
                            return e
                    except Exception:
                        pass
            if time.time() >= fim:
                return None
            time.sleep(0.5)

    def aberto(self):
        try:
            return self.driver is not None and bool(self.driver.window_handles)
        except Exception:
            return False

    def abrir(self, espera_login=180):
        """Abre o WhatsApp Web e espera estar logado (QR Code na 1ª vez)."""
        if not self.aberto():
            from selenium import webdriver
            opts = webdriver.ChromeOptions()
            opts.add_argument(f"--user-data-dir={PERFIL_WHATS}")
            opts.add_argument("--start-maximized")
            self.driver = webdriver.Chrome(options=opts)
            self.driver.get("https://web.whatsapp.com")
        if self._achar(SEL_LOGADO, timeout=15):
            return True
        self.log("WhatsApp Web aberto: se aparecer o QR Code, leia com o celular "
                 "(WhatsApp > Aparelhos conectados > Conectar aparelho).")
        return self._achar(SEL_LOGADO, timeout=espera_login) is not None

    def fechar(self):
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass
        self.driver = None

    # ---------------------------------------------------------------- conversa
    def _tela(self):
        try:
            return _norm(self.driver.execute_script("return document.body ? document.body.innerText.slice(0, 3000) : ''"))
        except Exception:
            return ""

    def _destravar(self):
        """O que você fazia na mão: se aparecer 'Usar aqui', clica; senão recarrega a página e espera entrar."""
        from selenium.webdriver.common.by import By
        try:
            for b in self.driver.find_elements(By.XPATH, "//button|//div[@role='button']"):
                if _norm(b.text) in TXT_USAR_AQUI and b.is_displayed():
                    b.click()
                    self.log("WhatsApp Web: cliquei em 'Usar aqui'.")
                    time.sleep(4)
                    if self._achar(SEL_LOGADO, timeout=30):
                        return True
        except Exception:
            pass
        self.log("WhatsApp Web travado — recarregando a página sozinho…")
        try:
            self.driver.get("https://web.whatsapp.com")
        except Exception:
            try:
                self.driver.refresh()
            except Exception:
                return False
        return self._achar(SEL_LOGADO, timeout=90) is not None

    def _tentar_abrir(self, nome):
        from selenium.webdriver.common.by import By
        from selenium.webdriver.common.keys import Keys
        try:  # fecha painel/pesquisa que tenha ficado aberto
            self.driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        except Exception:
            pass
        busca = self._achar(SEL_BUSCA, timeout=6)
        if not busca:
            bt = self._achar(SEL_BOTAO_BUSCA, timeout=2)
            if bt:
                try:
                    bt.click()
                except Exception:
                    pass
                busca = self._achar(SEL_BUSCA, timeout=4)
        if not busca:
            return False
        busca.click()
        busca.send_keys(Keys.CONTROL, "a")
        busca.send_keys(Keys.BACKSPACE)
        busca.send_keys(nome)
        time.sleep(2.5)
        # clica no resultado com o nome exato; se não achar, Enter no 1º (conferido depois)
        alvo = None
        try:
            for e in self.driver.find_elements(By.CSS_SELECTOR, "#pane-side span[title], #side span[title]"):
                if _norm(e.get_attribute("title")) == _norm(nome) and e.is_displayed():
                    alvo = e
                    break
        except Exception:
            alvo = None
        if alvo is not None:
            try:
                alvo.click()
            except Exception:
                busca.send_keys(Keys.ENTER)
        else:
            busca.send_keys(Keys.ENTER)
        return self._achar(SEL_CAIXA, timeout=15) is not None

    def _abrir_conversa(self, nome, telefone):
        if telefone:
            self.driver.get(f"https://web.whatsapp.com/send?phone={telefone}")
            return self._achar(SEL_CAIXA, timeout=60) is not None
        if self._tentar_abrir(nome):
            return True
        # 2ª tentativa: destrava (Usar aqui / recarrega) e tenta de novo — antes você fazia isso na mão
        tela = self._tela()
        motivo = next((t for t in TXT_TRAVADO if t in tela), "")
        self.log(f"Não abri a conversa de '{nome}' de primeira" + (f" ({motivo})" if motivo else "") + " — destravando…")
        if self._destravar() and self._tentar_abrir(nome):
            self.log(f"WhatsApp Web destravado, conversa de '{nome}' aberta.")
            return True
        try:
            self.driver.save_screenshot("whats_erro.png")
            with open("whats_erro.txt", "w", encoding="utf-8") as f:
                f.write(time.strftime("%d/%m/%Y %H:%M:%S") + f" - não abri a conversa de {nome}. Tela: {tela[:300]}")
        except Exception:
            pass
        return False

    def _conversa_e_de(self, nome):
        cab = self._achar(SEL_CABECALHO, timeout=5)
        if not cab:
            return False, "(não li o topo da conversa)"
        texto = cab.text.split("\n")[0]
        return _norm(nome) in _norm(cab.text), texto

    def enviar(self, nome, texto, telefone=""):
        """Envia `texto` (pode ter várias linhas) para o contato. Retorna (ok, detalhe)."""
        from selenium.webdriver.common.keys import Keys
        if not self.abrir():
            return False, "WhatsApp Web não está logado (leia o QR Code)."
        if not self._abrir_conversa(nome, telefone):
            return False, f"não consegui abrir a conversa de '{nome}'."
        ok, topo = self._conversa_e_de(nome)
        if not ok:
            return False, f"a conversa aberta é '{topo}', não '{nome}'. Nada foi enviado."
        caixa = self._achar(SEL_CAIXA, timeout=10)
        caixa.click()
        caixa.send_keys(Keys.CONTROL, "a")
        caixa.send_keys(Keys.BACKSPACE)
        linhas = texto.split("\n")
        for i, linha in enumerate(linhas):
            if linha:
                caixa.send_keys(linha)
            if i < len(linhas) - 1:
                caixa.send_keys(Keys.SHIFT, Keys.ENTER)  # quebra de linha sem enviar
        time.sleep(0.5)
        caixa.send_keys(Keys.ENTER)
        time.sleep(2)
        caixa = self._achar(SEL_CAIXA, timeout=5)
        if caixa is not None and caixa.text.strip():
            return False, "a mensagem ficou na caixa de texto e não saiu. Confira na tela."
        return True, f"enviado para {topo}"

    # ---------------------------------------------------------------- leitura (respostas)
    def com_nao_lidas(self, nomes):
        """Dos nomes, quais aparecem na lista de conversas com bolinha de não lida."""
        if not self.aberto():
            return []
        try:
            achados = self.driver.execute_script("""
                var nomes = arguments[0].map(n => n.toLowerCase());
                var linhas = Array.from(document.querySelectorAll('#pane-side [role="listitem"], #pane-side [role="row"]'));
                var r = [];
                linhas.forEach(l => {
                  var txt = (l.innerText || '').toLowerCase();
                  var badge = l.querySelector('[aria-label*="não lida"],[aria-label*="nao lida"],[aria-label*="unread"]');
                  if (!badge) return;
                  nomes.forEach((n, i) => { if (txt.startsWith(n) || txt.split('\\n')[0].includes(n)) r.push(arguments[0][i]); });
                });
                return r;""", list(nomes)) or []
            return list(dict.fromkeys(achados))
        except Exception:
            return []

    def mensagens_recebidas(self, nome, telefone=""):
        """Abre a conversa e devolve as mensagens RECEBIDAS visíveis:
        [{id, pre, texto}] (pre = '[14:32, 07/10/2026] Contato Alfa: ')."""
        if not self.abrir(espera_login=5):
            return None
        if not self._abrir_conversa(nome, telefone):
            return None
        ok, _ = self._conversa_e_de(nome)
        if not ok:
            return None
        time.sleep(1.5)
        return self.driver.execute_script("""
            return Array.from(document.querySelectorAll('#main .message-in [data-pre-plain-text]')).map(e => {
              var pre = e.getAttribute('data-pre-plain-text') || '';
              var t = e.querySelector('.selectable-text');
              var txt = t ? t.innerText : (e.innerText || '');
              return {id: pre + '|' + txt.slice(0, 80), pre: pre, texto: txt};
            });""") or []

    # ---------------------------------------------------------------- nomes das conversas
    def listar_conversas(self, rolagens=80):
        """Rola a lista de conversas e devolve só os NOMES (não abre nenhuma conversa)."""
        if not self.abrir(espera_login=5):
            return []
        nomes = []
        try:
            self.driver.execute_script("var p=document.querySelector('#pane-side'); if(p) p.scrollTop=0;")
            time.sleep(1)
            parado = 0
            for _ in range(rolagens):
                lote = self.driver.execute_script("""
                    return Array.from(document.querySelectorAll('#pane-side [role="listitem"] span[title], #pane-side [role="row"] span[title]'))
                      .map(e => e.getAttribute('title')).filter(t => t);""") or []
                antes = len(nomes)
                for t in lote:
                    if t not in nomes:
                        nomes.append(t)
                fim = self.driver.execute_script("""
                    var p=document.querySelector('#pane-side'); if(!p) return true;
                    var a=p.scrollTop; p.scrollTop=a+p.clientHeight*0.9; return p.scrollTop===a;""")
                time.sleep(0.7)
                parado = parado + 1 if len(nomes) == antes else 0
                if fim and parado >= 2:
                    break
            self.driver.execute_script("var p=document.querySelector('#pane-side'); if(p) p.scrollTop=0;")
        except Exception as e:
            self.log(f"Não consegui ler a lista de conversas: {e}")
        # tira linhas que não são nome (ex.: prévia de mensagem com ':' ou horas)
        return [n for n in nomes if len(n) <= 60 and not re.match(r"^\d{1,2}:\d{2}$", n)]
