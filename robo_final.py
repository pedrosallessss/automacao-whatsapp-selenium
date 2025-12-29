import pandas as pd
from datetime import datetime
import os
import time
import re
import random
from io import BytesIO

# Bibliotecas de Imagem e Clipboard
from PIL import Image
import win32clipboard

# Bibliotecas do Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains # <--- Nova importação vital
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def limpar_telefone(numero):
    if pd.isna(numero): return None
    limpo = re.sub(r'\D', '', str(numero))
    if len(limpo) < 8: return None
    if len(limpo) in [10, 11]: limpo = "55" + limpo
    return limpo

def pegar_imagem_aleatoria():
    pasta_img = os.path.join(os.getcwd(), "Lembretes")
    if not os.path.exists(pasta_img): return None
    arquivos = [f for f in os.listdir(pasta_img) if f.endswith(('.png', '.jpg', '.jpeg'))]
    if not arquivos: return None
    return os.path.join(pasta_img, random.choice(arquivos))

def copiar_imagem_para_clipboard(caminho_imagem):
    try:
        image = Image.open(caminho_imagem)
        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        win32clipboard.CloseClipboard()
        return True
    except Exception as e:
        print(f"❌ Erro ao copiar imagem: {e}")
        return False

def automacao_selenium():
    # ================= CONFIGURAÇÃO =================
    NOME_DO_ARQUIVO = 'dados_monitoria.xls' 
    COL_ALUNO = 'Aluno'              
    COL_TELEFONE = 'Telefone Aluno'         
    COL_DIA = 'Dias Agendamento'     
    COL_HORA = 'Horas Agendamento'   
    # ================================================

    print("=========================================")
    print("🤖 ROBÔ DE MONITORIA - CORREÇÃO DE ENVIO")
    print("=========================================")

    # --- 1. LEITURA ---
    if not os.path.exists(NOME_DO_ARQUIVO):
        print(f"❌ Arquivo não encontrado.")
        return

    try:
        df = pd.read_excel(NOME_DO_ARQUIVO)
        df.columns = [str(c).strip() for c in df.columns]
        if COL_TELEFONE not in df.columns:
            print(f"❌ Coluna '{COL_TELEFONE}' não encontrada.")
            return
    except Exception as e:
        print(f"❌ Erro excel: {e}")
        return

    dias_map = {0:'Segunda-Feira', 1:'Terça-Feira', 2:'Quarta-Feira', 3:'Quinta-Feira', 4:'Sexta-Feira', 5:'Sábado', 6:'Domingo'}
    hoje = dias_map[datetime.now().weekday()]
    
    print(f"📅 Dia: {hoje}")
    hora_alvo = input("⏰ Horário (ex: 14:00): ")

    df_filtrado = df[
        (df[COL_DIA].astype(str).str.contains(hoje, na=False, regex=False)) &
        (df[COL_HORA].astype(str).str.contains(hora_alvo, na=False, regex=False))
    ]

    if df_filtrado.empty:
        print(f"⚠️ Ninguém encontrado.")
        return

    # --- 2. SELEÇÃO ---
    print("\n📋 LISTA DE ENVIO:")
    lista_envio = []
    lista_temp = []
    for i, (_, row) in enumerate(df_filtrado.iterrows()):
        nome = row[COL_ALUNO]
        fone = limpar_telefone(row[COL_TELEFONE])
        status = "✅ OK" if fone else "❌ Sem Zap"
        print(f"[{i}] | {status:<10} | {nome}")
        lista_temp.append({'nome': nome, 'fone': fone})

    print("\n[ENTER]=Todos | [0,2]=Específicos | [N]=Cancelar")
    escolha = input("Escolha: ")
    if escolha.lower() == 'n': return
    elif escolha.strip() == "":
        lista_envio = [c for c in lista_temp if c['fone']]
    else:
        try:
            indices = [int(x.strip()) for x in escolha.split(',')]
            for idx in indices:
                if 0 <= idx < len(lista_temp) and lista_temp[idx]['fone']:
                    lista_envio.append(lista_temp[idx])
        except: return

    if not lista_envio: return

    # --- 3. NAVEGADOR ---
    print(f"\n🚀 Iniciando navegador...")
    dir_path = os.getcwd()
    profile_path = os.path.join(dir_path, "Perfil_Zap")
    options = Options()
    options.add_argument(f"user-data-dir={profile_path}")
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.get("https://web.whatsapp.com")
    
    # Cria o controlador de ações (teclado virtual)
    actions = ActionChains(driver)

    print("Aguardando WhatsApp...")
    try:
        WebDriverWait(driver, 120).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
        )
        print("✅ Conectado!")
        time.sleep(2)
    except:
        print("⚠️ Verifique o login.")

    # --- 4. ENVIO ---
    for i, aluno in enumerate(lista_envio):
        primeiro_nome = str(aluno['nome']).split()[0].title()
        numero = aluno['fone']
        caminho_imagem = pegar_imagem_aleatoria()
        
        msg = f"Olá {primeiro_nome}! Passando aqui para lembrar que sua aula começa às {hora_alvo}. Te aguardamos aqui! 🖥️"
        
        print(f"[{i+1}/{len(lista_envio)}] Enviando para {primeiro_nome}...")
        
        # 1. Copia Imagem
        if caminho_imagem:
            copiar_imagem_para_clipboard(caminho_imagem)
        
        # 2. Abre Conversa
        link = f"https://web.whatsapp.com/send?phone={numero}"
        driver.get(link)
        
        try:
            # 3. Foca na caixa de texto principal
            caixa_texto = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
            )
            time.sleep(3) # Espera carregar bem
            
            # 4. COLA A IMAGEM (Ctrl + V)
            caixa_texto.send_keys(Keys.CONTROL, 'v')
            
            # 5. ESPERA A TELA DE PRÉ-VISUALIZAÇÃO ABRIR (CRUCIAL)
            print("   📸 Imagem colada, aguardando prévia...")
            time.sleep(3) 
            
            # 6. DIGITA A LEGENDA (MÉTODO NOVO: ACTION CHAINS)
            # Como acabamos de colar, o foco JÁ ESTÁ na legenda. Não precisamos procurar o elemento.
            # Apenas mandamos o robô digitar as teclas.
            actions.send_keys(msg).perform()
            time.sleep(1)
            actions.send_keys(Keys.ENTER).perform()
            
            # 7. SEGURANÇA: Se o Enter não enviou, clica no botão verde manualmente
            try:
                time.sleep(2)
                botao_enviar = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
                botao_enviar.click()
                print("   (Botão de envio clicado por segurança)")
            except:
                pass # Se já enviou com o Enter, o botão não vai existir, então segue a vida

            print(f"   ✅ Enviado!")
            time.sleep(5) 

        except Exception as e:
            print(f"   ❌ Falha: {e}")

    print("\n🏁 FINALIZADO! Fechando em 5s...")
    time.sleep(5)
    driver.quit()

if __name__ == "__main__":
    automacao_selenium()