import streamlit.web.cli as stcli
import sys
import os
import threading
import time
import socket
import signal
import webview
import urllib.request

# --- CONFIGURAÇÕES DE ATUALIZAÇÃO ---
GITHUB_URL = "https://raw.githubusercontent.com/Agrivalle-BioDigital/Gest-o-de-Projetos/refs/heads/main/dashboard.py"

def obter_diretorio_base():
    """Garante que o arquivo seja salvo/lido na mesma pasta do .exe, e não na pasta temporária"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

LOCAL_APP = os.path.join(obter_diretorio_base(), "dashboard.py")

def atualizar_dashboard():
    """Baixa silenciosamente a última versão do código do GitHub"""
    print("Verificando atualizações no GitHub...")
    try:
        urllib.request.urlretrieve(GITHUB_URL, LOCAL_APP)
        print("Código atualizado com sucesso!")
    except Exception as e:
        print(f"Aviso: Não foi possível atualizar (sem internet?). Usando versão local. Erro: {e}")

# --- LÓGICA DO STREAMLIT E WEBVIEW ---
def start_streamlit():
    """Inicia o Streamlit em uma thread separada com correção de sinais"""
    sys.argv = [
        "streamlit",
        "run",
        LOCAL_APP,
        "--global.developmentMode=false",
        "--server.headless=true",
        "--server.port=8501",
        "--server.address=localhost",
    ]

    os.environ['STREAMLIT_SERVER_HEADLESS'] = 'true'
    original_signal = signal.signal
    signal.signal = lambda x, y: None
    
    try:
        stcli.main()
    except SystemExit:
        pass
    finally:
        signal.signal = original_signal

def wait_for_server(port=8501):
    """Aguarda o servidor estar online para evitar tela branca"""
    retries = 0
    while retries < 30: # Tenta por 15 segundos
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(1)
                if s.connect_ex(('localhost', port)) == 0:
                    return True
        except:
            pass
        time.sleep(0.5)
        retries += 1
    return False

if __name__ == "__main__":
    # 1. Tenta atualizar o arquivo dashboard.py primeiro
    atualizar_dashboard()
    
    # Verifica se o arquivo existe (caso seja a primeira vez abrindo sem internet)
    if not os.path.exists(LOCAL_APP):
        webview.create_window("Erro Fatal", html="<h1>Erro: Arquivo do sistema não encontrado e sem internet para baixar.</h1>")
        webview.start()
        sys.exit(1)

    # 2. Inicia o Streamlit
    t = threading.Thread(target=start_streamlit)
    t.daemon = True
    t.start()

    # 3. Abre a Janela do Aplicativo
    if wait_for_server(8501):
        webview.create_window("Gestão de Projetos", "http://localhost:8501")
        webview.start()
    else:
        webview.create_window("Erro", html="<h1>Erro: O servidor Streamlit não iniciou a tempo.</h1>")
        webview.start()