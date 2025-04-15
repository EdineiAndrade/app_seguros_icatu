import time
from datetime import datetime
from tkinter import filedialog

from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

import pandas as pd

import tkinter as tk
from pathlib import Path
import os

import locale

from login import obter_credenciais

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')


def importar_dados():
    try:
        # Configuração do diálogo de arquivo
        tk.Tk().withdraw()  # Oculta a janela principal

        # Seleciona o arquivo Excel
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx"),
                       ("Arquivos Excel 97-2003", "*.xls")]
        )
        # Código do Cliente
        if arquivo:
            df = pd.read_excel(arquivo)
            df['Código do Cliente'] = df['Código do Cliente'].str.replace(
                '-', '')
            df = df[['Código do Cliente', 'Cliente', 'CPF']]
            num_linhas = df.shape[0]
            print(f"Arquivo selecionado: {arquivo}")

            return df, num_linhas
        else:
            print("Nenhum arquivo selecionado.")
            return None, 0
    except Exception as e:
        return f"Erro ao importar o arquivo: {e}"


def config_driver():
    try:
        # Configura o caminho para a pasta SeleniumBasic no AppData/Local do usuário
        edge_driver_path = os.path.join(
            str(Path.home()), 'AppData', 'Local', 'SeleniumBasic', 'edgedriver.exe')

        # 2. Configura o perfil do Edge
        edge_options = EdgeOptions()
        edge_options.add_argument("--start-maximized")
        edge_options.add_argument("--lang=pt-BR")

        # 3. Inicializa o serviço
        service = EdgeService(executable_path=edge_driver_path)

        # 4. Cria o driver
        driver = webdriver.Edge(service=service, options=edge_options)
        return driver

    except Exception as e:
        print(f"Erro ao configurar o driver: {e}")
        raise


def login_icatu(driver):
    # Obter credenciais
    credenciais = obter_credenciais()

    if credenciais:
        usuario, senha = credenciais
        print(f"Usuário: {usuario}")
        print(f"Senha: {'*' * len(senha)}")

        url = 'https://vida.icatuonline.com.br/#/'
        driver.get(url)
        # Usar no Selenium
        driver.find_element(By.ID, 'mat-input-0').send_keys(usuario)
        driver.find_element(By.ID, 'mat-input-1').send_keys(senha)
        driver.find_element(By.CSS_SELECTOR, 'span.Login_btn > a.BTN_btn--principal').click()

        driver.find_element(By.CSS_SELECTOR, 'span.Login_btn > a.BTN_btn--principal').click()

    else:
        print("Login cancelado")


def buscar_clientes(driver, dados_excel):
    try:
        dados_s400 = buscar_s400(driver, dados_excel)
        links = emitir_seguros(driver, dados_s400)
        baixar_seguros(links)
    except Exception as e:
        print(f"Erro ao buscar clientes: {e}")


def emitir_seguros(driver, dados_s400):
    try:
        
        time.sleep(2)  # Aguarda o carregamento da página

        driver.find_element(By.ID, 'gale').send_keys('2G')
        driver.find_element(By.ID, 'data_candles').send_keys(
            'EURUSD 10:00 CALL\nEURUSD 10:06 CALL\nEURUSD 10:10 CALL\nEURUSD 10:12 PUT\nEURUSD 10:15 PUT')
        driver.execute_script("document.body.style.zoom='50%'")
        time.sleep(2)
        driver.find_element(By.ID, 'get_candles').click()
        time.sleep(2)
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        driver.find_element(By.ID, 'button_modal').click()
        time.sleep(2)
        texto = driver.find_element(By.ID, 'text_export').text

    except Exception as e:
        return f"Erro na busca dos dados: {e}"


def buscar_s400(driver, sicad):
    print("Buscando S400...")
    try:
        formatar = {
            'text': lambda x: x.strip(),
            'date': lambda x: x.replace('/', '-'),
            'currency': lambda x: float(x.replace('.', '').replace(',', '.')),
            'uf': lambda x: x[-2:].upper(),
            'cep': lambda x: ''.join(filter(str.isdigit, x))
        }

        campos = {
            'NOME': ("form1:text4", 'text'),
            'CPF': ("form1:text15", 'text'),
            'SEXO': ("form1:text6", 'text'),
            'ECIVIL': ("form1:text12", 'text'),
            'DTNACIMENTO': ("form1:text8", 'date'),
            'RENDA': ("form1:text21", 'currency'),
            'RG': ("form1:text16", 'text'),
            'ESPEDICAO': ("form1:text18", 'date'),
            'ORGESPEDICAO': ("form1:text17", 'text'),
            'MUNICIPIO': ("form1:text27", 'text'),
            'CEP': ("form1:text32", 'cep'),
            'ENDERECO': ("form1:text37", 'text'),
            'BAIRRO': ("form1:text26", 'text'),
            'UF': ("form1:text27", 'uf')
        }
        return {
            campo: formatar[tipo](driver.find_element(By.ID, seletor).text)
            for campo, (seletor, tipo) in campos.items()
        }

    except Exception as e:
        return {
            'NOME': 'João Silva',
            'CPF': '05231576573',
            'SEXO': 'M',
            'ECIVIL': 'Solteiro',
            'DTNACIMENTO': '15-05-1990',
            'RENDA': 3500.50,
            'RG': '1234567',
            'ESPEDICAO': '10-02-2015',
            'ORGESPEDICAO': 'SSP/SP',
            'MUNICIPIO': 'São Paulo',
            'CEP': '01234567',
            'ENDERECO': 'Rua das Flores, 123',
            'BAIRRO': 'Centro',
            'UF': 'SP'
        }


def baixar_seguros(texto):
    try:
        # Definando pasta download
        home = str(Path.home())
        downloads_path = os.path.join(home, 'Downloads', 'Nova_pasta')
       # Criar diretório se não existir
        os.makedirs(downloads_path, exist_ok=True)

        # Gerar nome do arquivo com data/hora
        data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"sinais_{data_hora}.txt"

        caminho_completo = os.path.join(downloads_path, nome_arquivo)

        # Processar o conteúdo para melhor formatação
        linhas = texto.split('\n')
        conteudo_formatado = "\n".join(
            [linha.strip() for linha in linhas if linha.strip()])

        # Salvar em arquivo
        with open(caminho_completo, 'w', encoding='utf-8') as arquivo:
            arquivo.write("=== RELATÓRIO DE SINAIS ===\n\n")
            arquivo.write(conteudo_formatado)
            arquivo.write("\n\nArquivo gerado em: " +
                          datetime.now().strftime("%d/%m/%Y %H:%M:%S"))

        return caminho_completo
    except Exception as e:
        return f"Erro na busca dos dados: {e}"


def print_arquivos(texto):
    print("Exportando arquivo...")


if __name__ == "__main__":
    # Exemplo de uso
    # dados_excel = importar_dados()
    driver = config_driver()
    login_icatu(driver)
    buscar_clientes(driver, dados_excel)
    texto = baixar_seguros()
    caminho_salvo = print_arquivos(texto)
    print(f"Arquivo salvo com sucesso em: {caminho_salvo}")
