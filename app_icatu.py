import time
from datetime import datetime
from tkinter import filedialog

from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

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

            print(f"Arquivo selecionado: {arquivo}")

            return df
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


def login_icatu(driver, credenciais):
    # Obter credenciais

    if credenciais:
        usuario, senha = credenciais
        print(f"Usuário: {usuario}")
        print(f"Senha: {'*' * len(senha)}")

        url = 'https://vida.icatuonline.com.br/#/'
        driver.get(url)
        # Usar no Selenium
        driver.find_element(By.ID, 'mat-input-0').send_keys(usuario)
        driver.find_element(By.ID, 'mat-input-1').send_keys(senha)
        driver.find_element(
            By.CSS_SELECTOR, 'span.Login_btn > a.BTN_btn--principal').click()

        timeout = 10  # Define a timeout value in seconds
        botao = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button.btn-cancel.mat-button"))
        )
        botao.click()

    else:
        print("Login cancelado")


def buscar_clientes(driver, dados_excel):
    try:
        # Verifica se dados_excel é um DataFrame e se tem dados
        if dados_excel.empty:
            raise ValueError(
                "Dados de entrada não são um DataFrame válido ou estão vazios")
        # Verifica se a coluna necessária existe
        # Substitua pelo nome da sua coluna de iteração (ex: 'CNPJ', 'ID', etc.)
        num_linhas = dados_excel.shape[0]
        print(f"Número de linhas a serem processadas: {num_linhas}")
        coluna_chave = 'Código do Cliente'
        if coluna_chave not in dados_excel.columns:
            raise KeyError(
                f"Coluna '{coluna_chave}' não encontrada no DataFrame")

        # Lista para armazenar resultados
        resultados = []

        # Itera sobre cada linha do DataFrame
        for index, linha in dados_excel.iterrows():
            try:
                # Extrai o valor da coluna chave
                valor_chave = linha[coluna_chave]
                print(f"\nProcessando: {valor_chave}")

                # Busca dados no S400 (adaptar conforme sua função)
                # Passa a linha inteira ou só o valor necessário
                dados_s400 = buscar_s400(driver, linha)

                # Emite seguros (adaptar conforme sua função)
                links = emitir_seguros(driver, dados_s400)

                # Baixa documentos (se aplicável)
                if links:
                    baixar_seguros(links)

                # Adiciona resultado à lista
                resultados.append({
                    'chave': valor_chave,
                    'status': 'Sucesso',
                    'links': links
                })

            except Exception as e:
                print(f"Erro ao processar {valor_chave}: {str(e)}")
                resultados.append({
                    'chave': valor_chave,
                    'status': 'Erro',
                    'erro': str(e)
                })

        # Retorna resultados como DataFrame
        return pd.DataFrame(resultados)

    except Exception as e:
        print(f"Erro geral na execução: {str(e)}")
        raise  # Re-lança a exceção para tratamento externo


def emitir_seguros(driver, dados_s400):
    try:

        time.sleep(2)  # Aguarda o carregamento da página
        driver.execute_script("document.body.style.zoom='80%'")
        driver.find_element(
            By.XPATH, '/html/body/app-root/app-seguro-bnb/div[1]/div[2]/div[2]/div/div[8]/div/a').click()
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.ID, 'mat-input-10'))
        )
        driver.find_element(By.ID, 'mat-input-10').send_keys(dados_s400['CPF'])
        time.sleep(2)
        driver.find_element(
            By.ID, 'mat-select-3').send_keys(dados_s400['ECIVIL'])
        driver.find_element(By.ID, 'mat-input-15').send_keys('Agricultor')
        driver.find_element(By.XPATH, '//*[@id="mat-option-159"]/span').click()
        driver.find_element(By.ID, 'mat-input-16').clear()
        driver.find_element(
            By.ID, 'mat-input-16').send_keys(dados_s400['RENDA'])
        driver.find_element(By.ID, 'mat-input-17').clear()
        driver.find_element(By.ID, 'mat-input-17').send_keys(dados_s400['RG'])
        driver.find_element(By.ID, 'mat-input-18').clear()
        driver.find_element(
            By.ID, 'mat-input-18').send_keys(dados_s400['ESPEDICAO'])
        driver.find_element(By.ID, 'mat-input-19').clear()
        driver.find_element(
            By.ID, 'mat-input-19').send_keys(dados_s400['ORGESPEDICAO'])

        # TRATAR OD DADOS DE UF ANTES
        driver.find_element(By.ID, 'mat-select-5').send_keys(dados_s400['UF'])

        driver.find_element(By.ID, 'mat-input-20').clear()
        driver.find_element(
            By.ID, 'mat-input-20').send_keys(dados_s400['telefone'])

        driver.find_element(By.ID, 'mat-input-21').clear()
        driver.find_element(
            By.ID, 'mat-input-21').send_keys(dados_s400['email'])

        time.sleep(2)
        # ir para próxima página
        driver.find_element(
            By.XPATH, '//*[@id="mat-tab-content-0-0"]/div/form/div[12]/span[2]/a').click()

        # página de CEP
        driver.find_element(By.ID, 'mat-input-2').send_keys(dados_s400['CEP'])
        driver.find_element(
            By.ID, 'mat-input-3').send_keys(dados_s400['ENDERECO'])
        driver.find_element(
            By.ID, 'mat-input-4').send_keys(dados_s400['NUMERO'])
        driver.find_element(
            By.ID, 'mat-input-6').send_keys(dados_s400['BAIRRO'])
        time.sleep(2)
        driver.find_element(
            By.XPATH, '//*[@id="mat-tab-content-0-1"]/div/app-address-data/div[2]/span[2]/a').click()
        time.sleep(2)
        driver.find_element(
            By.XPATH, '//*[@id="mat-tab-content-0-2"]/div/app-payment-details/div[2]/span[2]/a').click()
        time.sleep(2)
        driver.find_element(
            By.XPATH, '//*[@id="mat-tab-content-0-3"]/div/app-signature/div/div[1]/div[2]/button[1]/span').click()
        time.sleep(2)
        driver.find_element(
            By.XPATH, '//*[@id="mat-dialog-1"]/app-modal-dialog-common/div/div[2]/button[2]').click()
        time.sleep(2)
        driver.find_element(
            By.XPATH, '//*[@id="mat-tab-content-0-4"]/div/app-conclusion/div[7]/span[2]/a').click()
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        driver.find_element(By.ID, 'button_modal').click()

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
            'CPF': '095.103.405-70',
            'ECIVIL': 'Solteiro',
            'RENDA': 3500.50,
            'RG': '1234567',
            'ESPEDICAO': '10-02-2015',
            'ORGESPEDICAO': 'SSP/SP',
            'MUNICIPIO': 'São Paulo',
            'CEP': '46650000',
            'ENDERECO': 'Rua das Flores, 123',
            'BAIRRO': 'Centro',
            'UF': 'Bahia',
            'telefone': '77999272367',
            'email': ''
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
    print("Iniciando o processo...")
    print("Importando dados do Excel...")
    dados_excel = importar_dados()
    print("Coletando dados de Usuário e Senha...")
    credenciais = obter_credenciais()
    print("Abrindo navegador Edge...")
    driver = config_driver()
    print("Fazendo login na Icatu...")
    login_icatu(driver, credenciais)
    print("Buscando clientes...")
    resultado = buscar_clientes(driver, dados_excel)
    print(f"Processo finalizado com sucesso!")
