import os  # Biblioteca para interagir com o sistema operacional, como manipular diretórios e arquivos.
import smtplib  # Biblioteca para enviar e-mails utilizando o protocolo SMTP.
import zipfile  # Biblioteca para criar e extrair arquivos ZIP.
from email.mime.multipart import MIMEMultipart  # Classe para criar mensagens de e-mail com várias partes.
from email.mime.base import MIMEBase  # Classe para criar partes MIME básicas para anexos de e-mail.
from email.mime.text import MIMEText # Classe para incorporar corpo do texto
from email import encoders  # Módulo para codificar e decodificar anexos MIME em e-mails.
from datetime import datetime  # Classe para trabalhar com datas e horários.
import xml.etree.ElementTree as ET  # Biblioteca para analisar e manipular documentos XML.
import tkinter as tk  # Biblioteca para criar interfaces gráficas.
from tkinter import messagebox  # Submódulo de tkinter para exibir caixas de diálogo de mensagem.
import tkcalendar  # Biblioteca para adicionar um widget de calendário à interface gráfica.

CONFIG_FILE = 'config.xml'  # Nome do arquivo de configuração que será lido.

def ler_configuracoes_smtp():
    global host, port, username, password, destinatario, empresas, assunto, corpo

    try:
        tree = ET.parse(CONFIG_FILE)
        root = tree.getroot()

        host = root.find('smtp').text
        port = int(root.find('porta').text)
        username = root.find('remetente').text
        password = root.find('senha').text
        destinatario = root.find('destinatario').text

        empresas = []
        for empresa_element in root.findall('empresa'):
            diretorio_xml = empresa_element.find('diretorio_xml').text
            diretorio_pdf = empresa_element.find('diretorio_pdf').text
            empresas.append((diretorio_xml, diretorio_pdf))
            
        mensagem = root.find('mensagem')
        assunto = mensagem.find('assunto').text
        corpo = mensagem.find('corpo').text

    except (IOError, ET.ParseError) as e:
        print(f'Erro ao ler o arquivo de configuração: {e}')
        raise

    
def compactar_arquivos(empresas, mes_desejado):
    try:
        arquivos_zip = []  # Lista para armazenar os caminhos dos arquivos ZIP criados

        # Obter o diretório do código
        diretorio_codigo = os.path.dirname(os.path.abspath(__file__))  # Obtém o diretório do código atual

        for empresa_numero, empresa in enumerate(empresas, 1):  # Enumera as empresas com base em seus índices (1, 2, 3, ...)
            diretorio_xml, diretorio_pdf = empresa  # Obtém os diretórios XML e PDF da empresa atual

            # Define o diretório de saída como o diretório do código
            diretorio_saida = diretorio_codigo  # O diretório de saída dos arquivos compactados será o mesmo do código

            # Verifica se existem arquivos XML e PDF com o mês desejado nos diretórios correspondentes
            arquivos_xml_existem = verificar_arquivos_existem(mes_desejado, diretorio_xml)
            arquivos_pdf_existem = verificar_arquivos_existem(mes_desejado, diretorio_pdf)

            if arquivos_xml_existem or arquivos_pdf_existem:  # Se existirem arquivos XML ou PDF com o mês desejado
                compactar_zip(mes_desejado, diretorio_xml, diretorio_pdf, diretorio_saida, empresa_numero)
                # Chama a função para compactar os arquivos XML e PDF no diretório de saída, específico para a empresa atual

                print('Arquivos compactados e salvos como ZIP')  # Exibe uma mensagem informando que os arquivos foram compactados

                arquivos_zip.append(os.path.join(diretorio_saida, f'arquivos_empresa_{empresa_numero}.zip'))
                # Adiciona o caminho completo do arquivo ZIP criado à lista de arquivos_zip

            else:
                print('Não possui arquivos com esse mês')  # Exibe uma mensagem informando que não há arquivos com o mês desejado

        enviar_email(username, password, destinatario, arquivos_zip)
        # Chama a função para enviar o e-mail contendo os arquivos ZIP para o destinatário

    except FileNotFoundError as e:  # Trata exceção quando ocorre um erro ao abrir um arquivo
        print(f'Erro ao abrir o arquivo: {str(e)}')  # Exibe uma mensagem de erro informando o arquivo que não foi encontrado

    except Exception as e:  # Trata exceção genérica caso ocorra um erro inesperado
        print(f'Ocorreu um erro inesperado: {str(e)}')  # Exibe uma mensagem de erro informando o erro ocorrido



def verificar_arquivos_existem(mes_desejado, diretorio):
    """
    Verifica se existem arquivos com o mês e ano desejados no diretório especificado.

    Args:
        mes_desejado (datetime): Objeto datetime contendo o mês e ano desejados.
        diretorio (str): Diretório a ser verificado.

    Returns:
        bool: True se existirem arquivos com o mês e ano desejados, False caso contrário.
    """
    for nome_arquivo in os.listdir(diretorio):
        caminho_arquivo = os.path.join(diretorio, nome_arquivo)
        timestamp_modificacao = os.path.getmtime(caminho_arquivo)  # Obtém o timestamp de modificação do arquivo
        data_modificacao = datetime.fromtimestamp(timestamp_modificacao)  # Converte o timestamp para objeto datetime

        if nome_arquivo.endswith('.xml') or nome_arquivo.endswith('.pdf'):  # Verifica se o arquivo é XML ou PDF
            if data_modificacao.month == mes_desejado.month and data_modificacao.year == mes_desejado.year:
                # Verifica se o mês e ano de modificação do arquivo correspondem ao mês e ano desejados
                return True  # Retorna True se existir pelo menos um arquivo com o mês e ano desejados

    return False  # Retorna False se nenhum arquivo for encontrado com o mês e ano desejados

def compactar_zip(mes_desejado, diretorio_xml, diretorio_pdf, diretorio_saida, empresa_numero): # Define o nome do arquivo ZIP de saída com base no número da empresa
    nome_arquivo_zip = f'arquivos_empresa_{empresa_numero}.zip'
    caminho_arquivo_zip = os.path.join(diretorio_saida, nome_arquivo_zip) # Define o caminho completo para o arquivo ZIP de saída

    with zipfile.ZipFile(caminho_arquivo_zip, 'w', zipfile.ZIP_DEFLATED) as zipf: # Cria um arquivo ZIP e adiciona os arquivos XML e PDF desejados
        for nome_arquivo in os.listdir(diretorio_xml):  # Percorre os arquivos no diretório de XML
            if nome_arquivo.endswith('.xml'):
                caminho_arquivo = os.path.join(diretorio_xml, nome_arquivo)
                timestamp_modificacao = os.path.getmtime(caminho_arquivo) # Obtém o timestamp de modificação do arquivo
                data_modificacao = datetime.fromtimestamp(timestamp_modificacao) # Converte o timestamp em uma data e hora

                if data_modificacao.month == mes_desejado.month and data_modificacao.year == mes_desejado.year: # Verifica se o arquivo foi modificado no mês e ano desejados
                    zipf.write( 
                        caminho_arquivo,
                        f'xml_{empresa_numero}/{os.path.basename(caminho_arquivo)}' # Adiciona o arquivo ao arquivo ZIP com um diretório específico
                    )
                    print(f"Arquivo XML selecionado: {caminho_arquivo}")

        for nome_arquivo in os.listdir(diretorio_pdf): # Percorre os arquivos no diretório de PDF
            if nome_arquivo.endswith('.pdf'):
                caminho_arquivo = os.path.join(diretorio_pdf, nome_arquivo)
                timestamp_modificacao = os.path.getmtime(caminho_arquivo)   # Obtém o timestamp de modificação do arquivo
                data_modificacao = datetime.fromtimestamp(timestamp_modificacao) # Converte o timestamp em uma data e hora

                if data_modificacao.month == mes_desejado.month and data_modificacao.year == mes_desejado.year: # Verifica se o arquivo foi modificado no mês e ano desejados
                    zipf.write(
                        caminho_arquivo,
                        f'pdf_{empresa_numero}/{os.path.basename(caminho_arquivo)}' # Adiciona o arquivo ao arquivo ZIP com um diretório específico
                    )
                    print(f"Arquivo PDF selecionado: {caminho_arquivo}")


def enviar_email(username, password, destinatario, arquivos_zip):
    msg = MIMEMultipart()  # Cria um objeto MIMEMultipart para compor o e-mail
    msg['From'] = username  # Define o remetente do e-mail
    msg['To'] = destinatario  # Define o destinatário do e-mail
    msg['Subject'] = assunto  # Define o assunto do e-mail
    
    # Adiciona o corpo do e-mail
    texto = corpo
    msg.attach(MIMEText(texto, 'plain'))  # Adiciona o corpo do e-mail como texto simples

        
    for arquivo_zip in arquivos_zip: 
        with open(arquivo_zip, 'rb') as file: # Abre o arquivo ZIP em modo leitura binária
            part_zip = MIMEBase('application', 'zip')  # Cria uma parte MIME do tipo 'application/zip'
            part_zip.set_payload(file.read()) # Define o conteúdo da parte MIME como o conteúdo do arquivo ZIP
            encoders.encode_base64(part_zip) # Codifica o conteúdo da parte MIME em base64
            part_zip.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(arquivo_zip)}"')  # Adiciona o cabeçalho 'Content-Disposition' à parte MIME com o nome do arquivo ZIP
            msg.attach(part_zip) # Anexa a parte MIME ao objeto da mensagem

    servidor_smtp = host
    porta_smtp = port # Configurações de servidor SMTP e porta com base no domínio do remetente

    with smtplib.SMTP(servidor_smtp, porta_smtp) as smtp: # Cria uma conexão com o servidor SMTP
        smtp.starttls() # Inicia a criptografia TLS
        smtp.login(username, password) # Realiza login no servidor SMTP usando o remetente e a senha
        smtp.send_message(msg) # Envia a mensagem por e-mail

    for arquivo_zip in arquivos_zip:
        os.remove(arquivo_zip)  # Remove o arquivo .zip do diretório de saída

    print('E-mail enviado com sucesso!')


def enviar_form():
    mes_desejado = datepicker.get_date()
    mes_desejado_str = mes_desejado.strftime('%m/%Y')
    try:
        ler_configuracoes_smtp()
        compactar_arquivos(empresas, mes_desejado)
        messagebox.showinfo('Envio de E-mail', 'E-mail enviado com sucesso!')
    except ValueError:
        messagebox.showerror('Erro', 'Formato de data inválido. Certifique-se de usar o formato mês e ano, utilizando dois dígitos para o mês e quatro para o ano (MM/YYYY).')

if __name__ == '__main__':
    
    # Create a janela
    window = tk.Tk()
    window.title('Envio de E-mail')
    window.geometry('300x100')
    # Obtém a largura e a altura da tela
    largura_tela = window.winfo_screenwidth()
    altura_tela = window.winfo_screenheight()

    # Calcula as coordenadas X e Y para centralizar a janela
    posicao_x = int(largura_tela / 2 - window.winfo_width() / 2)
    posicao_y = int(altura_tela / 2 - window.winfo_height() / 2)

    # Define a posição da janela no centro da tela
    window.geometry(f"+{posicao_x}+{posicao_y}")
    # Label com texto de seleção de mês
    label_mes = tk.Label(window, text='Selecione o mês desejado:')
    label_mes.pack()

    # Calendário para seleção de data
    class CalendarioTraduzido(tkcalendar.Calendar):  # Chama o construtor da classe pai tkcalendar.Calendar
        def __init__(self, master=None, **kw):
            super().__init__(master, **kw)
            # Define a tradução dos meses do calendário
            self._months = [
                'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
            ]
            self._weekdays = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'] # Define a tradução dos dias da semana do calendário

    datepicker = tkcalendar.DateEntry(window, date_pattern='dd/mm/yyyy', calendar=CalendarioTraduzido)
    datepicker.pack()

    # Botão enviar que chama todas as funções
    button_enviar = tk.Button(window, text='Enviar', command=enviar_form)
    button_enviar.pack()

    # Inicia a janela em loop
    window.mainloop()