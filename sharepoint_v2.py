#: UTILITÁRIO DE SHAREPOINT

#: Import libs
from io import BytesIO
import os
import sys
import extract_n_load.utils.varEnv as varEnv

# Trabalhar com API do 365 - Sharepoint  ---> pip install office365-REST-Python-Client
from office365.runtime.auth.authentication_context import AuthenticationContext 
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.creation_information import FileCreationInformation
from office365.sharepoint.files.file import File 

# Trabalhar com excel
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import NamedStyle

# Obter variaveis do .env
from decouple import Config, RepositoryEnv
def _resource_path():
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath("..")

    env_path = os.path.abspath(os.path.join(base_path,'.env'))
    #env_path = os.path.abspath(os.path.join(base_path,'..', '.env'))
    
    return env_path

config = Config(RepositoryEnv(_resource_path()))


#/:
#: Definição de Variáveis
SITE_URL = config("URL_SHP")
USERNAME = config("USER_SHP")
PASSWORD = config("PASS_SHP")
TARGET_FOLDER_LOG = config("DESTINO_LOG")


#/: 
#: Definição de Funções
def autenticacao_sharepoint(site:str = SITE_URL, username:str = USERNAME, password:str = PASSWORD) -> ClientContext:
    """Função de autenticação no sharepoint. Garanta que seu usuário não necessita de MFA.

    Args:
        site (str, optional): Url do site do sharepoint. Defaults to SITE_URL.
        username (str, optional): Nome de usuario. Defaults to USERNAME.
        password (str, optional): Senha. Defaults to PASSWORD.

    Returns:
        ClientContext: Contexto Web autenticado no sharepoint
    """
    ctx_auth = AuthenticationContext(site)
    ctx_auth.acquire_token_for_user(username, password)
    ctx = ClientContext(site, ctx_auth)
    
    return ctx

       
def log_erro(lista_erro:list,file_name:str,ctx:ClientContext):
    """Upload de lista de Erros em um arquivo txt para Sharepoint

    Args:
        lista_erro (list): Conteudo q será escrito
        file_name (str): Nome do arquivo
        ctx (ClientContext): Contexto do sharepoint obtido da função de autenticação
    """
       
    file_info = FileCreationInformation()
    file_content = "\n".join(lista_erro) # Conteúdo do arquivo a ser enviado
    file_info.content = file_content.encode('utf-8')  # Codificar o conteúdo do arquivo
    file_info.url = TARGET_FOLDER_LOG + '/' + file_name

    # Upload do arquivo
    target_folder = ctx.web.get_folder_by_server_relative_url(TARGET_FOLDER_LOG)
    uploaded_file = target_folder.files.add(file_info.url,file_info.content)
    ctx.execute_query()
    
    
def upload_file(original_file:str,target_folder_url:str,ctx:ClientContext):
    """Upload de arquivo local para o sharepoint

    Args:
        original_file (str): Caminho local do arquivo que será alocado no sharepoint
        target_folder_url (str): Caminho da pasta de Sharepoint
        ctx (ClientContext): Contexto do sharepoint obtido da função de autenticação
    """
    
    # Convertendo o dataframe para um arquivo Excel em memória
    # abre o arquivo localmente
    with open(original_file,"rb") as content_file:
        file_content = content_file.read()
    # Define nome do arquivo e path (target)
    dir, name = os.path.split(target_folder_url)
    # Escreve arquivo no sharepoint
    file = ctx.web.get_folder_by_server_relative_url(dir).upload_file(name, file_content).execute_query()
    
    
def ler_objeto(arquivo:str, ctx:ClientContext) -> BytesIO:
    """Ler um arquivo alocado no sharepoint

    Args:
        arquivo (str): caminho do arquivo no SharePoint
        ctx (ClientContext): Contexto do sharepoint obtido da função de autenticação

    Returns:
        BytesIO: Ponteiro do conteúdo do arquivo. Este retorno pode ser utilizado em uma função de leitura como _pd.read_excel()_
    """
    # Abre o arquivo binário no contexto atual e na URL relativa ao servidor especificada nas propriedades do projeto
    response = File.open_binary(ctx, arquivo)
    # Cria um objeto de arquivo de bytes na memória
    bytes_file_obj = BytesIO()
    # Escreve o conteúdo binário da resposta no objeto de arquivo de bytes
    bytes_file_obj.write(response.content)
    # Retorna o ponteiro do arquivo para o início do arquivo de bytes
    bytes_file_obj.seek(0)
    
    return bytes_file_obj


def escrita_excel(destino_download:str,dados,file_sharepoint:str,ctx:ClientContext):
    """
    Escreve arquivo em excel no sharepoint

    Args:
        destino_download (str): caminho local onde o arquivo será temporariamente armazenado antes de subir para Sharepoint
        dados (pd.dataFrame): dataframe contendo os dados que serão escritos em excel
        file_sharepoint (str): caminho de destino no sharepoint
        ctx (ClientContext): Contexto do sharepoint obtido da função de autenticação
    """
    
    # Executar escrita de arquivo excel
    diretorio_atual = os.getcwd()
    destino_download = diretorio_atual + rf"{destino_download}"
    dados.to_excel(destino_download,index=False)
    
    # Escrita do excel no sharepoint
    upload_file(destino_download,file_sharepoint,ctx)
    
    # Excluir arquivo local
    os.remove(destino_download)
    

options = ["Inconsistente", "Vulnerável", "Mitigado", "Corrigido", "Aceito", "Eliminado da Ferramenta"]
datecols = ["G", "I", "L", "M"]
optionscols = 'H2:H1048576'
    
def escrita_excel_format(destino_download:str,dados,file_sharepoint:str,ctx:ClientContext,
                         options:list = options, optionscols:str = optionscols,
                         datecols:list = datecols):
    """
    Escreve arquivo em excel no sharepoint com formatação de data nas colunas especificadas e validação de dados em uma coluna.

    Args:
        destino_download (str): caminho local onde o arquivo será temporariamente armazenado antes de subir para Sharepoint
        dados (pd.dataFrame): dataframe contendo os dados que serão escritos em excel
        file_sharepoint (str): caminho de destino no sharepoint
        ctx (ClientContext): Contexto do sharepoint obtido da função de autenticação
        options (list, optional): Lista de opções para a coluna com validação de dados. 
        optionscols (str, optional): Range de linhas da coluna que se quer aplicar a validação. _Ex: 'H1:H9999'_
        datecols (list, optional): Lista de colunas que se quer configurar com o formato de Data. _Ex: ['A','B']_
    """
    
    # Executar escrita de arquivo excel
    diretorio_atual = os.getcwd()
    destino_download = diretorio_atual + rf"{destino_download}"
    dados.to_excel(destino_download,index=False)
    
    # Abrindo o arquivo com openpyxl para adicionar a validação de dados
    wb = load_workbook(destino_download)
    ws = wb.active

    ###### VALIDAÇÃO DE CAMPOS
    # Definindo a lista suspensa para a coluna "Status"
    
    #status_range = f"D2:D1048576"  # Aplica a validação para toda a coluna D (exceto o cabeçalho)

    # Criar a validação de dados (sem afetar valores existentes)
    dv = DataValidation(type="list", formula1=f'"{",".join(options)}"', allow_blank=True)
    #dv.ranges.append(status_range)

    # Adicionar a validação na planilha
    ws.add_data_validation(dv)

    # Adicionar validação às celulas
    dv.add(optionscols)
    
    ###### FORMATAR COMO DATA
    date_style = NamedStyle(name="date_style", number_format="DD/MM/YYYY")

    # Aplicar o estilo às colunas de data
    for col in datecols:
        for row in range(2, ws.max_row + 1):  # Começa da linha 2 para não alterar o cabeçalho
            cell = ws[f"{col}{row}"]
            cell.style = date_style
    
    #### SALVAR
    # Salvando o arquivo Excel final no diretório
    wb.save(destino_download)
        
    # Escrita do excel no sharepoint
    upload_file(destino_download,file_sharepoint,ctx)
    
    # Excluir arquivo local
    os.remove(destino_download)