from dotenv import load_dotenv
load_dotenv()

from playwright.sync_api import sync_playwright
from src.browser import abrir_navegador, pagina_aging, filtros_bi, exportar_dados
from src.processing import ler_arquivos, tratamento, salvar_arquivo
from utils.dates import data_atual, formatar_data, formatar_data_aging
from src.pivot import atualizar_tabelas

import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    filename='logs\historico.log',
    filemode='a',
    encoding='utf-8'
)

data_ontem = data_atual()
data_formatada = formatar_data(data_ontem)
data_aging = formatar_data_aging(data_ontem)

# Exportação de arquivo BI
data_bi = None

with sync_playwright() as pw:
    try:
        logging.info('Extraindo dados do Power BI...')
        browser = abrir_navegador(pw)

        aging = pagina_aging(browser)

        data_bi = filtros_bi(aging, data_formatada)

        exportar_dados(aging, data_aging)

        browser.close()
        logging.info('Extração concluída!')
    except Exception as e:
        logging.error(f'Ocorreu uma falha ao extrair os dados: {e}')

# Tratamento de dados
try:
    logging.info('Começando o tratamento de dados..')
    dados = ler_arquivos()

    tratamento(dados, data_ontem, data_bi)

    salvar_arquivo(data_aging, dados["aging"])

except Exception as e:
    logging.exception(f'Ocorreu uma falha ao realizar o tratamento: {e}')

# Atualização de tabelas dinâmicas
try:
    logging.info('Atualizando as tabelas dinâmicas..')
    atualizar_tabelas(data_aging)

except Exception as e:
    logging.exception(f'Houve um problema ao atualizar as tabelas: {e}')



