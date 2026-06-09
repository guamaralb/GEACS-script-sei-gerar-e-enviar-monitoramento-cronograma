import logging
import os
import sys
from datetime import datetime
from pathlib import Path
from time import sleep

import pandas as pd
from dotenv import load_dotenv
from sei_automacao.acoes_completas.processo import incluir_doc_externo
from sei_automacao.driver.iniciar_driver import iniciar_driver
from sei_automacao.genericos.botoes import clicar_salvar_btnSalvar
from sei_automacao.inicio.efetuar_login import efetuar_login
from sei_automacao.menu.abrir_menu import abrir_menu
from sei_automacao.menu.iniciar_processo import (
    clicar_iniciar_processo,
    preencher_especificacao_processo,
    selecionar_tipo_processo,
)
from sei_automacao.ubiquo.acessar_processo import acessar_processo
from sei_automacao.utils.selecionar_nivel_acesso import selecionar_nivel_acesso
from sei_automacao.utils.trocar_iframe import trocar_iframe

load_dotenv()

SEI_LOGIN = os.getenv('SEI_LOGIN')
SEI_SENHA = os.getenv('SEI_SENHA')
SEI_ORGAO = os.getenv('SEI_ORGAO')

sys.stdout.reconfigure(encoding='utf-8')
HOJE = datetime.now().strftime('%Y_%m_%d')
AGORA = datetime.now().strftime('%Y%m%d_%H%M%S')

DIR_BASE = Path(__file__).resolve().parent.parent

DIR_DADOS = DIR_BASE / 'dados'
DIR_LOGS = DIR_BASE / 'logs' / 'enviar'

DIR_DADOS_INPUTS = DIR_DADOS / 'inputs'
DIR_DADOS_REGS = DIR_DADOS / 'regionais'

DIR_REDE_FFE = Path('H:/Sisso-Arquivos/Estatistica')

ARQ_LOG = DIR_LOGS / f'enviar-{AGORA}.log'
ARQ_CONTATOS_REGIONAIS = DIR_DADOS_INPUTS / 'Contatos Regionais.xlsx'

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(ARQ_LOG, encoding='utf-8'),
        logging.StreamHandler(),
    ],
)

logger = logging.getLogger(__name__)


def main() -> None:
    driver = iniciar_driver()

    driver.maximize_window()
    driver.get('https://www.sei.mg.gov.br')

    efetuar_login(driver, SEI_LOGIN, SEI_SENHA, SEI_ORGAO)
    sleep(2)

    acessar_processo(driver, num_processo='2010.01.0054229/2026-05')

    sleep(1)

    df_regionais = pd.read_excel(
        ARQ_CONTATOS_REGIONAIS, sheet_name='teste', index_col=False
    )
    logger.info(
        'Gerando processos por regional (%d regionais)', len(df_regionais)
    )

    for _, reg in df_regionais.iterrows():
        municipio = reg['MUNICIPIO']
        arq_rem_regional = DIR_DADOS_REGS / f'remessas-{municipio}-{HOJE}.xlsx'
        espec_processo = (
            f'TESTE - Monitoramento de Cronograma - Regional de {municipio}'
        )

        abrir_menu(driver)
        sleep(1)
        clicar_iniciar_processo(driver)
        sleep(1)
        selecionar_tipo_processo(
            driver,
            tipo_processo='Pedidos, Oferecimentos e Informações Diversas',
        )
        sleep(1)
        preencher_especificacao_processo(driver, especificacao=espec_processo)
        selecionar_nivel_acesso(driver, nivel_acesso='Público')
        clicar_salvar_btnSalvar(driver)
        trocar_iframe(driver, iframe='ifrConteudoVisualizacao')
        incluir_doc_externo(
            driver,
            tipo_doc='Planilha',
            num='',
            formato='Nato-digital',
            data='09/06/2026',
            caminho_anexo=arq_rem_regional,
            nivel_acesso='Público',
        )

        sleep(100)
        logger.info('Regional %s: %s gerado', municipio, arq_rem_regional.name)

    logger.info('Processamento concluído')


if __name__ == '__main__':
    main()
