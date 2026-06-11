import logging
import os
import sys
from datetime import datetime
from pathlib import Path
from time import sleep

import pandas as pd
from dotenv import load_dotenv
from sei_automacao.acoes_completas.processo import (
    enviar_email,
)

# from sei_automacao.acoes_completas.processo import incluir_doc_externo
from sei_automacao.driver.iniciar_driver import iniciar_driver
from sei_automacao.inicio.efetuar_login import efetuar_login
from sei_automacao.ubiquo.acessar_processo import acessar_processo
from txt_email import gerar_assunto_email, gerar_corpo_email

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
DIR_DADOS_REGS = DIR_DADOS / 'dl' / 'regionais'

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
    driver = iniciar_driver(driver_type='edge')

    driver.maximize_window()
    driver.get('https://www.sei.mg.gov.br')

    efetuar_login(driver, SEI_LOGIN, SEI_SENHA, SEI_ORGAO)
    sleep(2)

    acessar_processo(driver, num_processo='2010.01.0055691/2026-10')

    sleep(1)

    df_regionais = pd.read_excel(
        ARQ_CONTATOS_REGIONAIS, sheet_name='teste', index_col=False
    )
    logger.info(
        'Gerando processos por regional (%d regionais)', len(df_regionais)
    )

    path_anexo_pdf = (
        DIR_DADOS_INPUTS / 'Orientações por situação da remessa - '
        '2026 (Em revisão, Liberado e Pag. Processado Zerado).pdf'
    )

    for _, reg in df_regionais.iterrows():
        municipio = reg['MUNICIPIO']
        email_regional = reg['EMAIL_UNIDADE_REGIONAL']

        path_arq_rem_regional = (
            DIR_DADOS_REGS / f'remessas-{municipio}-{HOJE}.xlsx'
        )

        acessar_processo(driver, num_processo='2010.01.0055691/2026-10')

        # espec_processo = (
        #     f'TESTE - Monitoramento de Cronograma - Regional de {municipio}'
        # )

        # iniciar_processo(
        #     driver,
        #     tipo_processo='Pedidos, Oferecimentos e Informações Diversas',
        #     especificacao=espec_processo,
        #     nivel_acesso='Público',
        # )
        # incluir_doc_externo(
        #     driver,
        #     tipo_doc='Planilha',
        #     num=path_arq_rem_regional.name + AGORA,
        #     formato='Nato-digital',
        #     data='09/06/2026',
        #     path_anexo=path_arq_rem_regional,
        #     nivel_acesso='Público',
        # )
        # incluir_doc_externo(
        #     driver,
        #     tipo_doc='PDF',
        #     num='Orientações por situação da remessa - 2026',
        #     formato='Nato-digital',
        #     data='09/06/2026',
        #     path_anexo=path_anexo_pdf,
        #     nivel_acesso='Público',
        #     fecha_alerta=True,
        # )
        enviar_email(
            driver=driver,
            email_de='IPSEMG/GEACS <faturamento.pagamento@ipsemg.mg.gov.br>',
            emails_para=[email_regional],
            assunto=gerar_assunto_email(),
            corpo_email=gerar_corpo_email(
                municipio=municipio, nome_anexo_pdf=path_anexo_pdf.name
            ),
            nivel_acesso='Público',
        )
        logger.info(
            'Regional %s: %s gerado', municipio, path_arq_rem_regional.name
        )

    logger.info('Processamento concluído')


if __name__ == '__main__':
    main()
