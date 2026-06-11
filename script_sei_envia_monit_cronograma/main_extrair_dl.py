import logging
import sys
from datetime import datetime
from pathlib import Path

import jaydebeapi
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

HOJE = datetime.now().strftime('%Y_%m_%d')
HOJE_DBEAVER = datetime.now().strftime('%Y%m%d')
AGORA = datetime.now().strftime('%Y%m%d_%H%M%S')

DIR_BASE = Path(__file__).resolve().parent.parent

DIR_DADOS = DIR_BASE / 'dados'
DIR_LOGS = DIR_BASE / 'logs' / 'extrair' / 'dl'

DIR_DADOS_INPUTS = DIR_DADOS / 'inputs'
DIR_DADOS_REGS = DIR_DADOS / 'dl' / 'regionais'
DIR_DADOS_GERAL = DIR_DADOS / 'dl' / 'geral'


ARQ_LOG = DIR_LOGS / f'extrair-{AGORA}.log'
ARQ_COMPLETO = DIR_DADOS_GERAL / f'dl_completo-{HOJE_DBEAVER}.csv'
ARQ_REM_EM_CADASTRAMENTO = (
    DIR_DADOS_GERAL / f'remessas_em_cadastramento-dl-{HOJE}.xlsx'
)
ARQ_REM_PAG_PROCESSADO = (
    DIR_DADOS_GERAL / f'remessas_pag_processado_zeradas-dl-{HOJE}.xlsx'
)
ARQ_MONITORAMENTO = DIR_DADOS_GERAL / f'monitoramento-{HOJE}.xlsx'
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

_COLS_MONIT = [
    'regional',
    'municipio',
    'prestador_matricula',
    'prestador_nome',
    'ramo',
    'referencia',
    'remessa',
    'valor_pagar',
]


def preparar_df_monit(df: pd.DataFrame) -> pd.DataFrame:
    return df[_COLS_MONIT].assign(observacoes='')


def connect_to_db():
    conn = jaydebeapi.connect(
        'com.cloudera.impala.jdbc.Driver',
        (
            'jdbc:impala://dlmg.prodemge.gov.br:21051;'
            'AuthMech=3;'
            'SSL=1;'
            'transportMode=sasl;'
            'AllowHostNameCNMismatch=1;'
            'AllowSelfSignedCerts=1;'
            'UID=755247;'
            'PWD=Gmespi0825'
        ),
        jars=r'C:\Users\m755247\Downloads\ClouderaImpala_JDBC-2.6.15.1017.zip',
    )

    return conn


def salvar_xlsx_por_abas(abas: dict[str, pd.DataFrame], destino: Path) -> None:
    with pd.ExcelWriter(destino, engine='xlsxwriter') as writer:
        for nome_aba, df in abas.items():
            sheet = nome_aba[:31]
            df.to_excel(writer, sheet_name=sheet, index=False)
            fmt_decimal = writer.book.add_format({'num_format': '#,##0.00'})
            worksheet = writer.sheets[sheet]
            for col_idx, col_name in enumerate(df.columns):
                raw = df[col_name].astype(str).str.len().max()
                max_dados = 0 if pd.isna(raw) else int(raw)
                largura = max(len(col_name), max_dados) + 2
                fmt = fmt_decimal if col_name.startswith('valor') else None
                worksheet.set_column(col_idx, col_idx, largura, fmt)


def main() -> None:
    # conn = connect_to_db()

    # df_completo = pd.read_sql(
    #     """
    #     SELECT
    #         nome_agencia_regional_contrato_prestador AS regional,
    #         nome_regional_contrato_prestador AS regional,
    #         numero_matricula_prestador AS prestador_matricula,
    #         nome_prestador AS prestador_nome,
    #         sigla_ramo_atividade AS ramo,
    #         data_referencia_fatura AS referencia,
    #         numero_remessa AS remessa,
    #         NULL AS observacoes
    #     FROM db_dlipsemg_reporting.faturamento_detalhado
    #     WHERE
    #         tipo_contrato = 'CREDPJ'
    #         AND data_referencia_fatura >= '2022-01-01'
    #         AND data_referencia_fatura <= '2025-12-01'
    #     ORDER BY
    #         regional ASC,
    #         referencia ASC,
    #         prestador_matricula ASC
    #     """,
    #     con=conn,
    # )

    df_completo = pd.read_csv(
        ARQ_COMPLETO,
        sep=';',
        encoding='utf-8',
        index_col=False,
    )

    logger.info(
        'FFEs completos salvos: %s (%d linhas)',
        ARQ_COMPLETO.name,
        len(df_completo),
    )

    df_rem_em_cadastramento = df_completo[
        df_completo['situacao'] == 'CADASTRO'
    ]
    df_rem_em_cadastramento.to_excel(
        ARQ_REM_EM_CADASTRAMENTO, index=False, engine='xlsxwriter'
    )
    logger.info(
        'EM CADASTRAMENTO salvo: %s (%d linhas)',
        ARQ_REM_EM_CADASTRAMENTO.name,
        len(df_rem_em_cadastramento),
    )

    df_rem_pag_processado_vl_zerado = df_completo[
        (df_completo['situacao'] == 'PAGPROC')
        & (df_completo['valor_pagar'] == 0)
    ]
    df_rem_pag_processado_vl_zerado.to_excel(
        ARQ_REM_PAG_PROCESSADO, index=False, engine='xlsxwriter'
    )
    logger.info(
        'PAG. PROCESSADO zerado salvo: %s (%d linhas)',
        ARQ_REM_PAG_PROCESSADO.name,
        len(df_rem_pag_processado_vl_zerado),
    )

    df_em_revisao = df_completo[df_completo['situacao'] == 'EMREVISAO']

    df_liberado_pag = df_completo[df_completo['situacao'] == 'LIBPAGTO']

    df_pag_proc_zerado = df_completo[
        (df_completo['situacao'] == 'PAGPROC')
        & (df_completo['valor_pagar'] == 0)
    ]

    logger.info(
        'Monitoramento: EM REVISAO=%d | LIBERADO PAG=%d | PAG PROC ZERADO=%d',
        len(df_em_revisao),
        len(df_liberado_pag),
        len(df_pag_proc_zerado),
    )

    df_monitoramento = pd.concat(
        [df_em_revisao, df_liberado_pag, df_pag_proc_zerado],
        ignore_index=True,
    )

    salvar_xlsx_por_abas(
        {'MONITORAMENTO': df_monitoramento}, ARQ_MONITORAMENTO
    )
    logger.info(
        'Monitoramento consolidado salvo: %s (%d linhas)',
        ARQ_MONITORAMENTO.name,
        len(df_monitoramento),
    )

    df_regionais = pd.read_excel(ARQ_CONTATOS_REGIONAIS, index_col=False)
    logger.info(
        'Gerando arquivos por regional (%d regionais)', len(df_regionais)
    )

    for _, reg in df_regionais.iterrows():
        regional = reg['MUNICIPIO']
        arq_rem_regional = (
            DIR_DADOS_REGS / f'AGREGADO-remessas-dl-{regional}-{HOJE}.xlsx'
        )

        salvar_xlsx_por_abas(
            {
                'EM REVISAO': preparar_df_monit(
                    df_em_revisao[df_em_revisao['regional'] == regional]
                ),
                'LIBERADO PAGAMENTO': preparar_df_monit(
                    df_liberado_pag[df_liberado_pag['regional'] == regional]
                ),
                'PAGAMENTO PROCESSADO ZERADO': preparar_df_monit(
                    df_pag_proc_zerado[
                        df_pag_proc_zerado['regional'] == regional
                    ]
                ),
            },
            arq_rem_regional,
        )
        logger.info('Regional %s: %s gerado', regional, arq_rem_regional.name)

    logger.info('Processamento concluído')


if __name__ == '__main__':
    main()
