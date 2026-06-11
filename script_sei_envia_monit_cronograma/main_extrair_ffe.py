import logging
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
HOJE = datetime.now().strftime('%Y_%m_%d')
AGORA = datetime.now().strftime('%Y%m%d_%H%M%S')

DIR_BASE = Path(__file__).resolve().parent.parent

DIR_DADOS = DIR_BASE / 'dados'
DIR_LOGS = DIR_BASE / 'logs' / 'extrair' / 'ffe'

DIR_DADOS_INPUTS = DIR_DADOS / 'inputs'
DIR_DADOS_REGS = DIR_DADOS / 'ffe' / 'regionais'
DIR_DADOS_GERAL = DIR_DADOS / 'ffe' / 'geral'

DIR_REDE_FFE = Path('H:/Sisso-Arquivos/Estatistica')

ARQ_LOG = DIR_LOGS / f'extrair-{AGORA}.log'
ARQ_FFE_COMPLETO = DIR_DADOS_GERAL / f'ffes_completos-{HOJE}.xlsx'
ARQ_REM_EM_CADASTRAMENTO = (
    DIR_DADOS_GERAL / f'remessas_em_cadastramento-ffe-{HOJE}.xlsx'
)
ARQ_REM_PAG_PROCESSADO = (
    DIR_DADOS_GERAL / f'remessas_pag_processado_zeradas-ffe-{HOJE}.xlsx'
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


def path_ffe(ano: int) -> Path:
    return DIR_REDE_FFE / f'FFE_{ano}.CSV'


_COLS_MONIT = [
    'MUNICIPIO',
    'COD_PESSOA_PRESTADOR',
    'NOME_PRESTADOR',
    'RAMO_ATIVIDADE',
    'MES_REFERENCIA',
    'NUM_REMESSA',
    'VALOR_PAGAR',
]


def preparar_df_monit(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df[_COLS_MONIT]
        .rename(columns={'COD_PESSOA_PRESTADOR': 'MATRICULA'})
        .assign(OBSERVAÇÕES='')
    )


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
                fmt = fmt_decimal if col_name.startswith('VAL') else None
                worksheet.set_column(col_idx, col_idx, largura, fmt)


def gerar_df_ffes_completo(ano_ini: int, ano_final: int) -> pd.DataFrame:
    dfs_ffe: list[pd.DataFrame] = []

    for ano in range(ano_ini, ano_final + 1):
        arq_ffe = path_ffe(ano)
        df_ffe = pd.read_csv(
            arq_ffe,
            sep=';',
            skipinitialspace=True,
            encoding='latin1',
            skipfooter=1,
            engine='python',
            index_col=False,
        )
        logger.info('FFE_%d carregado: %d linhas', ano, len(df_ffe))
        dfs_ffe.append(df_ffe)

    df_ffe_completo: pd.DataFrame = pd.concat(dfs_ffe, ignore_index=True)
    logger.info('FFEs concatenados: %d linhas no total', len(df_ffe_completo))

    val_cols = [c for c in df_ffe_completo.columns if c.startswith('VAL')]
    df_ffe_completo[val_cols] = df_ffe_completo[val_cols].apply(
        lambda s: pd.to_numeric(
            s.astype(str).str.replace(',', '.', regex=False), errors='coerce'
        )
    )

    total = len(df_ffe_completo)
    df_ffe_completo_filtrado = df_ffe_completo[
        (df_ffe_completo['TIPO_CONTRATO'] == 'CREDENCIADO PESSOA JURIDICA')
        & (
            (df_ffe_completo['MES_REFERENCIA'] >= '2022/01')
            & (df_ffe_completo['MES_REFERENCIA'] <= '2025/12')
        )
        & (df_ffe_completo['SITUACAO'] != 'PAGAMENTO EFETUADO')
    ]
    logger.info(
        'Filtro aplicado: %d/%d linhas retidas (PJ, 2022/01-2025/12)',
        len(df_ffe_completo_filtrado),
        total,
    )

    return df_ffe_completo_filtrado


def main() -> None:
    df_ffes = gerar_df_ffes_completo(2022, 2026)

    df_ffes.to_excel(ARQ_FFE_COMPLETO, index=False, engine='xlsxwriter')
    logger.info(
        'FFEs completos salvos: %s (%d linhas)',
        ARQ_FFE_COMPLETO.name,
        len(df_ffes),
    )

    df_rem_em_cadastramento = df_ffes[
        df_ffes['SITUACAO'] == 'EM CADASTRAMENTO'
    ]
    df_rem_em_cadastramento.to_excel(
        ARQ_REM_EM_CADASTRAMENTO, index=False, engine='xlsxwriter'
    )
    logger.info(
        'EM CADASTRAMENTO salvo: %s (%d linhas)',
        ARQ_REM_EM_CADASTRAMENTO.name,
        len(df_rem_em_cadastramento),
    )

    df_rem_pag_processado_vl_zerado = df_ffes[
        (df_ffes['SITUACAO'] == 'PAGAMENTO PROCESSADO')
        & (df_ffes['VALOR_PAGAR'] == 0)
    ]
    df_rem_pag_processado_vl_zerado.to_excel(
        ARQ_REM_PAG_PROCESSADO, index=False, engine='xlsxwriter'
    )
    logger.info(
        'PAG. PROCESSADO zerado salvo: %s (%d linhas)',
        ARQ_REM_PAG_PROCESSADO.name,
        len(df_rem_pag_processado_vl_zerado),
    )

    df_em_revisao = df_ffes[df_ffes['SITUACAO'] == 'EM REVISAO']
    df_liberado_pag = df_ffes[df_ffes['SITUACAO'] == 'LIBERADO PARA PAGAMENTO']
    df_pag_proc_zerado = df_ffes[
        (df_ffes['SITUACAO'] == 'PAGAMENTO PROCESSADO')
        & (df_ffes['VALOR_PAGAR'] == 0)
    ]
    logger.info(
        'Monitoramento: EM REVISAO=%d | LIBERADO PAG=%d | PAG PROC ZERADO=%d',
        len(df_em_revisao),
        len(df_liberado_pag),
        len(df_pag_proc_zerado),
    )

    df_monitoramento = preparar_df_monit(
        pd.concat(
            [df_em_revisao, df_liberado_pag, df_pag_proc_zerado],
            ignore_index=True,
        )
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
        municipio = reg['MUNICIPIO']
        nome_ffe = reg['UNIDADE_REGIONAL_FFE']
        arq_rem_regional = (
            DIR_DADOS_REGS / f'remessas-ffe-{municipio}-{HOJE}.xlsx'
        )

        salvar_xlsx_por_abas(
            {
                'EM REVISAO': preparar_df_monit(
                    df_em_revisao[
                        df_em_revisao['UNIDADE_REGIONAL'] == nome_ffe
                    ]
                ),
                'LIBERADO PAGAMENTO': preparar_df_monit(
                    df_liberado_pag[
                        df_liberado_pag['UNIDADE_REGIONAL'] == nome_ffe
                    ]
                ),
                'PAGAMENTO PROCESSADO ZERADO': preparar_df_monit(
                    df_pag_proc_zerado[
                        df_pag_proc_zerado['UNIDADE_REGIONAL'] == nome_ffe
                    ]
                ),
            },
            arq_rem_regional,
        )
        logger.info('Regional %s: %s gerado', municipio, arq_rem_regional.name)

    logger.info('Processamento concluído')


if __name__ == '__main__':
    main()
