import logging
import sys
from datetime import datetime
from pathlib import Path

import jaydebeapi
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

HOJE = datetime.now().strftime('%Y_%m_%d')
AGORA = datetime.now().strftime('%Y%m%d_%H%M%S')

DIR_BASE = Path(__file__).resolve().parent.parent

DIR_DADOS = DIR_BASE / 'dados'
DIR_LOGS = DIR_BASE / 'logs' / 'extrair' / 'dl'

DIR_DADOS_INPUTS = DIR_DADOS / 'inputs'
DIR_DADOS_REGS = DIR_DADOS / 'dl' / 'regionais'
DIR_DADOS_GERAL = DIR_DADOS / 'dl' / 'geral'


ARQ_LOG = DIR_LOGS / f'extrair-{AGORA}.log'
ARQ_COMPLETO = DIR_DADOS_INPUTS / f'dl_completo-{HOJE}.csv'
ARQ_REM_EM_CADASTRAMENTO = (
    DIR_DADOS_GERAL / f'rem_em_cadastramento-dl-{HOJE}.xlsx'
)
ARQ_REM_PAG_PROCESSADO = (
    DIR_DADOS_GERAL / f'rem_pag_processado_zeradas-dl-{HOJE}.xlsx'
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

_COLS_COM_REM = [
    'regional',
    'municipio',
    'prestador_matricula',
    'prestador_nome',
    'ramo',
    'referencia',
    'remessa',
    'tipo',
    'valor_pagar',
]


_COLS_COM_RB = [
    'regional',
    'municipio',
    'prestador_matricula',
    'prestador_nome',
    'ramo',
    'referencia',
    'rb',
    'tipo',
    'valor_pagar',
]


def agregar_por_remessa(df: pd.DataFrame) -> pd.DataFrame:
    return df[_COLS_COM_REM].assign(observacoes='')


def agregar_por_rb(df: pd.DataFrame) -> pd.DataFrame:
    return df[_COLS_COM_RB].assign(observacoes='')


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


def carregar_df_completo() -> pd.DataFrame:
    df_completo = pd.read_csv(
        ARQ_COMPLETO,
        sep=';',
        encoding='utf-8',
        index_col=False,
    )

    df_completo_filtro_tipo = df_completo[df_completo['tipo'] != 'DIGITADA']
    print(df_completo_filtro_tipo)

    logger.info(
        'DL completo salvo: %s (%d linhas)',
        ARQ_COMPLETO.name,
        len(df_completo_filtro_tipo),
    )

    return df_completo_filtro_tipo


def agregar_situacoes(
    df_completo_filtro_tipo: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    _group_rem = [c for c in _COLS_COM_REM if c != 'valor_pagar']
    _group_rb = [c for c in _COLS_COM_RB if c != 'valor_pagar']

    df_agg_por_rem = (
        df_completo_filtro_tipo[_COLS_COM_REM + ['situacao']]
        .groupby(_group_rem + ['situacao'], dropna=False)
        .agg(valor_pagar=('valor_pagar', 'sum'))
        .reset_index()
    )

    df_agg_por_rb = (
        df_completo_filtro_tipo[_COLS_COM_RB + ['situacao']]
        .groupby(_group_rb + ['situacao'], dropna=False)
        .agg(valor_pagar=('valor_pagar', 'sum'))
        .reset_index()
    )

    return df_agg_por_rem, df_agg_por_rb


def salvar_rem_em_cadastramento(df_agg_por_rem: pd.DataFrame) -> None:
    df_rem_em_cadastramento = agregar_por_remessa(
        df_agg_por_rem[df_agg_por_rem['situacao'] == 'CADASTRO']
    )
    df_rem_em_cadastramento.to_excel(
        ARQ_REM_EM_CADASTRAMENTO, index=False, engine='xlsxwriter'
    )
    logger.info(
        'EM CADASTRAMENTO salvo: %s (%d linhas)',
        ARQ_REM_EM_CADASTRAMENTO.name,
        len(df_rem_em_cadastramento),
    )


def salvar_rem_pag_processado_zerado(df_agg_por_rb: pd.DataFrame) -> None:
    df_rem_pag_processado_vl_zerado = agregar_por_rb(
        df_agg_por_rb[
            (df_agg_por_rb['situacao'] == 'PAGPROC')
            & (df_agg_por_rb['valor_pagar'] == 0)
        ]
    )
    df_rem_pag_processado_vl_zerado.to_excel(
        ARQ_REM_PAG_PROCESSADO, index=False, engine='xlsxwriter'
    )
    logger.info(
        'PAG. PROCESSADO zerado salvo: %s (%d linhas)',
        ARQ_REM_PAG_PROCESSADO.name,
        len(df_rem_pag_processado_vl_zerado),
    )


def salvar_monitoramento(
    df_em_revisao: pd.DataFrame,
    df_liberado_pag: pd.DataFrame,
    df_pag_proc_zerado: pd.DataFrame,
) -> None:
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


def gerar_relatorio_regional(
    regional: str,
    df_em_revisao: pd.DataFrame,
    df_liberado_pag: pd.DataFrame,
    df_pag_proc_zerado: pd.DataFrame,
) -> None:
    arq_rem_regional = DIR_DADOS_REGS / f'rem-dl-{regional}-{HOJE}.xlsx'

    df_rev = df_em_revisao[df_em_revisao['regional'] == regional]
    df_lib = df_liberado_pag[df_liberado_pag['regional'] == regional]
    df_zer = df_pag_proc_zerado[df_pag_proc_zerado['regional'] == regional]

    df_resumo = pd.DataFrame({
        'STATUS': [
            'EM REVISAO',
            'LIBERADO PAGAMENTO',
            'PAGAMENTO PROCESSADO ZERADO',
        ],
        'QUANTIDADE': [
            len(df_rev),
            len(df_lib),
            len(df_zer),
        ],
    })

    abas = {
        'RESUMO': df_resumo,
    }

    if not df_rev.empty:
        abas['EM REVISAO'] = agregar_por_remessa(df_rev)

    if not df_lib.empty:
        abas['LIBERADO PAGAMENTO'] = agregar_por_remessa(df_lib)

    if not df_zer.empty:
        abas['PAGAMENTO PROCESSADO ZERADO'] = agregar_por_rb(df_zer)

    salvar_xlsx_por_abas(abas, arq_rem_regional)

    logger.info('Regional %s: %s gerado', regional, arq_rem_regional.name)


def gerar_relatorios_regionais(
    df_em_revisao: pd.DataFrame,
    df_liberado_pag: pd.DataFrame,
    df_pag_proc_zerado: pd.DataFrame,
) -> None:
    df_regionais = pd.read_excel(ARQ_CONTATOS_REGIONAIS, index_col=False)
    logger.info(
        'Gerando arquivos por regional (%d regionais)', len(df_regionais)
    )

    for _, reg in df_regionais.iterrows():
        gerar_relatorio_regional(
            reg['MUNICIPIO_DL'],
            df_em_revisao,
            df_liberado_pag,
            df_pag_proc_zerado,
        )


def main() -> None:
    df_completo_filtro_tipo = carregar_df_completo()

    df_agg_por_rem, df_agg_por_rb = agregar_situacoes(df_completo_filtro_tipo)

    salvar_rem_em_cadastramento(df_agg_por_rem)
    salvar_rem_pag_processado_zerado(df_agg_por_rb)

    df_em_revisao = df_agg_por_rem[df_agg_por_rem['situacao'] == 'EMREVISAO']
    df_liberado_pag = df_agg_por_rem[df_agg_por_rem['situacao'] == 'LIBPAGTO']
    df_pag_proc_zerado = df_agg_por_rb[
        (df_agg_por_rb['situacao'] == 'PAGPROC')
        & (df_agg_por_rb['valor_pagar'] == 0)
    ]

    salvar_monitoramento(df_em_revisao, df_liberado_pag, df_pag_proc_zerado)

    gerar_relatorios_regionais(
        df_em_revisao, df_liberado_pag, df_pag_proc_zerado
    )

    logger.info('Processamento concluído')


if __name__ == '__main__':
    main()
