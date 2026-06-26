import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Consolidador de Extratos", layout="wide")

st.title("🏦 Conversor de Extratos Bancários")
st.markdown("Faça upload do extrato bancário em Excel e baixe o arquivo agrupado e formatado para contabilização.")

uploaded_file = st.file_uploader("📎 Selecione o arquivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:

        # ---------------------------------------------------
        # LEITURA DO ARQUIVO
        # ---------------------------------------------------

        df_raw = pd.read_excel(uploaded_file, header=None, dtype=str, engine='openpyxl')

        linha_cabecalho = None
        for idx, row in df_raw.iterrows():
            if row.astype(str).str.strip().str.lower().isin(['deb/credit']).any():
                linha_cabecalho = idx
                break

        if linha_cabecalho is None:
            st.error("❌ Cabeçalho com 'Deb/Credit' não encontrado no arquivo.")
            st.stop()

        df = pd.read_excel(uploaded_file, header=linha_cabecalho, dtype=str, engine='openpyxl')
        df.columns = df.columns.str.strip()
        df = df.fillna('')

        # ---------------------------------------------------
        # FILTROS INICIAIS
        # ---------------------------------------------------

        df = df[df['Deb/Credit'] == "Credito"]

        historico_filters = [
            'BIN','BANRISUL','CREDZ','ELOSGATE','GETNET','GLOBAL','CIELO','REDE',
            'CONTAS A RECEBER TRANSI','STONE','PAGSEGURO','FISERV','PAGSEG','SISPAG','SFPAY','PIX TRANSF  Nu Pay', 'VERO BANRI',
        ]

        documento_filters = ['12109247','FISERV','REDE-','CIELO']

        df_filtered = df[
            df['Historico'].str.contains('|'.join(historico_filters), na=False) |
            df['Documento'].str.contains('|'.join(documento_filters), na=False)
        ]

        df_filtered = df_filtered[~df_filtered['Historico'].str.contains('MORAIS', na=False)]

        # ---------------------------------------------------
        # CÓDIGOS MACAÉ / BAURU
        # ---------------------------------------------------

        codigos_macae = [
            '91046446','91046449','2808379700','2808379697',
            '12627602','12627703','191807527','191807614'
        ]

        codigos_bauru = [
            '91270743','91270749','2808377759','2808377740',
            '87807580','87808153','12633893','12651489',
            '86571982','86572679'
        ]

        mask_macae = df_filtered['Historico'].str.contains('|'.join(codigos_macae), na=False)
        mask_bauru = df_filtered['Historico'].str.contains('|'.join(codigos_bauru), na=False)

        # ---------------------------------------------------
        # LIMPEZA DOS DADOS
        # ---------------------------------------------------

        df_filtered['Agencia'] = df_filtered['Agencia'].apply(lambda x: str(x)[-4:] if x else x)
        # conta que não pode perder formatação
        conta_excecao = '0610897103'

        df_filtered['Conta'] = df_filtered['Conta'].astype(str).str.strip()

        df_filtered['Conta'] = df_filtered['Conta'].apply(
        lambda x: x if x == conta_excecao else str(int(float(x))) if x.replace('.', '', 1).isdigit() else x
        ) 
        df_filtered['Filial'] = df_filtered['Filial'].apply(lambda x: str(x)[:4] if x else x)

        df_filtered['Ocorrencia'] = df_filtered['Ocorrencia'].fillna("N/A")
        df_filtered['Data'] = pd.to_datetime(df_filtered['Data'], errors='coerce')
        df_filtered['Valor'] = pd.to_numeric(df_filtered['Valor'], errors='coerce').fillna(0).round(2)

        # ---------------------------------------------------
        # IDENTIFICAÇÃO DO CANAL
        # ---------------------------------------------------

        def get_natureza(historico, ocorrencia, documento):

            if 'BANRISUL' in historico: return 'BANRISUL'
            elif 'BIN' in historico: return 'BIN'
            elif 'CREDZ' in historico or '12109247000120' in documento: return 'CREDZ'
            elif 'GETNET' in historico: return 'GETNET'
            elif 'GLOBAL' in historico: return 'GLOBAL'
            elif 'CIELO' in historico or 'CIELO' in documento: return 'CIELO'
            elif 'REDE' in historico or 'REDE' in documento: return 'REDE'
            elif 'VERO' in ocorrencia: return 'BIN'
            elif 'PAGSEGURO' in historico: return 'PAGSEGURO'
            elif 'PAGSEG' in historico: return 'TEDPAGSEG'
            elif 'FISERV' in historico or 'FISERV' in documento: return 'BIN'
            elif 'STONE' in historico: return 'STONE'
            elif 'SISPAG' in historico: return 'BIN'
            elif 'VERO BANRI' in historico: return 'VERO'    
            elif 'PIX TRANSF  Nu Pay' in historico: return 'NUPAY'

            return None

        df_filtered['Historico'] = df_filtered.apply(
            lambda row: get_natureza(row['Historico'], row['Ocorrencia'], row['Documento']),
            axis=1
        )

        # ---------------------------------------------------
        # MAPA DE NATUREZA CONTÁBIL
        # ---------------------------------------------------

        natureza_map = {
            'BANRISUL':'A10801',
            'BIN':101113,
            'CREDZ':101115,
            'GETNET':101112,
            'GLOBAL':'A10806',
            'CIELO':101118,
            'REDE':101111,
            'TEDPAGSEG':101117,
            'STONE':101122,
            'PAGSEGURO':101117,
            'SISPAG PAGSEG':101117,
            'NUPAY':101121,
            'VERO':100116,
        }

        df_filtered['Natureza'] = df_filtered['Historico'].map(natureza_map)

        # ---------------------------------------------------
        # AJUSTE MACAÉ / BAURU
        # ---------------------------------------------------

        df_filtered.loc[mask_macae | mask_bauru, 'Natureza'] = 100135

        df_filtered.loc[mask_macae, 'Historico'] = df_filtered.loc[mask_macae, 'Historico'] + ' - MACAE'
        df_filtered.loc[mask_bauru, 'Historico'] = df_filtered.loc[mask_bauru, 'Historico'] + ' - BAURU'

        # ---------------------------------------------------
        # AGRUPAMENTO
        # ---------------------------------------------------

        df_grouped = df_filtered.groupby(
            ['Filial','Data','Historico','Natureza','Banco','Agencia','Conta']
        ).agg({'Valor':'sum'}).reset_index()

        # Ajuste de formato da data
        df_grouped['Data'] = df_grouped['Data'].dt.strftime('%d/%m/%Y')

        # ---------------------------------------------------
        # COMPLEMENTO CONTÁBIL
        # ---------------------------------------------------

        df_grouped['TIPO'] = 'R'
        df_grouped['NUMERARIO'] = 'CD'
        df_grouped['NUM CHEQUE'] = ''
        df_grouped['C. Custo debito'] = ''
        df_grouped['C. Custo credito'] = ''
        df_grouped['Item debito'] = ''
        df_grouped['Item credito'] = ''
        df_grouped['Cl Valor deb'] = ''
        df_grouped['Cl Valor crd'] = ''

        colunas_ordenadas = [
            'Filial','Data','NUMERARIO','TIPO','Valor','Natureza','Banco','Agencia','Conta',
            'NUM CHEQUE','Historico','C. Custo debito','C. Custo credito',
            'Item debito','Item credito','Cl Valor deb','Cl Valor crd'
        ]

        df_grouped = df_grouped[colunas_ordenadas]

        

        # ---------------------------------------------------
        # PREVIEW
        # ---------------------------------------------------

        st.subheader("🔎 Prévia do resultado")

        st.dataframe(
            df_grouped.head(50),
            use_container_width=True
        )

        st.caption(f"Mostrando 50 linhas de um total de {len(df_grouped)} registros.")

        # ---------------------------------------------------
        # DOWNLOAD EXCEL
        # ---------------------------------------------------

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_grouped.to_excel(writer, index=False)

        output.seek(0)

        st.success("✅ Arquivo processado com sucesso!")

        st.download_button(
            label="⬇️ Baixar Excel Formatado",
            data=output,
            file_name=f"consolidado_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")
