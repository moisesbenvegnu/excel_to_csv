# -*- coding: utf-8 -*-
"""
Created on Fri Jul 15 09:22:25 2022

@author: Moises
"""

import streamlit as st
import pandas as pd
import openpyxl

st.title('Conversor Excel para CSV')

st.image('capa.png')

uploaded_file = st.file_uploader("Faça o upload do arquivo em excel clicando abaixo")

if uploaded_file is not None:
    
    
    planilha = pd.read_excel(uploaded_file, sheet_name = None)

    abas_dados = []
    for aba in planilha.values():
        abas_dados.append(aba)
    abas_titulos = []
    for titulo in planilha:
        abas_titulos.append(titulo)

    st.markdown('**O Arquivo em excel possui as seguintes abas que serão convertidas em arquivos CSV:**')

    st.info(abas_titulos)

    st.markdown('**Seguem abaixo os arquivos em CSV para download:**')

    for i in range(len(abas_dados)):

        st.download_button(
             label='Download ' + abas_titulos[i] + '.csv',
             data=abas_dados[i].to_csv(index = False, sep = ';').encode('utf-8'),
             file_name = abas_titulos[i] + '.csv',
             mime='text/csv',
         )


        if st.button('Pré-visualizar o arquivo ' + abas_titulos[i] + '.csv'):
            try:
                st.dataframe(abas_dados[i])
            except:
                st.info('Não foi possível pré-visualizar este arquivo')

    #st.info('Não foi possível realizar a conversão. Verifique se o arquivo utilizado atende aos requisitos para a conversão.')

st.markdown('Desenvolvido por Moises A. Benvegnu')
