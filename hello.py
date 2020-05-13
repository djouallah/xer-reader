import sys
import pandas as pd
import altair as alt
import streamlit as st
import numpy as np
def clean(grouped,table):
    DF = grouped.get_group(table)
    DF.dropna(axis=1, how='all',inplace=True)
    new_header = DF.iloc[0] #grab the first row for the header
    DF = DF[1:] #take the data less the header row
    DF.columns = new_header #set the header row as the df header
    del DF['%F']
    return DF
    
st.title('Read XER')
uploaded_file = st.file_uploader("Choose an XER file", type="xer")
if uploaded_file is not None:
    df = pd.read_csv(uploaded_file,sep='\t',names=range(100), encoding= 'unicode_escape',dtype=str)
    df.loc[df[0] == '%T', 'table'] = df[1]
    df['table'].fillna(method='ffill', inplace=True)
    dff=df.loc[df[0].isin(['%R','%F'])]
    tablelist= dff['table'].unique()
    tables =pd.DataFrame(tablelist)
    tables.columns = ['tables']
    grouped = dff.groupby(dff.table)
    df_list= {}
    #####################################
    for x in tablelist:
        df = clean(grouped,x)
        df_name = { x: df}
        df_list.update(df_name)
project_name = df_list["PROJECT"]['proj_short_name']
st.table(project_name)
option = st.sidebar.selectbox('Select Tables?',tables['tables'])
TASK = df_list[option]
st.table(TASK)