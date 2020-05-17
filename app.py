import sys
import pandas as pd
import altair as alt
import streamlit as st
import numpy as np
from sqlalchemy import create_engine
engine = create_engine('sqlite://', echo=False)
def clean(grouped,table):
    DF = grouped.get_group(table)
    DF.dropna(axis=1, how='all',inplace=True)
    new_header = DF.iloc[0] #grab the first row for the header
    DF = DF[1:] #take the data less the header row
    DF.columns = new_header #set the header row as the df header
    del DF['%F']
    return DF
@st.cache
def load_data(uploaded_file):
    df = pd.read_csv(uploaded_file,sep='\t',names=range(100), encoding= 'unicode_escape',dtype=str)
    df.loc[df[0] == '%T', 'table'] = df[1]
    df['table'].fillna(method='ffill', inplace=True)
    data=df.loc[df[0].isin(['%R','%F'])]
    return data
st.title('Read XER')
uploaded_file = st.file_uploader("Choose an XER file", type="xer")
if uploaded_file is not None:
    dff=load_data(uploaded_file)
    tablelist= dff['table'].unique()
    tables =pd.DataFrame(tablelist)
    tables.columns = ['tables']
    grouped = dff.groupby(dff.table)
    df_list= {}
    #####################################
    for x in tablelist:
        df = clean(grouped,x)
        df.to_sql(x, con=engine)
PROJECT =pd.read_sql("SELECT proj_id,proj_short_name FROM PROJECT",engine) 
values = PROJECT['proj_short_name'].tolist()
options = PROJECT['proj_id'].tolist()
dic = dict(zip(options, values))
proj_id_var= st.sidebar.selectbox('Select Project', options, format_func=lambda x: dic[x])
###### Shwing sme stats about the file
result =pd.read_sql("SELECT count(*) as Number_task FROM TASK where proj_id="+proj_id_var,engine) 
st.subheader("show some stats about the file")  
st.write (result)