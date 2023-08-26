import streamlit as st
import sys
import streamlit as st
import time
import numpy as np
from PIL import Image
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder
from st_aggrid.shared import GridUpdateMode
import plotly.express as px
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# Your app code goes here
@st.cache_data()
def load_data():
    df =  pd.read_excel ('bdd.xlsx')
    df=df.drop(['Region', 'City','Area','District ID', 'District Name', 'Salesman Name', 'Customer No','BUID', 'Points','Price', 'Value', 'Discount','Qty'], axis=1)
    return df





# try both connected=False (the default) and connected=True
st.title('SARL ANDROMED DISTRIBUTION -ALGER CENTRE- 2023')
st.text('Application Créée Par: ALLOUCHE KENZA')

st.header('DATA BASE:')
uploaded_files = st.file_uploader("Choose a CSV file", accept_multiple_files=True)
for uploaded_file in uploaded_files:
    bytes_data = uploaded_file.read()
    st.write("filename:", uploaded_file.name)
    st.write(bytes_data)



df= load_data()
dfg=df.groupby(['Salesman No']).sum()
df2=df.groupby(['Item Name','Item ID','Salesman No','Customer Name'], as_index=False).sum()
#df2['OBJECTIF CA HT']=df2['Net']-50
#df2['Realisation']=df2['Net']-df2['OBJECTIF CA HT']
#df3=pd.read_excel('ff.xlsx')
#df3

# Add column using np.where()
#df2['Discount_rating'] = df2['OBJECTIF CA HT'].where(df2['Item Name'].isin(df3['Item Name'])& df2['Salesman No'].isin(df3['Salesman No']))
#print(df2)

gb = GridOptionsBuilder.from_dataframe(
        df2, enableRowGroup=True, enableValue=True, enablePivot=True
    )
#gb.configure_pagination(enabled=True) #Add pagination

#gb = GridOptionsBuilder.from_dataframe(df)
gb.configure_pagination(paginationAutoPageSize=True, paginationPageSize=50) #Add pagination
gb.configure_side_bar() #Add a sidebar
#gb.configure_selection( groupSelectsChildren = "Group checkbox select children") #Enable multi-row selection

gridOptions = gb.build()

grid_response = AgGrid(
    df2,
    gridOptions=gridOptions,
    data_return_mode='AS_INPUT', 
    update_mode='MODEL_CHANGED', 
    fit_columns_on_grid_load=False,
    theme = 'streamlit',
   # theme='streamlit',# ['streamlit', 'alpine', 'balham', 'material']
    #height=900,
    #width=5,
    
    
    enable_enterprise_modules=True,
     
    
    reload_data=True
)

data = grid_response['data']
selected = grid_response['selected_rows'] 
df = pd.DataFrame(selected)

