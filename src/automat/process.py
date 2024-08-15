import pandas as pd
from openpyxl import *


#Selecciona el archivo a trabajar
input_cols = [0,1,5,6,9,13]
anexo_evento = "E:/PRUEBA/anexo1.xlsx"
dfanx = pd.read_excel(anexo_evento, header=0, usecols=input_cols)

#Maestra de tecnologias que son suceptibles de automatización
input_cols1 = [0]
anexo_evento = "E:/PRUEBA/TECN SUSCEPTIBLES AUTOMATIZACION BASE DEFINITIVA.xlsx"


#filtra tarifas con valor inferior a $3.000.000
dfanx = dfanx[dfanx["TARIFA NEGOCIADA*"] <= 3000000]

#Validar tecnologias Activas/Inactivas (Dejar solo activas)
dfanx = dfanx[dfanx["ESTADO"] == "Activo"]

#Quitar duplicados
dfanx = dfanx.drop_duplicates(['COD TECNOLOGIA* (RIPS)', 'DESC TECNOLOGIA*'], keep='last')
    
#Convierte a numero los valores para realizar los cruces de informacion.
def convertir_a_numero(value):
    try:
        return int(value)
    except ValueError:
        return value

dfanx['COD TECNOLOGIA* (RIPS)'] = dfanx['COD TECNOLOGIA* (RIPS)'].apply(convertir_a_numero)

#dfanx.to_excel("E:/PRUEBA/Nuevo.xlsx", header=True,index=False, sheet_name="FINAL")

#CRUCES DE INFORMACION

#cruce de anexo con maestra de tecnologias suseptibles para automatizacion para determinar las que NO se hacen.
df_auto = pd.read_excel(anexo_evento, header=0, sheet_name="NO SE HACE", usecols=input_cols1)
criterio_no = dfanx.merge(df_auto, how='left', left_on='COD TECNOLOGIA* (RIPS)', right_on='cod_tecnologia', indicator='Criterio_no')

#cruce de anexo con maestra de tecnologias suseptibles para automatizacion para determinar las que SI se hacen.
df_auto = pd.read_excel(anexo_evento, header=0, sheet_name="SI SE HACE", usecols=input_cols1)
criterio_si = dfanx.merge(df_auto, how='left', left_on='COD TECNOLOGIA* (RIPS)', right_on='cod_tecnologia', indicator='Criterio_si')

#Filtro de datos que no cruzan para definir que tecnologias no estan en la maestra de tecnologias suceptibles de automatización, ya sea en el "SI SE HACEN" o "NO SE HACEN", 
#El resultado se exporta para que sea analizado por el equipo de automatizaciones.

df_no = criterio_no[criterio_no['Criterio_no'] == 'left_only']
df_si = criterio_si[criterio_si['Criterio_si'] == 'left_only']

no_cruce = pd.merge(df_si,df_no, how='inner', on='COD TECNOLOGIA* (RIPS)')

#Filtro de datos que si cruzan, para determinar que tecnologias no se van a automatizar y cuales si.
df_filtro_no = criterio_no[criterio_no['Criterio_no'] == 'both']
df_filtro_si = criterio_si[criterio_si['Criterio_si'] == 'both']



#Exportar archivo excel resultado
with pd.ExcelWriter('E:/PRUEBA/Resultado.xlsx', engine='openpyxl') as writer:
    df_filtro_no.to_excel(writer, sheet_name='CRITERIO_NO', index=False)
    df_filtro_si.to_excel(writer, sheet_name='CRITERIO_SI', index=False)
    no_cruce.to_excel(writer, sheet_name='RESULTADO', index=False)



