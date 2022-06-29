import pandas as pd
from datetime import datetime

time_nombre = 'archivo.xlsx'
nombre_salida = datetime.today().strftime('%Y.%m.%d')+' Estado de stock general tek3.xlsx'
name_file = time_nombre
name_sheet = "Estado_General_de_Stock-25_05_2"
time_sheet = datetime.today().strftime('%d_%m_2')
print(time_nombre)
name_sheet ='Estado_General_de_Stock-'+f'{time_sheet}'


df = pd.read_excel(time_nombre, sheet_name=name_sheet,header=3)
pd.DataFrame(df)
df['Cant. Disponible'] = df['Cant. Disponible'].astype(int)
df['Cant. Pte. Liberación'] = df['Cant. Pte. Liberación'].astype(int)
df['Cant. en Simulación'] = df['Cant. en Simulación'].astype(int)

eit = df[['Cód. Item','Cant. Disponible','Cant. Pte. Liberación','Cant. en Simulación']]
def resta(i,a,p):
    return i - a - p

eit['dispEIT'] = eit.apply(lambda f: resta(f['Cant. Disponible'], f['Cant. Pte. Liberación'], f['Cant. en Simulación']), axis=1)
eitFinal = eit[['Cód. Item','dispEIT']]
writer = pd.ExcelWriter('Resultado_'+f'{nombre_salida}', engine='xlsxwriter')
eitFinal.to_excel(writer, sheet_name= "resumen",index=False)
df.to_excel(writer, sheet_name= "original",index=False)
#df.to_excel(f'{time}'+ " Estado de stock general tek3.xlsx",index=False, sheet_name= "original")
writer.save()
writer.close()