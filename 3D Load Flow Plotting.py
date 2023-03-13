import os
import sys
import pandas as pd
import numpy as np
import plotly.graph_objects as go

## Inputs
Excel_file = '3D_Loadflow.xlsx'
sheet_name = ''  # En caso de no especificar nombre de la hoja, introducir 0
maximum_loadflow_limit = 600  # MVA o % sobrecarga
days_range = [1,365]  #  Se introduce un rango de valores. El primer valor nunca debe ser 0, como mínimo debe ser 1. El segundo valor puede ser como máximo 365 (o 366 en año bisiesto)
name = ''  # Nombre del elemento a estudiar
factor_escala_eje_x = 4 # Con este parámetro se modifica la longitud del eje X. Si es 1, la gráfica será un cubo; si introduce otro valor, el eje X será 'n' veces el eje Y y Z


## Lectura del Excel --> Creacion DataFrame
Excel_file_pd = pd.read_excel(Excel_file)

## Creacion de los vectores de lso ejes X,Y,Z del plot

Active_Power = []  # Eje Z
hours = []  # Eje Y
days = []  # Eje X

number_days = (days_range[1] - days_range[0])+1

for day in range(number_days):  # Creación Eje X
    real_day = days_range[0] + day
    days.append(real_day)


for hour in range(24):   # Creación Eje Y
    hours.append(hour)


for hour in range(24):  # Creación Eje Z
    vector_z = []
    for day in range(number_days):
        real_day = days_range[0] + day
        real_hour = (hour + (real_day*24))-24
        vector_z.append(Excel_file_pd[Excel_file_pd.columns[0]].iloc[real_hour])
    Active_Power.append(vector_z)

        
## Pasamos de Lista a Numpy Array
X_array = np.array(days)
Y_array= ['23','22','21','20','19','18','17','16','15','14','13','12','11','10','9','8','7','6','5','4','3','2','1','0']
X_array, Y_array = np.meshgrid(X_array, Y_array)
Z_array = np.array(Active_Power)
Z_array_invertido=Z_array[::-1,:]


## Meter plano limite
X_plano = np.arange(days_range[0]-1, days_range[1]+2, 1)
Y_plano = np.arange(0, 24, 1)
X_plano, Y_plano = np.meshgrid(X_plano, Y_plano)
Z_plano = (X_plano + Y_plano)*0 + maximum_loadflow_limit

Y_plano= ['23','22','21','20','19','18','17','16','15','14','13','12','11','10','9','8','7','6','5','4','3','2','1','0']

#### Graficar en 3D
fig = go.Figure(data=[
    go.Surface(z=Z_plano ,x= X_plano,y=Y_plano, showscale=False, colorscale='darkmint', opacity=0.75),
    go.Surface(z=Z_array_invertido ,x= X_array,y=Y_array, colorscale='YlOrRd')])
fig.update_layout(title={"text": 'Flujo de cargas horario de: '+ name, "y": 0.9, "x": 0.5, "xanchor": "center", "yanchor": "top", "font_size": 28},
                  margin={"l": 0, "r": 0, "t": 0, "b": 0},
                  scene = dict(
                    xaxis_title='Dia', xaxis_title_font_size=18,xaxis_ticks='outside',
                    yaxis_title='Hora', yaxis_title_font_size=18,yaxis_ticklen=24,
                    yaxis_tickvals = np.arange(0, 24, 2),
                    zaxis_title='Potencia Activa (MW)', zaxis_title_font_size=18,
                    aspectratio_x=factor_escala_eje_x, aspectratio_y=1,aspectratio_z=1))

        

fig.show() 
    
