# Librerías a utilizar
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
import random
import string
import base64

# Configuración de la página
st.set_page_config(page_title="Cálculo de Matriz Ybus", layout="wide")

# Personalización de estilos
st.markdown("""
    <style>
    .css-18e3th9, .css-1y4v4l9, .css-1v0mbdj {background-color: lavenderblush; color: black;}
    .css-1l0l5lz {color: white;}
    .title {text-align: center; font-size: 40px; font-weight: bold; color: deeppink;}  
    [data-testid="stSidebar"] {background-color: mistyrose;}
    .subheader {text-align: center; font-size: 20px; font-weight: bold; color: palevioletred;}
    </style>
    """, unsafe_allow_html=True)


# Función para calcular la matriz Ybus
def calcular_matriz_ybus(data, data_generadores):
    data = data.dropna(how="any")
    nodos = sorted(set(data['Nodo origen']).union(set(data['Nodo destino'])))
    n = len(nodos)
    
    nodo_indices = {nodo: idx for idx, nodo in enumerate(nodos)}
    Ybus = np.zeros((n, n), dtype=complex)
    
    # Procesar admitancias de líneas
    for _, row in data.iterrows():
        i, j = nodo_indices[row['Nodo origen']], nodo_indices[row['Nodo destino']]
        Y = complex(row['Conductancia de la línea'], row['Susceptancia de la línea'])  # Solo X en la parte imaginaria para fuera de la diagonal
        Y_shunt = complex(0, row['(Y/2)'])  # Solo en la diagonal principal
        
        Ybus[i, j] -= Y
        Ybus[j, i] -= Y
        Ybus[i, i] += Y + Y_shunt
        Ybus[j, j] += Y + Y_shunt
    
    # Agregar admitancias de generadores
    for _, row in data_generadores.iterrows():
        if pd.notna(row['Conductancia del generador']) and pd.notna(row['Susceptancia del generador']):  
            Y_gen = complex(row['Conductancia del generador'], row['Susceptancia del generador'])
            i = nodo_indices[row['Nodo']]
            Ybus[i, i] += Y_gen
    
    return np.round(Ybus, 6), nodos

# Función para exportar matriz Ybus a Excel
def generar_nombre_aleatorio():
    return "matriz_Ybus_" + ''.join(random.choices(string.ascii_letters + string.digits, k=8)) + ".xlsx"

def exportar_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=True, sheet_name='Matriz Ybus')
    output.seek(0)
    return output

# Botón para descargar el archivo con nombre aleatorio
if 'Ybus' in locals():
    # Botón para descargar matriz Ybus en Excel
    st.download_button(
        label="Descargar matriz Ybus en Excel",
        data=exportar_excel(df_Ybus),
        file_name="matriz_Ybus.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Configuración de la interfaz
st.sidebar.text('Seleccione una opción:')

if 'pagina' not in st.session_state:
    st.session_state.pagina = "Cálcular matriz"

if st.sidebar.button("Cálcular matriz"):
    st.session_state.pagina = "Cálcular matriz"
if st.sidebar.button("Creadores"):
    st.session_state.pagina = "Creadores"

# Página principal
if st.session_state.pagina == "Cálcular matriz":
    
    st.markdown("<h1 class='title'>Cálculo de Matriz de Admitancia Nodal (Ybus)</h1>", unsafe_allow_html=True)

    st.markdown("### Ingresar los valores de la admitancia de cada línea.")
   
   ## Ingreso de datos de impedancias
    data = pd.DataFrame({
       'Nodo origen': pd.Series(dtype='int'),
        'Nodo destino': pd.Series(dtype='int'),
        'Conductancia de la línea': pd.Series(dtype='float'),
        'Susceptancia de la línea': pd.Series(dtype='float'),
        '(Y/2)': pd.Series(dtype='float'),
    })
    
    data = st.data_editor(data, num_rows="dynamic", key="tabla_datos", use_container_width=True)

    # Ingreso de datos de generadores
    st.markdown("### Ingrese los valores de la admitancia de los elementos aislados:")
    data_generadores = pd.DataFrame(columns=['Nodo', 'Conductancia del generador', 'Susceptancia del generador'])

    data_generadores = st.data_editor(data_generadores, num_rows="dynamic", key="tabla_generadores", use_container_width=True)
    
    if st.button("Calcular matriz Ybus"):

        data = data.dropna(how="any")

        if data.empty:
            st.warning("La tabla está vacía o contiene filas incompletas. Por favor, revise los datos.")
        else:
            # Asegurar que la columna 'Nodo' existe en data_generadores
            if 'Nodo' not in data_generadores.columns:
                data_generadores['Nodo'] = None

            # Asegurar que 'Nodo' existe en data
            data['Nodo'] = data['Nodo origen']

            # Convertir la parte real e imaginaria en número complejo
            if not data_generadores.dropna(how="any").empty:
                # Convertir las columnas de generadores a tipo numérico, manejando errores
                data_generadores['Conductancia del generador'] = pd.to_numeric(data_generadores['Conductancia del generador'], errors='coerce').fillna(0)
                data_generadores['Susceptancia del generador'] = pd.to_numeric(data_generadores['Susceptancia del generador'], errors='coerce').fillna(0)

                # Crear la columna Y_gen como número complejo
                data_generadores['Y_gen'] = data_generadores['Conductancia del generador'] + 1j * data_generadores['Susceptancia del generador']

            else:
                data_generadores['Y_gen'] = None

            # Eliminar valores NaN en 'Nodo' de data_generadores
            data_generadores = data_generadores.dropna(subset=['Nodo'])

            # Convertir a tipo entero
            data_generadores['Nodo'] = data_generadores['Nodo'].astype(int)

            # Hacer el merge asegurando que 'Nodo' es int en ambas tablas
            data_final = data.merge(data_generadores[['Nodo', 'Y_gen']], on='Nodo', how='left')

            # Calcular matriz Ybus
            Ybus, nodos = calcular_matriz_ybus(data, data_generadores)

            st.success("Matriz Ybus calculada correctamente:")
            
            # Mostrar la matriz Ybus
            df_Ybus = pd.DataFrame(Ybus, index=[f"Nodo {n}" for n in nodos], columns=[f"Nodo {n}" for n in nodos])
            st.dataframe(df_Ybus, use_container_width=True)
    
            # Botón para descargar matriz Ybus en Excel
            st.download_button(
                label="Descargar matriz Ybus en Excel",
                data=exportar_excel(df_Ybus),
                file_name="matriz_Ybus.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
                

elif st.session_state.pagina == "Creadores":
    st.markdown("<h1 class='title'>Autores</h1>", unsafe_allow_html=True)
    st.write("Este programa fue desarrollado por el equipo de ingeniería eléctrica de la Universidad del Norte conformado por:")
    st.markdown("""
    - Isabella María Chacón Villa
    - Juan Camilo Pombo Muñoz
    """)


