import pandas as pd
import matplotlib.pyplot as plt

# Pueden usarse bucles infinitos??? son realmente infinitos???
def menu():
    datos = lectura_archivo()
    datos = limpieza(datos)
    while True:
        opcion = (input("""Elige el numero de tu opcion:
        1- Influencer con mas seguidores.
        2- Los Tres paises con más Influencers.
        3- Cuenta de Marca con más seguidores.
        4- Salir
        """))

        match opcion :
            case '1':
                print("El top 5 Influencers con mas seguidores son:", end="\n")
                influencer(datos)
                

                # llama una funcion que entra o retorna un valor que se va la funcion deseada
            case '2':
                print("Los Tres paises con más Influencers son:", end="\n\n")
                top_paises(datos)
              
                # llama una funcion que sale o retorna un valor que se va la funcion deseada
            case '3':
                print("Cuenta de Marca con mas seguidores",  end="\n\n")
                marca(datos)
                
            case '4':
                print("Elegiste Salir")
                break
            case _:
                print("Error. Debes seleccionar una opción del 1 al 4")
                




def lectura_archivo():
    archivo = "C:\\Users\\jazmin\\Documents\\Otros programas\\Proyectos\\python\\cienda_datos_ifts\\Datos usuarios de instagram.xlsx"
    df_datos = pd.read_excel(archivo, sheet_name='Hoja1')
    df_datos["País"]=df_datos["País"].replace(u"\xa0", "")

    #Renombra encabezado
    df_datos=df_datos.rename(columns={'Cuenta de marca':'Marca'})
    return df_datos

 # Cambio de tipo de dato y encabezados 
def limpieza(df_datos):
    df_datos["Seguidores(millones)"]= pd.to_numeric(df_datos["Seguidores(millones)"],errors='coerce')
 
    return df_datos 

def influencer(df_datos):
    # 1 paso
    df_influencer = df_datos.iloc[:,[2,3,5]]
    # Organiza datos
    df_influencer.sort_values(by='Seguidores(millones)', ascending=False)
    # Filtra datos  top 5
    print(df_influencer.loc[0:5, ['Propietario',	'Seguidores(millones)', 'País']])
    top_5_influencer = df_influencer.loc[0:5, ['Propietario',	'Seguidores(millones)', 'País']]
    #top_5_influencer.plot(kind='bar', x='Propietario', y='Seguidores(millones)', colormap='viridis')
    #plt.show()
    top_5_influencer.plot(kind='scatter', x='País', y='Propietario', s=top_5_influencer['Seguidores(millones)'], c='Seguidores(millones)',colormap='viridis', alpha=0.5)
    plt.show()
    return df_influencer

# 2 paso
def top_paises(df_datos):
    df_paises = df_datos.iloc[:,[1,5]]
    df_paises = df_paises[['País', 'Usuario']].groupby('País')
    df_paises = df_paises.count().sort_values(by='Usuario', ascending=False).head(3)
    print(df_paises)
    df_paises.plot(kind='barh', y='Usuario',  colormap='viridis')
    plt.show()
    return df_paises

# 3 paso

def marca(df_datos): 

    df_marca = df_datos.iloc[:,[2,3,6]]
    df_marca = df_marca.loc[(df_marca.Marca == 'Sí')].sort_values(by='Seguidores(millones)', ascending=False)
    print(df_marca)
    df_marca.set_index('Propietario').plot(kind='pie', y='Seguidores(millones)', title='Propietarios que son Marca' ,colormap='viridis', figsize=(5, 5), legend=False, autopct='%1.1f%%', shadow=True)
    plt.show()
    return df_marca

if __name__ == '__main__':
    menu()