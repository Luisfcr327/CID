#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd     
import pyodbc
from datetime import date, timedelta
import calendar
import statistics
import warnings
import schedule
import time
import tkinter as tk
import win32com.client as win32
from tkinter import simpledialog

from tkinter import ttk
from tkinter import messagebox
from tkcalendar import DateEntry


# In[2]:


#Objeto tienda y sus caracteristicas
class Tienda():
    def __init__(self, tienda):
        self.numero_tienda = tienda
        self.socio = self.definir_socios(tienda)
        self.stock = 0,
        #self.consumo_diario = [1,2,3]
        self.promedio = 0,
        self.tienda_vision = 00000000
        self.sucursal = ""

    def definir_socios(self, tienda):
        socio_dict = {
            '00': 'CyA',
            '01': 'BRADESCARD',
            '02': 'LOB',
            '03': 'SUBURBIA',
            '04': 'AURRERA',
            '05': 'SHASA',
            '06': 'GCC',
            '07': 'PROMODA',
            '08': 'CCP'
        }
        socio = tienda[:2]
        return socio_dict.get(socio, 'OTRO')
            
    def agregar_stock(self,stock_agregado):
        self.stock = self.stock + stock_agregado

    def quitar_stock(self,stock_consumido):
        #self.consumo_diario.append(stock_consumido)
        self.stock = self.stock - stock_consumido
    
    def ajustar_stock(self,nuevo_stock):
        self.stock = nuevo_stock

    #def calcular_promedio(self):
    #    self.promedio = statistics.mean(self.consumo_diario)
        
    def asignar_promedio(self, promedio):
        self.promedio = promedio

    def get_stock(self):
        return self.stock
        
    def get_numero_tienda(self):
        return self.numero_tienda
        
    def get_socio(self):
        return self.socio

    def __str__(self):
        return f"Tienda: {self.numero_tienda} \n Socio: {self.socio} \n Stock: {self.stock} \n "

    def to_dict(self):
        return {
            "Tienda": self.numero_tienda,
            "Socio": self.socio,
            "Stock" : self.stock,
            #"Consumo diario": ", ".join(map(str, self.consumo_diario)),
            "Promedio diario" : self.promedio,
            "Tienda_Vision" : self.tienda_vision,
            "Sucursal" : self.sucursal
        }


# In[3]:


#AQUI ESTA TODO LO DE LA BASE DE DATOS
def abrir_db(servidor, db):
#Conexion a la base de datos
    server = servidor
    database = db


    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"Trusted_Connection=yes;"
    )

    try:
        connection=pyodbc.connect(conn_str)
        #print("Conexion exitosa")
    except Exception as e:
        messagebox.showerror(
            "Error Conexion DB",
            f"No se pudo conectar a la base de datos"
        )
            
    return connection
                    
def ejecutar_query(query,connection):
    #Creamos el objeto cursor para poder ejecutar consultas
    cursor = connection.cursor()

    # Realizar las consultas y el resultado lo guardamos en otro dataframe
    df = pd.read_sql(query, connection)
        
    return df
    
def cerrar_db(connection):
    # Cerrar la conexión
    connection.close()
    
    return connection


# In[4]:


#QUERY'S
def query_bid(fecha_inicio, fecha_final):
    
    query = f"""
    SELECT 
        NumeroTienda,
        COUNT(*) AS CantidadRegistros
    FROM [Tarjeta_Digital].[dbo].[TD_Embozado]
    WHERE FechaEmbozado >= {fecha_inicio} 
      AND FechaEmbozado < {fecha_final}
      AND (StatusEmbozado = 'Embozado con éxito' OR StatusEmbozado = 'Mal embozado' or StatusEmbozado IS null )
    GROUP BY NumeroTienda
    ORDER BY CantidadRegistros DESC;    
    """
    return query

def query_soc(fecha_inicio, fecha_final):

    query = f"""
    SELECT 
        Socio_2 + '0000' + Dsc_Tienda_C AS NumeroTienda,
        COUNT(*) AS CantidadRegistros
    FROM (
        SELECT 
            RIGHT('000' + CAST(CAST(b.Dsc_Tienda AS INT) AS VARCHAR), 3) AS Dsc_Tienda_C,
            RIGHT('0' + CAST(CAST(b.Id_Socio AS INT) AS VARCHAR), 2) AS Socio_2
        FROM Tabla_Bitacora_InventarioEmb a
        LEFT JOIN Cat_Sucursal_Socio b 
            ON a.tienda = b.Id_Sucursal
        WHERE fecha > {fecha_inicio} AND fecha < {fecha_final}
          AND (estatus = '0' OR estatus = '1' OR estatus IS null)
    ) t
    GROUP BY 
        Socio_2 + '0000' + Dsc_Tienda_C
    ORDER BY 
        CantidadRegistros DESC;
        
    """

    return query

def query_plasticos_bid(fecha_inicio, fecha_final):

    query = f"""
    SELECT [NumeroTienda],
          [Socio],
          [StatusEmbozado],
          [FechaEmbozado]
      FROM [Tarjeta_Digital].[dbo].[TD_Embozado]
    
    WHERE FechaEmbozado >= {fecha_inicio}
    AND FechaEmbozado < {fecha_final}

    """

    return query

def query_plasticos_soc(fecha_inicio, fecha_final):
    
    query = f"""
    SELECT 
    
    	Socio_2 + '0000' + Dsc_Tienda_C AS NumeroTienda,
    	
    	CASE 
            WHEN Id_Socio = 0 THEN 'CyA'
            WHEN Id_Socio = 1 THEN 'BRADESCARD'
            WHEN Id_Socio = 2 THEN 'LOB'
            WHEN Id_Socio = 3 THEN 'SUBURBIA'
            WHEN Id_Socio = 4 THEN 'AURRERA'
            WHEN Id_Socio = 5 THEN 'SHASA'
            WHEN Id_Socio = 6 THEN 'GCC'
            WHEN Id_Socio = 7 THEN 'PROMODA'
            WHEN Id_Socio = 8 THEN 'CCP'
        END AS Socio,
    
    	Case
    		WHEN estatus = '0' THEN 'Mal embozado'
    		WHEN estatus = '1' THEN 'Embozado con éxito'
    		WHEN estatus = '2' THEN 'No tomo el plastico'
    		WHEN estatus IS NULL THEN 'NO_ESPEC'
    	END AS StatusEmbozado,
    
    
    	[fecha] AS FechaEmbozado
    
    FROM (
        SELECT 
            RIGHT('000' + CAST(CAST(b.Dsc_Tienda AS INT) AS VARCHAR), 3) AS Dsc_Tienda_C,
            RIGHT('0' + CAST(CAST(b.Id_Socio AS INT) AS VARCHAR), 2) AS Socio_2,
    		[estatus],
    		[Id_Socio],
    		[fecha]
        FROM Tabla_Bitacora_InventarioEmb a
        LEFT JOIN Cat_Sucursal_Socio b 
            ON a.tienda = b.Id_Sucursal
        WHERE fecha > {fecha_inicio} AND fecha < {fecha_final}
    ) t
    """

    return query
    

def query_emozados_soc(fecha_inicio, fecha_final):
    query = f"""

    select cast(fecha as date) as fecha, cast(Dsc_Tienda as int) as Tienda ,Dsc_Socio,Dsc_Socio + ' ' + Dsc_Producto as Producto, count(*) as tarjetas_usadas from Tabla_Bitacora_InventarioEmb a left join Cat_Sucursal_Socio b on tienda = Id_Sucursal 
    left join Cat_Socios c on b.Id_Socio = c.Id_Socio 
    left join Cat_Producto d on a.producto = d.Id_Producto
    
    where fecha >= {fecha_inicio}
    and fecha < {fecha_final}
    and b.Id_Socio not in ( '1','5')
    and (estatus <> '2' or estatus is null)
    
    group by cast(fecha as date),cast(Dsc_Tienda as int), Dsc_Socio,Dsc_Socio + ' ' + Dsc_Producto
    order by cast(fecha as date),cast(Dsc_Tienda as int),Dsc_Socio,Dsc_Socio + ' ' + Dsc_Producto   
    """

    return query

def query_embozados_bid(fecha_inicio, fecha_final):

    query = f"""

    select cast(fechaEmbozado as date) as fecha, cast(SUBSTRING(NumeroTienda,3,10) as int) as tienda,Socio,
    CASE WHEN socio + ' '+ producto = 'AURRERA BANKCARD' THEN 'AURRERA BK LB' 
    WHEN socio + ' '+ producto = 'AURRERA PLCC' THEN 'AURRERA PLCC LB' 
    WHEN socio + ' '+ producto = 'GCC BANKCARD' THEN 'GCC BK LB' 
    WHEN socio + ' '+ producto = 'GCC PLCC' THEN 'GCC PLCC LB' 
    WHEN socio + ' '+ producto = 'PROMODA PLCC' THEN 'PROMODA PLCC LB' 
    WHEN socio + ' '+ producto = 'PROMODA BANKCARD' THEN 'PROMODA BK LB'  else 'REVISAR' END as [Producto],
     count (*) as tarjetas_usadas from TD_Embozado
    
    where FechaEmbozado >= {fecha_inicio}
    and FechaEmbozado < {fecha_final}
    and (StatusEmbozado <> 'No tomó el plástico' or StatusEmbozado is null)
    group by  cast(fechaEmbozado as date), cast(SUBSTRING(NumeroTienda,3,10) as int),Socio,socio +' '+ producto
    order by cast(fechaEmbozado as date), cast(SUBSTRING(NumeroTienda,3,10) as int),Socio,socio +' '+ producto
    
    """
    return query    


# In[5]:


#Aqui van a estar los df que se generan de la base de datos
def bid():

    fecha_hoy = "'" + date.today().strftime("%Y-%m-%d") + "'"
    fecha_ant = "'" + (date.today() - timedelta(days=1)).strftime("%Y-%m-%d") + "'"
    
    conn = abrir_db("MP-VW-DB-017", "Tarjeta_Digital")
    df_bid = ejecutar_query(query_bid(fecha_ant, fecha_hoy), conn)
    cerrar_db(conn)
    
    return df_bid

def soc():

    fecha_hoy = "'" + date.today().strftime("%Y-%m-%d") + "'"
    fecha_ant = "'" + (date.today() - timedelta(days=1)).strftime("%Y-%m-%d") + "'"
    
    conn = abrir_db("MP-VW-DB-023", "ORIGINACION_MIGRACION")
    df_soc = ejecutar_query(query_soc(fecha_ant, fecha_hoy), conn)
    cerrar_db(conn)
    
    return df_soc

def bid_consumo(fecha_inicio, fecha_final):

    #Lo comentado es que toma en cuenta todo el mes
    #Pero ahora lo voy a cambiar para que tome un rango de fechas que quiera el usuario
    
    #hoy = date.today()
    
    # Primera fecha del mes
    #fecha_inicio = "'" + date(hoy.year, hoy.month, 1).strftime("%Y-%m-%d") + "'"
    
    # Última fecha del mes (como ya lo tienes)
    #ultimo_dia = calendar.monthrange(hoy.year, hoy.month)[1]
    #fecha_final = "'" + date(hoy.year, hoy.month, ultimo_dia).strftime("%Y-%m-%d") + "'"


    #f_inicio = date.fromisoformat(fecha_inicio)
    #f_final = date.fromisoformat(fecha_final)
    
    #dif_dias = (f_final - f_inicio).days

    fecha_inicio = "'" + fecha_inicio + "'"
    fecha_final = "'" + fecha_final + "'"
    
    conn = abrir_db("MP-VW-DB-017", "Tarjeta_Digital")
    df_bid = ejecutar_query(query_bid(fecha_inicio, fecha_final), conn)
    cerrar_db(conn)

    #df_bid["Promedio"] = df_bid["CantidadRegistros"] / dif_dias
    
    return df_bid

def soc_consumo(fecha_inicio, fecha_final):

    #Lo comentado es que toma en cuenta todo el mes
    #Pero ahora lo voy a cambiar para que tome un rango de fechas que quiera el usuario
    
    #hoy = date.today()
    
    # Primera fecha del mes
    #fecha_inicio = "'" + date(hoy.year, hoy.month, 1).strftime("%Y-%m-%d") + "'"
    
    # Última fecha del mes (como ya lo tienes)
    #ultimo_dia = calendar.monthrange(hoy.year, hoy.month)[1]
    #fecha_final = "'" + date(hoy.year, hoy.month, ultimo_dia).strftime("%Y-%m-%d") + "'"

    #f_inicio = date.fromisoformat(fecha_inicio)
    #f_final = date.fromisoformat(fecha_final)
    
    #dif_dias = (f_final - f_inicio).days
    
    fecha_inicio = "'" + fecha_inicio + "'"
    fecha_final = "'" + fecha_final + "'"
    
    conn = abrir_db("MP-VW-DB-023", "ORIGINACION_MIGRACION")
    df_soc = ejecutar_query(query_soc(fecha_inicio, fecha_final), conn)
    cerrar_db(conn)
    
    #df_soc["Promedio"] = df_soc["CantidadRegistros"] / dif_dias
    
    return df_soc

    
def embozado_bid(fecha_inicio, fecha_final):

    fecha_inicio = "'" + fecha_inicio + "'"
    fecha_final = "'" + fecha_final + "'"
    
    conn = abrir_db("MP-VW-DB-017", "Tarjeta_Digital")
    df_embo_bid = ejecutar_query(query_embozados_bid(fecha_inicio, fecha_final), conn)
    cerrar_db(conn)
    
    return df_embo_bid

def embozado_soc(fecha_inicio, fecha_final):

    fecha_inicio = "'" + fecha_inicio + "'"
    fecha_final = "'" + fecha_final + "'"
    
    conn = abrir_db("MP-VW-DB-023", "ORIGINACION_MIGRACION")
    df_embo_soc = ejecutar_query(query_emozados_soc(fecha_inicio, fecha_final), conn)
    cerrar_db(conn)

    df_embo_soc.rename(columns={"Dsc_Socio":"Socio", "Tienda":"tienda"}, inplace=True)
    
    
    return df_embo_soc


def merma_bid(fecha_inicio, fecha_final):

    fecha_inicio = "'" + fecha_inicio + "'"
    fecha_final = "'" + fecha_final + "'"

    conn = abrir_db("MP-VW-DB-017", "Tarjeta_Digital")
    df_bid = ejecutar_query(query_plasticos_bid(fecha_inicio, fecha_final), conn)
    cerrar_db(conn)
    
    return df_bid

def merma_soc(fecha_inicio, fecha_final):

    fecha_inicio = "'" + fecha_inicio + "'"
    fecha_final = "'" + fecha_final + "'"

    conn = abrir_db("MP-VW-DB-023", "ORIGINACION_MIGRACION")
    df_soc = ejecutar_query(query_plasticos_soc(fecha_inicio, fecha_final), conn)
    cerrar_db(conn)
    
    return df_soc
    
    


# In[6]:


#Aqui se desarrolla el control de inventario
#Se van a definir todas las tiendas con todo su stock origen, para que esto funcione el codigo necesita estar siempre corriendo (usar una compu como servidor)
def definir_tiendas():
    tiendas = []

    no_existe_archivo = True

    #Se verifica que exista el archivo y que no tenga datos en blanco
    while no_existe_archivo:
        try:
            df_tiendas = pd.read_excel("Respaldo.xlsx", dtype={"Tienda":str, "Tienda_Vision":str})
            no_existe_archivo = False
            
        except FileNotFoundError:
            input("¡¡Archivo Respaldo.xlsx no encontrado!!\n\nAsegurese que se encuentra con el nombre y ubicación correcta\n\nEscriba enter en el rectangulo para continuar\n\n")

    datos_vacios = True
    texto_stock = True

    while datos_vacios or texto_stock:
        df_tiendas = pd.read_excel("Respaldo.xlsx", dtype={"Tienda":str, "Tienda_Vision":str})
        
        try:
            df_tiendas["Stock"] = df_tiendas["Stock"] + 0
            texto_stock = False
        except TypeError:
            input("El archivo Respaldo.xlsx contiene texto en la columna Stock\nDebe contener unicamente numeros enteros\n")
            texto_stock = True
            
        if df_tiendas[['Stock', 'Tienda']].isna().any().any():
            input("Hay Tiendas o Stocks vacios en el archivo Respaldo\n")
            datos_vacios = True
        else:
            datos_vacios = False

    #df_tiendas["Consumo diario"] = df_tiendas["Consumo diario"].apply(lambda x: [int(i) for i in x.split(",")])
 
    for index, row in df_tiendas.iterrows():
        t = Tienda(row["Tienda"])
        t.ajustar_stock(row["Stock"])
        #t.consumo_diario = row["Consumo diario"]
        #t.calcular_promedio()
        t.promedio = 0
        t.tienda_vision = row["Tienda_Vision"]
        t.sucursal = row["Sucursal"]
        
        tiendas.append(t)

    #Se transforma en diccionario el numero de tienda para que sea mas eficiente las operaciones, es {numero de tienda:tienda}
    tiendas = {tienda.get_numero_tienda(): tienda for tienda in tiendas}

    return tiendas

def modificar_stock(columna_tienda, columna_plasticos, tiendas, metodo, df):

    #Se implementa una busqueda de tienda para a esa agregarle, establecerle o quitarle el stock, si en el diccionario existe entonces a esa tienda se le aplica la operacion
    for _, row in df.iterrows():    
        tienda = tiendas.get(row[columna_tienda])
        if tienda:
            getattr(tienda, metodo)(row[columna_plasticos])
            #tienda.calcular_promedio()
    
    return tiendas   

def eliminacion_diaria(tiendas):

    exportar_archivo(tiendas, "Respaldo_Anterior.xlsx")
    
    hoy = date.today()
    df_bid = bid()
    df_soc = soc()
    df_eliminar = pd.concat([df_soc,df_bid])
    tiendas = modificar_stock("NumeroTienda", "CantidadRegistros", tiendas, "quitar_stock", df_eliminar)
    #Esto es para pruebas
    nombre = "Stock_Eliminado " + str(hoy) + ".xlsx"
    df_eliminar.to_excel(nombre, index=False)
    
    if df_bid.empty:
        print("\tBid esta vacio")
    if df_soc.empty:
        print("\tSoc esta vacio")
        
    with open("Fecha ultima eliminacion.txt", "w") as archivo:
        archivo.write(str(hoy))

    exportar_archivo(tiendas, "Respaldo.xlsx")

    destinatario = "luis.campos@bradescard.com.mx"
    asunto = "Eliminacion diaria Exitosa"
    mensaje = "La eliminacion diario del control de inventario ha sido exitosa"
    enviar_correo(asunto, mensaje, destinatario)
    
    



# In[7]:


#Generacion de reportes
#Respaldo
def exportar_archivo(dic_tiendas, nombre):

    data = [tienda.to_dict() for tienda in dic_tiendas.values()]
    df = pd.DataFrame(data)
    #df.drop(columns={"Promedio diario"}, inplace=True)

    #Con try catxh para cuando el usuario lo tiene abierto o no lo logra exportar no de error si no un mensaje al usuario
    try:
        df.to_excel(nombre, index=False)
        
    except PermissionError:
        messagebox.showerror(
            "Archivo en uso",
            f"No se pudo guardar el archivo '{nombre}' porque está abierto.\n\nCiérralo e inténtalo de nuevo."
        )
    
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Ocurrió un error al generar el archivo:\n{str(e)}"
        )

#Reporte Stocks actuales
def reporte_stock(dic_tiendas, nombre, socio, tienda):

    data = [tienda.to_dict() for tienda in dic_tiendas.values()]
    df = pd.DataFrame(data)

    #En caso de que no se seleccione socio o tienda se pasa el query tal cual
    if socio != "Todos los socios" and socio != "":
        df = df[df["Socio"] == socio]
    if tienda != "":
        df = df[df["Tienda"] == tienda] 
    
    #df.drop(columns={"Promedio diario"}, inplace=True)

    #Con try catxh para cuando el usuario lo tiene abierto o no lo logra exportar no de error si no un mensaje al usuario
    try:
        df.to_excel(nombre, index=False)
        messagebox.showinfo("CID", "Reporte generado con exito")
        
    except PermissionError:
        messagebox.showerror(
            "Archivo en uso",
            f"No se pudo guardar el archivo '{nombre}' porque está abierto.\n\nCiérralo e inténtalo de nuevo."
        )
    
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Ocurrió un error al generar el archivo:\n{str(e)}"
        )

#Reporte de embozado
def reporte_embozado(fecha_inicio, fecha_final, socio, tienda):
    df_embozado_soc = embozado_soc(fecha_inicio, fecha_final)
    df_embozado_bid = embozado_bid(fecha_inicio, fecha_final)
    df_embozado = pd.concat([df_embozado_soc, df_embozado_bid])

    if socio != "Todos los socios" and socio != "":
        df_embozado = df_embozado[df_embozado["Socio"] == socio]
    if tienda != "":
        tienda = int(tienda)
        df_embozado = df_embozado[df_embozado["tienda"] == tienda] 

    try:
        df_embozado.to_excel("Reporte de Embozado.xlsx", index=False)
        messagebox.showinfo("CID", "Reporte generado con exito")
    
    except PermissionError:
        messagebox.showerror(
            "Archivo en uso",
            f"No se pudo guardar el archivo Reporte de Embozado porque está abierto.\n\nCiérralo e inténtalo de nuevo."
        )
    
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Ocurrió un error al generar el archivo:\n{str(e)}"
        )

#Reporte merma
def reporte_merma(fecha_inicio, fecha_final, socio, tienda):

    #Se consulta elquery de bid con fecha de inicio y final
    df_bid = merma_bid(fecha_inicio, fecha_final)
    df_soc = merma_soc(fecha_inicio, fecha_final)

    df_merma = pd.concat([df_bid, df_soc])

    #Se quita los de no tomo el plastico porque no cuentan para calcular el porcentaje
    df_merma = df_merma[df_merma["StatusEmbozado"] != "No tomó el plástico"]

    
    #En caso de que no se seleccione socio o tienda se pasa el query tal cual
    if socio != "Todos los socios" and socio != "":
        df_merma = df_merma[df_merma["Socio"] == socio]
    
    if tienda != "":
        df_merma = df_merma[df_merma["NumeroTienda"] == tienda]   

    #Se cambian los valores null por "NO_ESPEC" de la tabla para que si se tome en cuenta 
    df_merma["StatusEmbozado"].fillna("NO_ESPEC", inplace=True)

    #Se crea una tabla pivote la cual hace que los valores de status se hagan columnas y en filas las tiendas y que cantidad tienen de cada valor (de cada status cuantos tiene cada tienda)
    tabla = df_merma.pivot_table(
        index=['NumeroTienda','Socio'],
        columns='StatusEmbozado',
        aggfunc='size',
        fill_value=0
    )

    #En caso de que una tienda no tenga alguno de estos estatus para que no truene el codigo se agrega y se pone 0 en cantidad de casos de esa tienda
    columnas_esperadas = ["Embozado con éxito", "Mal embozado", "NO_ESPEC"]
    for col in columnas_esperadas:
        if col not in tabla.columns:
            tabla[col] = 0

    #Se calcula el porcentaje con las nuevas columnas creadas
    tabla["Porcentaje"] = round(((tabla["Mal embozado"] / (tabla["Embozado con éxito"] + tabla["Mal embozado"] + tabla["NO_ESPEC"]))*100),2)
    tabla["Porcentaje"] = tabla["Porcentaje"].astype(str) + "%"

    #Se exporta el archivo 
    try:
        tabla.to_excel("Reporte_merma.xlsx", sheet_name="Reporte merma")
        messagebox.showinfo("CID", "Reporte generado con exito")
        
    except PermissionError:
        messagebox.showerror(
            "Archivo en uso",
            f"No se pudo guardar el archivo Reporte_merma porque está abierto.\n\nCiérralo e inténtalo de nuevo."
        )
    
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Ocurrió un error al generar el archivo:\n{str(e)}"
        )

def reporte_consumo(fecha_inicio, fecha_final, tiendas, socio, tienda_filtrada):

    df_bid = bid_consumo(fecha_inicio, fecha_final)
    df_soc = soc_consumo(fecha_inicio, fecha_final)
    
    df_consumo = pd.concat([df_bid, df_soc], ignore_index=True).groupby('NumeroTienda', as_index=False)['CantidadRegistros'].sum()

    f_inicio = date.fromisoformat(fecha_inicio)
    f_final = date.fromisoformat(fecha_final)
    dif_dias = (f_final - f_inicio).days
    
    df_consumo["Promedio"] = df_consumo["CantidadRegistros"] / dif_dias
    
    
    for _, row in df_consumo.iterrows():    
        tienda = tiendas.get(row["NumeroTienda"])
        if tienda:
            tienda.asignar_promedio(row["Promedio"])

    data = [tienda.to_dict() for tienda in tiendas.values()]
    df = pd.DataFrame(data)


    #En caso de que no se seleccione socio o tienda se pasa el query tal cual
    if socio != "Todos los socios" and socio != "":
        df = df[df["Socio"] == socio]
    
    if tienda_filtrada != "":
        df = df[df["Tienda"] == tienda_filtrada] 

    
    #df.drop(columns={"Consumo diario"}, inplace=True)
    df["Dias para acabar stock"] = round(df["Stock"] / df["Promedio diario"])

    try:
        df.to_excel("Reporte Consumo.xlsx", index=False)
        messagebox.showinfo("CID", "Reporte generado con exito")
        
    except PermissionError:
        messagebox.showerror(
            "Archivo en uso",
            f"No se pudo guardar el archivo Reporte Consumo porque está abierto.\n\nCiérralo e inténtalo de nuevo."
        )
    
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Ocurrió un error al generar el archivo:\n{str(e)}"
        )

    return tiendas


# In[8]:


#Funciones de la interfaz grafica
def boton_agregar(tiendas):

        exportar_archivo(tiendas, "Respaldo_Anterior.xlsx")
    
        try:
            df_archivo = pd.read_excel("Agregar_Stock.xlsx", dtype={"Numero de tienda":str})
            if df_archivo[['Numero de tienda', 'Plasticos distribuidos']].isna().any().any():
                messagebox.showinfo("CID", "Existen casos con datos vacios en el archivo Agregar. \nFavor de corregirlos")
                
            else:
                try:
                    tiendas = modificar_stock("Numero de tienda", "Plasticos distribuidos", tiendas, "agregar_stock", df_archivo)
                    messagebox.showinfo("CID", "Stock modificado con éxito")
                    exportar_archivo(tiendas, "Respaldo.xlsx")
                
                except TypeError:
                    messagebox.showerror("CID", "El archivo Agregar.xlsx contiene texto en la columna Plasticos distribuidos\nDebe contener unicamente numeros enteros")
                    
                    
                
        except FileNotFoundError:
            messagebox.showerror(
                "Archivo no encontrado",
                f"No se encontro el archivo 'Agregar_Stock.xlsx'\n\nAsegure que exista con el nombre y ubicación correcta"
            )


def boton_ajustar(tiendas):

    exportar_archivo(tiendas, "Respaldo_Anterior.xlsx")
    
    try:
        df_archivo = pd.read_excel("Ajustar_Stock.xlsx", dtype = {"Numero de tienda":str})
        if df_archivo[['Numero de tienda', 'Plasticos distribuidos']].isna().any().any():
            messagebox.showerror("CID", "Existen casos con datos vacios en el archivo Ajustar. \nFavor de corregirlos")

        else:
            try:
                df_archivo["Plasticos distribuidos"] = df_archivo["Plasticos distribuidos"] + 0 
                tiendas = modificar_stock("Numero de tienda", "Plasticos distribuidos", tiendas, "ajustar_stock", df_archivo)
                messagebox.showinfo("CID", "Stock modificado con éxito")
                exportar_archivo(tiendas, "Respaldo.xlsx")

            except TypeError:
                    messagebox.showerror("CID", "El archivo Ajustar.xlsx contiene texto en la columna Plasticos distribuidos\nDebe contener unicamente numeros enteros")
            
            
    except FileNotFoundError:
        messagebox.showerror(
            "Archivo no encontrado",
            f"No se encontro el archivo 'Ajustar_Stock.xlsx'\n\nAsegure que exista con el nombre y ubicación correcta"
        )

def boton_eliminar(fecha_inicio, fecha_final, tiendas):
    
    #Esta opciojn es solo en caso de que la eliminacion automatica falle, se hara la eliminacion de un rango de fechas
    #Ejemplo de quie poner en el rango de fechas: En caso de que hoy fuera 25 y la ultima fecha registrada sea el 23 entonces poner en el rango de fechas del 23 al 25
    respuesta = simpledialog.askstring("Advertencia", "Estas seguro que quieres hacer la eliminación diaria manualmente (Esta opción es solo para cuando la eliminacion automatica falle)\nEscribe S para continuar o enter para regresar:")
    
    if respuesta == "S" or respuesta == "s":
        exportar_archivo(tiendas, "Respaldo_Anterior.xlsx")
        fecha_inicio = "'" + str(fecha_inicio) + "'"
        fecha_final = "'" + str(fecha_final) + "'"
        
        conn = abrir_db("MP-VW-DB-023", "ORIGINACION_MIGRACION")
        df_soc = ejecutar_query(query_soc(fecha_inicio, fecha_final), conn)
        cerrar_db(conn)
    
        conn = abrir_db("MP-VW-DB-017", "Tarjeta_Digital")
        df_bid = ejecutar_query(query_bid(fecha_inicio, fecha_final), conn)
        cerrar_db(conn)
    
        df_eliminar = pd.concat([df_soc,df_bid])
        tiendas = modificar_stock("NumeroTienda", "CantidadRegistros", tiendas, "quitar_stock", df_eliminar)
        
        df_eliminar.to_excel("Stock eliminado.xlsx", index=False)
        messagebox.showinfo("CID", "Stock modificado con éxito")

        exportar_archivo(tiendas, "Respaldo.xlsx")

def generar_archivo(socio, tienda, opc_reportes, tiendas, fecha_inicio, fecha_final):
    
    fecha_inicio = str(fecha_inicio.strftime("%Y-%m-%d"))
    fecha_final = str((fecha_final + timedelta(days=1)).strftime("%Y-%m-%d"))

    if opc_reportes == "":
        messagebox.showwarning("CID", "Selecciona una opción primero")
        return
    match opc_reportes:
        case "Reporte de Stock Actual":    
            reporte_stock(tiendas, "Reporte de Stock Actual.xlsx", socio, tienda)
        case "Reporte de Embozado":
            reporte_embozado(fecha_inicio, fecha_final, socio, tienda)
        case "Reporte de Merma":
            reporte_merma(fecha_inicio, fecha_final, socio, tienda)
        case "Reporte de Consumo":
            reporte_consumo(fecha_inicio, fecha_final, tiendas, socio, tienda)
        case "Eliminacion de Stocks (SOLO EN CASO DE QUE LA ELIMINACION AUTOMATICA FALLE)":
            boton_eliminar(fecha_inicio, fecha_final, tiendas)
            



# In[9]:


def enviar_correo(asunto, mensaje, destinatario):
    # Abrir aplicación de Outlook
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
 
    # Configuración del correo
    mail.To = destinatario
    #mail.CC = ""
    #Asunto
    mail.Subject = asunto
 
    # Ruta de la imagen de la firma (logo, etc.)
    #imagen_path = r"C:\Users\alan.arias\Documents\Proceso DBTA\imagen_firma.jpg"  # cámbialo a tu ruta
 
    # Adjuntar la imagen al correo
    #attachment = mail.Attachments.Add(imagen_path)
    #attachment.PropertyAccessor.SetProperty(
    #    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "FirmaLogo"
    #)
 
    # Firma en HTML con la imagen
    firma = """
            <div style="font-family:Arial, sans-serif; font-size:10pt; color:#000000;">
            <b>Luis Francisco Campos Ramirez</b><br>
            <b>Operaciones (Aclaraciones)</b><br>
                    Tel. +52 (33) 38842300<br>
                    Ext.2338<br>
                    luis.campos@bradescard.com.mx<br>
                    BRADESCARD MÉXICO, S. DE R.L<br>
                    Camino al Iteso #8310. Parque Industrial del Bosque 1, CP. 45609<br>
                    Tlaquepaque, Jalisco, México<br><br>
            </div>
                """
 
    mensaje = f"""
            <div style="font-family:Arial, sans-serif; font-size:11pt; color:#000000;">
            <p>{mensaje}</p>
            <br><br><br>
                    {firma}
            </div>
                """
 
    #Mensaje    
    mail.HTMLBody = mensaje
 
    #Archivo
    #archivo = r"C:\ruta\archivo.xlsx"
    #mail.Attachments.Add(archivo)
 
    # Enviar el correo
    mail.Send()
    


# In[10]:


#Interfaz Grafica:
def interfaz_grafica(tiendas):
    #  Ventana principal
    ventana = tk.Tk()
    ventana.title("CID")
    ventana.geometry("880x560")
    ventana.minsize(840, 520)
    ventana.configure(bg="#F3F4F8")  # fondo más limpio

    # Tema y estilos ttk
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except:
        pass

    # Paleta
    COLOR_BG = "#F3F4F8"     # fondo del cuadro
    COLOR_CARD = "#FFFFFF"   # tarjetas
    COLOR_TEXT = "#1F2330"
    COLOR_MUTED = "#6E7480"
    COLOR_PRIMARY = "#6C4EE3"     # Para el boton de "Generar"
    COLOR_PRIMARY_HOVER = "#5B41C5"
    COLOR_ADD = "#46265C"         # Boton de "Agregar"
    COLOR_ADD_HOVER = "#3D2051"
    COLOR_DANGER = "#CA0F14"      # Boton de "Ajustar"
    COLOR_DANGER_HOVER = "#A50C10"
    COLOR_BORDER = "#E5E7EF"

    # Tipografías
    ventana.option_add("*Font", "Arial 10")
    ventana.option_add("*TButton.Padding", 10)

    # Estilos base
    style.configure("TLabel", background=COLOR_CARD, foreground=COLOR_TEXT)
    style.configure("Muted.TLabel", background=COLOR_CARD, foreground=COLOR_MUTED)
    style.configure("Header.TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=("Arial", 18, "bold"))
    style.configure("SubHeader.TLabel", background=COLOR_BG, foreground=COLOR_MUTED, font=("Arial", 10))
    style.configure("Card.TFrame", background=COLOR_CARD, borderwidth=1, relief="solid")
    style.configure("TEntry", fieldbackground="#FFFFFF", padding=6)
    style.configure("TCombobox", fieldbackground="#FFFFFF", padding=6)

    # Botones
    style.configure(
        "Agregar.TButton",
        background=COLOR_ADD, foreground="#FFFFFF",
        relief="flat", borderwidth=0, font=("Arial", 10, "bold")
    )
    style.map("Agregar.TButton", background=[("active", COLOR_ADD_HOVER)])
    style.configure(
        "Ajustar.TButton",
        background=COLOR_DANGER, foreground="#FFFFFF",
        relief="flat", borderwidth=0, font=("Arial", 10, "bold")
    )
    style.map("Ajustar.TButton", background=[("active", COLOR_DANGER_HOVER)])
    style.configure(
        "Generar.TButton",
        background=COLOR_PRIMARY, foreground="#FFFFFF",
        relief="flat", borderwidth=0, font=("Arial", 10, "bold")
    )
    style.map("Generar.TButton", background=[("active", COLOR_PRIMARY_HOVER)])

    # Layout raíz 
    ventana.grid_rowconfigure(1, weight=1)
    ventana.grid_columnconfigure(0, weight=1)

    # Header
    header = tk.Frame(ventana, bg=COLOR_BG)
    header.grid(row=0, column=0, sticky="ew", padx=28, pady=(18, 10))
    header.grid_columnconfigure(0, weight=1)
    ttk.Label(header, text="CID", style="Header.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Label(header, text="Gestión de stock, ajustes y reportes.", style="SubHeader.TLabel").grid(row=1, column=0, sticky="w")

    # Contenedor principal
    container = tk.Frame(ventana, bg=COLOR_BG)
    container.grid(row=1, column=0, sticky="nsew", padx=28, pady=10)
    container.grid_columnconfigure(0, weight=1)
    container.grid_columnconfigure(1, weight=1)
    container.grid_rowconfigure(0, weight=1)

    # Columna izquierda: acciones
    left = ttk.Frame(container, style="Card.TFrame")
    left.grid(row=0, column=0, sticky="nsew", padx=(0, 12), pady=0, ipadx=6, ipady=6)
    left.grid_columnconfigure(0, weight=1)

    # Simular borde sutil
    left.configure(borderwidth=1)
    left["borderwidth"] = 1
    left["padding"] = (16, 16)

    ttk.Label(left, text="Acciones", font=("Arial", 11, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 8))

    boton1 = ttk.Button(
        left, text="Agregar Stock",
        command=lambda: boton_agregar(tiendas),
        style="Agregar.TButton"
    )
    boton1.grid(row=1, column=0, sticky="ew", pady=(4, 10), ipady=10)

    boton2 = ttk.Button(
        left, text="Ajustar Stock",
        command=lambda: boton_ajustar(tiendas),
        style="Ajustar.TButton"
    )
    boton2.grid(row=2, column=0, sticky="ew", pady=(0, 4), ipady=10)

    # Columna derecha: tarjeta de filtros
    right = ttk.Frame(container, style="Card.TFrame")
    right.grid(row=0, column=1, sticky="nsew", padx=(12, 0), pady=0)
    right.grid_columnconfigure(0, weight=1)
    right.grid_columnconfigure(1, weight=1)

    # Tarjeta con padding interno
    right.configure(borderwidth=1)
    right["padding"] = (20, 18)

    # Título de la tarjeta (en lugar de LabelFrame gris)
    title_bar = tk.Frame(right, bg=COLOR_CARD)
    title_bar.grid(row=0, column=0, columnspan=2, sticky="ew")
    ttk.Label(title_bar, text="Filtros y reportes", font=("Arial", 11, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 10))

    # Fila 1: Socio / Tienda
    ttk.Label(right, text="Socio:").grid(row=1, column=0, sticky="w", pady=(0, 4))
    ttk.Label(right, text="Tienda:").grid(row=1, column=1, sticky="w", pady=(0, 4))

    socios = ['Todos los socios', 'CyA', 'BRADESCARD', 'LOB', 'SUBURBIA', 'AURRERA', 'SHASA', 'GCC', 'PROMODA', 'CCP']
    combo_socios = ttk.Combobox(right, values=socios, state="readonly")
    combo_socios.grid(row=2, column=0, sticky="ew", padx=(0, 8))
    entrada_tiendas = ttk.Entry(right)
    entrada_tiendas.grid(row=2, column=1, sticky="ew", padx=(8, 0))

    # Espacio
    tk.Frame(right, height=6, bg=COLOR_CARD).grid(row=3, column=0, columnspan=2)

    # Fila 2: Reportes
    ttk.Label(right, text="Reportes:").grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 4))
    opciones = ["Reporte de Stock Actual", "Reporte de Embozado", "Reporte de Merma", "Reporte de Consumo", "Eliminacion de Stocks (SOLO EN CASO DE QUE LA ELIMINACION AUTOMATICA FALLE)"]
    combo_reporte = ttk.Combobox(right, values=opciones, state="readonly")
    combo_reporte.grid(row=5, column=0, columnspan=2, sticky="ew")

    # Espacio
    tk.Frame(right, height=6, bg=COLOR_CARD).grid(row=6, column=0, columnspan=2)

    # Fila 3: Fechas
    ttk.Label(right, text="Fecha inicio:").grid(row=7, column=0, sticky="w", pady=(0, 4))
    ttk.Label(right, text="Fecha fin:").grid(row=7, column=1, sticky="w", pady=(0, 4))

    date_inicio = DateEntry(
        right, width=12, background=COLOR_PRIMARY,
        foreground="white", borderwidth=0, date_pattern="dd/mm/yyyy"
    )
    date_inicio.grid(row=8, column=0, sticky="w")
    date_fin = DateEntry(
        right, width=12, background=COLOR_PRIMARY,
        foreground="white", borderwidth=0, date_pattern="dd/mm/yyyy"
    )
    date_fin.grid(row=8, column=1, sticky="w")

    # Botón Generar a la derecha
    boton_generar = ttk.Button(
        right, text="Generar",
        style="Generar.TButton",
        command=lambda: generar_archivo(
            combo_socios.get(),
            entrada_tiendas.get(),
            combo_reporte.get(),
            tiendas,
            date_inicio.get_date(),
            date_fin.get_date()
        )
    )
    boton_generar.grid(row=9, column=1, sticky="e", pady=(14, 0))

   
    schedule.every().day.at("15:38").do(eliminacion_diaria, tiendas)

    def revisar_schedule():
        schedule.run_pending()
        ventana.after(1000, revisar_schedule)

    revisar_schedule()


    ventana.mainloop()


# In[11]:


def main():

    #Carga la informacion de origen en un archivo donde se definen las tiendas y el stock,
    #esto es para la primera vez y el programa no deberia interrumpirse,
    #en caso de que se frene cargar de origen el respaldo qu debe generar el codigo en cada accion
    warnings.filterwarnings('ignore')
    tiendas = definir_tiendas()
    
    interfaz_grafica(tiendas)

    #Enviar correo si se cierra el programa (No contempla si se deja de ejecutar el codigo solo si se cierra)
    destinatario = "luis.campos@bradescard.com.mx"
    asunto = "Control de Inventario Cerrado"
    mensaje = "Se cerro el programa de Control de Inventario, favor de verificar"
    #enviar_correo(asunto, mensaje, destinatario)


main()

