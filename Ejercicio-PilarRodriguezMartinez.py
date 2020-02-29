# -*- coding: utf-8 -*-

import openpyxl
import pandas as pd
import random
from pandas import ExcelWriter
from pandas import ExcelFile
from bs4 import BeautifulSoup


# ---------- FUNCIONES AUXILIARES PARA EL PROGRAMA PRINCIPAL ----------

def func_Ganador (max_puntuacion):
	"""
	Función para calcular aleatoriamente el ganador del partido dada una puntuación máxima
	"""
	punt_Equipo_i1 = random.randint(0,int(max_puntuacion))
	punt_Equipo_i2 = random.randint(0,int(max_puntuacion))

	if punt_Equipo_i1 > punt_Equipo_i2:
		Ganador  = "Equipo_i1"
	elif punt_Equipo_i1 < punt_Equipo_i2:
		Ganador = "Equipo_i2"
	else:
		Ganador = "Empate"

	resultados = [Ganador, punt_Equipo_i1, punt_Equipo_i2]

	return resultados

def func_click(index,table):
	"""
	Función click para determinar si se hace click sobre una celda
	"""
	if (type(index) != "undefined"):
		table.rows()[index].classList.toggle("selected")
	return()

def fila_Seleccionada (cuerpo_HTML, tam_Tabla):
	"""
	Función para determinar cual es la fila seleccionada y cambiar el color de la misma
	"""
	soup = BeautifulSoup(cuerpo_HTML, "html.parser")
	Lista_index_table = soup.find(id="table")

	index = Lista_index_table[0]
	table = Lista_index_table[1]

	table.onclick = func_click(index,table)
	
	return()

# --------------------------------------------------------------------------

# ----------------- PROGRAMA PRINCIPAL -----------------

# 1. Lectura del nombre del archivo Excel del que se desean leer los datos
#    Como se sabe que los datos se encuentran en la hoja TestPython se leen los datos de esa hoja del Excel 

file_name = input('Por favor, escribir a continuación el nombre del archivo Excel (incluyendo el .xlsx) y dar a intro: ')
resp = ''

while resp != 's':
	print ('El nombre del archivo Excel del que deseas leer los datos es: ', file_name)
	resp = input('¿El nombre del archivo escrito es el correcto? (Contestar con s/n, en minusculas): ')
	if resp != 's':
		file_name = input('Por favor, escribir de nuevo el nombre del archivo: ')

df = pd.read_excel(file_name, sheet_name='TestPython')

# 2. Se pide introducir la puntuación máxima que se podrá alcanzar en los partidos para poder generar un número aleatorio
#    como resultado de los partidos que sea coherente (y no valga 100, o valores que no se podrian alcanzar como goles en un 
#    partido real)

max_puntuacion = input("Fija cual es la puntación máxima que se va a poder alcanzar en los partidos (debe ser un numero natural): ")
while int(max_puntuacion) < 0:
	max_puntuacion = input("La puntuación debe ser un numero positivo")

# 3. Se guardan los valores de las columnas del excel en una lista (lista_columnas)
#    Los valores de dicha lista se convierten a tipo String (y se guardan en el diccionario conversor) para poder ser utilizados 
#    en la llamada a pd.read_excel

lista_columnas = []
columnas = df.columns

for c in columnas:
	lista_columnas.append(c)

conversor = {col: str for col in lista_columnas}

df_actual = pd.read_excel('ExcelFile.xlsx', conversor = conversor)

# 4. La variable f_HTML se utiliza para abrir el Fichero Salida (html) en modo escritura ('w'), en el cual se escribirán los datos
#    que se desean mostrar en el fichero html
#    En la variable cuerpo_HTML se diseña la estructura que contendrá el fichero html

f_HTML = open('FicheroSalida.html', 'w')

cuerpo_HTML = """
<html>
	<head>
		<title>Ejercicio - Pilar Rodriguez Martinez </title>
		<p>O-o-O-O-o-O-O-o-O-[ Tabla de Enfrentamientos ]-O-o-O-O-o-O-O-o-O</p>
	</head>
	<body>
        <table id="table" border="1">
        	<tr>
            	<th>&nbspEquipos</th>
                <th>Puntuaciones</th>
            </tr>
        </table>
    </body>
</html>
"""

# Estos print son meramente informativos, para que sirvan como guía a quien lo ejecute de por donde va la ejecución del programa
# -----------------------------------------------------------------
print("-------------------->")
print("Lectura de Datos....")

print("--- INICIO DEL BUCLE --->")

# En la variable tam_Tabla guardamos el tamaño de la tabla de datos (necesario para crear el bucle que genera los ganadores de cada
# partido)
tam_Tabla = len(df_actual['Columna1'])  
print (tam_Tabla)
# -----------------------------------------------------------------

# --- CREACIÓN DEL CUERPO DEL HTML ---
# En cada iteración del siguiente bucle for se determina quien es el ganador de cada partido de la siguiente manera:
# 1. Se realiza una llamada a la funcion Ganador para determinar el ganador del partido i-esimo y cuales son las puntuaciones obtenidas
# 2. Se guardan las puntuaciones de los dos equipos en las variables punt_Equipo_i1 y punt_Equipo_i2
# 3. En función de si el ganador es el primer equipo o el segundo se generara un cuerpo del HTML distinto dado que se debe marcar en 
#    negrita al ganador correspondiente del partido. * Se ha tenido en cuenta una tercera casuística que es la de que no gane ni el 
#    equipo i1 ni el i2 (es decir, que se produzca un empate), en tal caso se ha decidido que ninguno de los dos equipos aparezca
#    en negrita en el HTML

for i in range(tam_Tabla):
	#-*- LLAMADA A LA FUNCION GANADOR 
	resultados = func_Ganador(max_puntuacion)

	Ganador = resultados[0]
	punt_Equipo_i1 = resultados[1]
	punt_Equipo_i2 = resultados[2]

	print("i:",i)
	print("Ganador: ",Ganador)

	if Ganador == "Equipo_i1":
		cuerpo_HTML += """<body>
						<table id="table" border="1">
							<tr>
            					<td><b>%s</b></td>
                				<td><b>%s %s %s</b></td>
            				</tr>
            				<tr>
            					<td>%s</td>
                				<td>%s %s %s</td>
            				</tr>
        				</table>
    				</body>"""%(df_actual['Columna1'][i],9*"&nbsp",str(punt_Equipo_i1),9*"&nbsp",df_actual['Columna2'][i],9*"&nbsp",str(punt_Equipo_i2),9*"&nbsp")

	elif Ganador == "Equipo_i2":
		cuerpo_HTML += """<body>
						<table id="table" border="1">
							<tr>
            					<td>%s</td>
                				<td>%s %s %s</td>
            				</tr>
            				<tr>
            					<td><b>%s</b></td>
                				<td><b>%s %s %s</b></td>
            				</tr>
            			</table>
            		</body>"""%(df_actual['Columna1'][i],9*"&nbsp",str(punt_Equipo_i1),9*"&nbsp",df_actual['Columna2'][i],9*"&nbsp",str(punt_Equipo_i2),9*"&nbsp")

	else:
		cuerpo_HTML += """<body>
						<table id="table" border="1">
							<tr>
            					<td>%s</td>
                				<td>%s %s %s</td>
            				</tr>
            				<tr>
            					<td>%s</td>
                				<td>%s %s %s</td>
            				</tr>
            			</table>
            		</body>"""%(df_actual['Columna1'][i],9*"&nbsp",str(punt_Equipo_i1),9*"&nbsp",df_actual['Columna2'][i],9*"&nbsp",str(punt_Equipo_i2),9*"&nbsp")


	cuerpo_HTML += "<br><br>"	


# - Una vez finalizado el bucle y creados todos los resultados de los partidos se escriben los datos correspondientes en el HTML	
# Y se debería realizar la siguiente llamada a la funcion (que en este caso esta comentada)

# fila_Seleccionada(cuerpo_HTML, tam_Tabla)

# Las funciones func_click y fila_Seleccionada deberian ser utilizadas para cambiar el color de la celda al hacer click
# Esta parte tendria que investigarla un poco mas ya que aunque recuerdo que se hace utilizando las herramientas que pongo en esas funciones
# necesitaria investigar un poco mas sobre como se hace y no queria retrasarme mas en la entrega del programa 				

f_HTML.write(cuerpo_HTML)
f_HTML.close()