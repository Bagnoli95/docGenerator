from docxtpl import DocxTemplate
from tqdm import tqdm
import pandas as pd
import os


"""
    docGenerator.
    Input: Documento excel. Tiene que estar en la misma carpeta que el .py, para simplificar las cosas.
    Primero, lee el excel para obtener las configuraciones que necesite de la pag Config.
    Segundo, procesar según los datos configurados en la segunda hoja para cargar en la plantilla de word.
    Output: Debe crear carpetas y dentro los documentos generados (docs, pdf o ambos).
    comando para ejecutar el instalable: py -3.10 setup.py py2exe

    comando para armar el .exe desde windows ==>> py -3.10 -m PyInstaller --onefile .\docGenerator.py
    debe estar ubicado en la carpeta donde esté el .py a compilar.

"""


# Funcion para generar los documentos
def procesar_documentos(contexto_doc):
    # Cargar la planilla
    plantilla = DocxTemplate(plantilla_nombre)
    # Cargar el contexto en la planilla
    plantilla.render(contexto_doc)
    ## Guardar la planilla
    # Darle un nombre
    nombre_archivo = contexto_doc['nombre_archivo']
    nombre_archivo = nombre_archivo + '.' + tipo_output
    # Guardar la planilla
    plantilla.save(carpeta + '\\' + nombre_archivo)


# Funcion para generar los pdf
def generar_pdf():
    print('(Aun no esta disponible la generación de pdf)')
    return

# Mensaje de bienvenida
print('_'*80)
print('\n\t\t\tBIENVENIDO AL GENERADOR DE ARCHIVOS')
print('_'*80)


# Cargar el archivo de Config
doc_excel = pd.read_excel('docGenerator_BBDD.xlsx', sheet_name=['Config', 'Datos'])
doc_config = doc_excel.get('Config')
doc_datos = doc_excel.get('Datos')
print('\n\tLOS DATOS DE CONFIGURACION SON:')
print('Config:\n', doc_config)
print('\n\tLOS DATOS A PROCESAR SON:')
print('Datos:\n', doc_datos)


# Recuperar valores de configuracion
plantilla_nombre = doc_config.iloc[0]['Valor Config']
print('\nPlantilla nombre: ', plantilla_nombre)
tipo_output = doc_config.iloc[1]['Valor Config']
print('Tipo resultado: ', tipo_output)
pos_procesado = doc_config.iloc[2]['Valor Config']
print('Pos procesado: ', pos_procesado)
nombre_generacion = doc_config.iloc[3]['Valor Config']
print('Nombre para los archivos: Toma de la primera columna de la página de datos\n',)
print('_'*80)

# Ejecutar o no el proceso
print('\n')
usuario_flag = input('\nDesea procesar los datos expuestos? (S/N): ')
print('\n')
print('_'*80)
print('\n')

if usuario_flag == 'S' or usuario_flag == 's':
    # Crear directorio si no existe
    carpeta = os.getcwd() # Obtiene la ruta de trabajo actual
    carpeta = carpeta + '\\Generados'
    if os.path.isdir(carpeta):
        print("La carpeta ya existe, se crearan los archivos dentro.")
    else:
        print("La carpeta se creó correctamente.")
        os.makedirs(carpeta)

    pbar_contador = len(doc_datos)
    pbar = tqdm(total=pbar_contador, desc="PROCESANDO")
    
    nombres_columnas = doc_datos.columns.values
    contexto = {}
    contador = 0
    # Recorrer las filas del dataFrame
    for index, row in doc_datos.iterrows():
        # Recorrer las columnas de una fila para agregar al contexto
        for i in range(len(nombres_columnas)):
            nombre = nombres_columnas[i]
            valor = row[nombre]
            contexto[nombre] = valor
        
        # Procesar las Planillas
        procesar_documentos(contexto)
        contador += 1
        pbar.update(1)

    pbar.close()
    if pos_procesado == 'pdf':
        generar_pdf()
    print(f'Proceso terminado, se crearon {contador} archivos\n')
else:
    print('Proceso cancelado.\n')

os.system('Pause')