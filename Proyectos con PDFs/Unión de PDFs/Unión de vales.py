import PyPDF2
#import os
import os
import pandas as pd


# Función para combinar pdf
def combinar_pdf(pdf1, pdf2, pdf_resultante):
    # Abre los archivos PDF
    with open(pdf1, 'rb') as archivo1, open(pdf2, 'rb') as archivo2:
        # Crea objetos PDFReader para ambos archivos
        lector1 = PyPDF2.PdfReader(archivo1)
        lector2 = PyPDF2.PdfReader(archivo2)
       
        # Crea un objeto PDFWriter
        escritor = PyPDF2.PdfWriter()
       
        # Agrega las páginas del primer PDF al escritor
        for pagina in range(len(lector1.pages)):
            escritor.add_page(lector1.pages[pagina])
       
        # Agrega las páginas del segundo PDF al escritor
        for pagina in range(len(lector2.pages)):
            escritor.add_page(lector2.pages[pagina])
       
        # Guarda el resultado en un nuevo archivo PDF
        with open(pdf_resultante, 'wb') as resultado:
            escritor.write(resultado)


# Cambio de directorio a sitio de trabajo
directorio = '/Users/leonardorebollarruelas/Documents/GitHub/LeonardoRR95/Proyectos con PDFs/Unión de PDFs/Vales'

# Definir el directorio de trabajo "Original"
os.chdir(directorio)

# Inicializar listas para nombres de carpetas y directorios
carpetas = []
directorios = []

# Obtener todas las entradas en el directorio principal
with os.scandir(directorio) as entradas:
    for entrada in entradas:
        if entrada.is_dir():
            carpetas.append(entrada.name)
            directorios.append(entrada.path)

direcciones = pd.DataFrame({'carpetas' : carpetas, 'directorios' : directorios})

direcciones['VALES'] = direcciones['directorios'] + '/VALES'
direcciones['DIGITALES'] = direcciones['directorios'] + '/DIGITALES'

directorio = {}


# Realiza la generación de un directorio
for i in range(len(direcciones)):
    #i = 0
   
    d_aseg = direcciones['directorios'][i]
   
    c_aseg = []
    c_dir = []
   
    with os.scandir(d_aseg) as entradas:
        for entrada in entradas:
            if entrada.is_dir():
                c_aseg.append(entrada.name)
                c_dir.append(entrada.path)
   
    if len(c_aseg) > 0:
       
        archivos_v = []
        directorios_v = []
       
        archivos_d = []
        directorios_d = []
       
        d_vales = direcciones['VALES'][i]
        d_digitales = direcciones['DIGITALES'][i]
       
        # Obtener todas las entradas en el directorio principal
        with os.scandir(d_vales) as entradas:
            for entrada in entradas:
                if entrada.is_file():
                    archivos_v.append(entrada.name)
                    directorios_v.append(entrada.path)
       
        vales = pd.DataFrame({'vale' : archivos_v, 'direcciones_v' : directorios_v})
       
        # Obtener todas las entradas en el directorio principal
        with os.scandir(d_digitales) as entradas:
            for entrada in entradas:
                if entrada.is_file():
                    archivos_d.append(entrada.name)
                    directorios_d.append(entrada.path)
       
        digitales = pd.DataFrame({'digital' : archivos_d, 'direcciones_d' : directorios_d})
       
        directorio[carpetas[i]] ={'vales' : vales, 'digitales': digitales}
       
    else:
        ""

claves = list(directorio.keys())

#Fecha
fecha = '/Unidos 07062024'

nuevo_dir = 'C:/Users/Certeza/Documents/Cosas de Leo/Union Vales 2024/VALES UNIDOS' + fecha

os.makedirs(nuevo_dir, exist_ok= True)

#Proceso para unir PDF con la función "combinar_pdf(pdf1, pdf2, pdf_resultante)"
for clave in claves:
    #clave = 'HDI'
    digitales = directorio[clave]['digitales']
    vales = directorio[clave]['vales']
   
    #Limpieza del nombre de la variable de ambos casos "digitales" y "vales"
    vales['clave'] = vales['vale'].str.replace(r'(GNP/.|QUA/.|AXA/.|HDI/.|/.pdf)', '', regex=True)
    vales['clave'] = vales['clave'].str.replace(' ', '', regex = True)
   
    digitales['clave'] = digitales['digital'].str.replace(r'(/.c/.pdf|/.C/.PDF|/.C/.PDF/.PDF)', '', regex=True)
    digitales['clave'] = digitales['clave'].str.replace(r'(/.pdf|/.PDF)', '', regex=True)
    digitales['clave'] = digitales['clave'].str.replace(' ', '', regex = True)
   
    tabla = pd.merge(vales, digitales, on = 'clave')
   
    d_trabajo = nuevo_dir + '/' + str(clave)
   
    os.makedirs(d_trabajo)
   
    os.chdir(d_trabajo)
   
    for i in range(len(tabla['clave'])):
        #i = 0
        v_nom = tabla['clave'].loc[i]
        nombre = v_nom + '.C.pdf'
        vale = tabla['direcciones_v'].loc[i]
        digi = tabla['direcciones_d'].loc[i]
        combinar_pdf(vale, digi, nombre)
