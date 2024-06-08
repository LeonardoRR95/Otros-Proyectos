import PyPDF2
import os
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader

# Marca de Agua JPG "FIJO"
m_agua = '/Users/leonardorebollarruelas/Documents/GitHub/LeonardoRR95/Proyectos con PDFs/Marca de Agua/agua.jpg'

# Directorio y archivos
directorio = '/Users/leonardorebollarruelas/Documents/GitHub/LeonardoRR95/Proyectos con PDFs/Marca de Agua/Polizas'

# Inicializar listas para nombres de carpetas y directorios
carpetas = []
directorios = []

# Obtener todas las entradas en el directorio principal
with os.scandir(directorio) as entradas:
    for entrada in entradas:
        if entrada.is_dir():
            carpetas.append(entrada.name)
            directorios.append(entrada.path)

# Crear el DataFrame
df = pd.DataFrame({'carpetas': carpetas, 'directorio': directorios})

##########################################################################
aseguradora = "HDI" #DEFINIR PARA PODER GENERAR LA MARCA DE AGUA CORRECTA#
##########################################################################

aseg_dir = df[df['carpetas'] == aseguradora].iloc[0,1]

os.chdir(aseg_dir)

# Inicializar listas para nombres de carpetas y directorios
archivos = []
directorios = []

# Obtener todas las entradas en el directorio principal
with os.scandir(aseg_dir) as entradas:
    for entrada in entradas:
        if entrada.is_file():
            archivos.append(entrada.name)
            directorios.append(entrada.path)

# Crear el DataFrame
df_arch = pd.DataFrame({'archivos': archivos, 'directorio': directorios}) #Contiene las direcciones de las polizas

if aseguradora == "GNP":
   
    # =============================================================================
    # GNP
    # =============================================================================
   
    for i in range(0, len(df_arch['archivos'])):
        #i = 0
        archivo = df_arch.iloc[i, 1]
        n_archivo = df_arch.iloc[i, 0]
        temp_pdf = aseg_dir + '/temp.pdf'
       
        # Tamaño y posición de la imagen para marca de Agua
        imagen_ancho = 140  # ancho de la imagen en puntos
        imagen_alto = 70   # alto de la imagen en puntos
        posicion_x = 440    # posición x de la imagen desde el borde izquierdo en puntos
        posicion_y = 413    # posición y de la imagen desde el borde inferior en puntos

        # Crear un nuevo PDF temporal con la imagen usando reportlab
        c = canvas.Canvas(temp_pdf, pagesize=letter)
        # Ajustar el tamaño de la imagen
        imagen = ImageReader(m_agua)
        c.drawImage(imagen, posicion_x, posicion_y, width=imagen_ancho, height=imagen_alto)
        c.showPage()
        c.save()

        # Leer los PDFs
        lector_original = PyPDF2.PdfReader(archivo)
        lector_imagen = PyPDF2.PdfReader(temp_pdf)

        # Crear un escritor para el nuevo PDF
        escritor = PyPDF2.PdfWriter()

        # Combinar la primera página con la imagen como marca de agua
        pagina_original = lector_original.pages[0]
        pagina_imagen = lector_imagen.pages[0]

        # Crear una nueva página que combine la original y la imagen
        pagina_original.merge_page(pagina_imagen)
        escritor.add_page(pagina_original)

        # Añadir el resto de las páginas del PDF original
        for i in range(1, len(lector_original.pages)):
            escritor.add_page(lector_original.pages[i])

        # Guardar el nuevo PDF
        with open(os.path.join(aseg_dir, n_archivo), 'wb') as salida:
            escritor.write(salida)

        # Limpiar el archivo temporal
        os.remove(temp_pdf)
       
elif aseguradora == "HDI":

    # =============================================================================
    # HDI
    # =============================================================================
   
    temp_pdf = aseg_dir + '/temp.pdf'
   
    for i in range(0, len(df_arch['archivos'])):
        #i = 0
        archivo = df_arch.iloc[i, 1]
        n_archivo = df_arch.iloc[i, 0]
        temp_pdf = aseg_dir + '/temp.pdf'
       
        # Tamaño y posición de la imagen para marca de Agua
        imagen_ancho = 300  # ancho de la imagen en puntos
        imagen_alto = 150   # alto de la imagen en puntos
        posicion_x = 200    # posición x de la imagen desde el borde izquierdo en puntos
        posicion_y = 250    # posición y de la imagen desde el borde inferior en puntos

        # Crear un nuevo PDF temporal con la imagen usando reportlab
        c = canvas.Canvas(temp_pdf, pagesize=letter)
        # Ajustar el tamaño de la imagen
        imagen = ImageReader(m_agua)
        c.drawImage(imagen, posicion_x, posicion_y, width=imagen_ancho, height=imagen_alto)
        c.showPage()
        c.save()

        # Leer los PDFs
        lector_original = PyPDF2.PdfReader(archivo)
        lector_imagen = PyPDF2.PdfReader(temp_pdf)

        # Crear un escritor para el nuevo PDF
        escritor = PyPDF2.PdfWriter()

        # Combinar la primera página con la imagen como marca de agua
        pagina_original = lector_original.pages[0]
        pagina_imagen = lector_imagen.pages[0]

        # Crear una nueva página que combine la original y la imagen
        pagina_original.merge_page(pagina_imagen)
        escritor.add_page(pagina_original)

        # Añadir el resto de las páginas del PDF original
        for i in range(1, len(lector_original.pages)):
            escritor.add_page(lector_original.pages[i])

        # Guardar el nuevo PDF
        with open(os.path.join(aseg_dir, n_archivo), 'wb') as salida:
            escritor.write(salida)

        # Limpiar el archivo temporal
        os.remove(temp_pdf)
       
elif aseguradora == "QUA":
   
    # =============================================================================
    # QUA
    # =============================================================================

    for i in range(0, len(df_arch['archivos'])):
        #i = 0
        archivo = df_arch.iloc[i, 1]
        n_archivo = df_arch.iloc[i, 0]
        temp_pdf = aseg_dir + '/temp.pdf'
       
       
        if n_archivo[:2] == '79':
           
            # Tamaño y posición de la imagen para marca de Agua
            imagen_ancho = 200  # ancho de la imagen en puntos
            imagen_alto = 100   # alto de la imagen en puntos
            posicion_x = 200    # posición x de la imagen desde el borde izquierdo en puntos
            posicion_y = 280    # posición y de la imagen desde el borde inferior en puntos

            # Crear un nuevo PDF temporal con la imagen usando reportlab
            c = canvas.Canvas(temp_pdf, pagesize=letter)
            # Ajustar el tamaño de la imagen
            imagen = ImageReader(m_agua)
            c.drawImage(imagen, posicion_x, posicion_y, width=imagen_ancho, height=imagen_alto)
            c.showPage()
            c.save()

            # Leer los PDFs
            lector_original = PyPDF2.PdfReader(archivo)
            lector_imagen = PyPDF2.PdfReader(temp_pdf)

            # Crear un escritor para el nuevo PDF
            escritor = PyPDF2.PdfWriter()

            # Combinar la primera página con la imagen como marca de agua
            pagina_original = lector_original.pages[2]
            pagina_uno = lector_original.pages[0]
            pagina_dos = lector_original.pages[1]
            pagina_imagen = lector_imagen.pages[0]

            # Crear una nueva página que combine la original y la imagen
            pagina_original.merge_page(pagina_imagen)
            escritor.add_page(pagina_uno)
            escritor.add_page(pagina_dos)
            escritor.add_page(pagina_original)

            # Añadir el resto de las páginas del PDF original
            for i in range(3, len(lector_original.pages)):
                escritor.add_page(lector_original.pages[i])

            # Guardar el nuevo PDF
            with open(os.path.join(aseg_dir, n_archivo), 'wb') as salida:
                escritor.write(salida)

            # Limpiar el archivo temporal
            os.remove(temp_pdf)

        else:
           
            # Tamaño y posición de la imagen para marca de Agua
            imagen_ancho = 200  # ancho de la imagen en puntos
            imagen_alto = 100   # alto de la imagen en puntos
            posicion_x = 200    # posición x de la imagen desde el borde izquierdo en puntos
            posicion_y = 280    # posición y de la imagen desde el borde inferior en puntos

            # Crear un nuevo PDF temporal con la imagen usando reportlab
            c = canvas.Canvas(temp_pdf, pagesize=letter)
            # Ajustar el tamaño de la imagen
            imagen = ImageReader(m_agua)
            c.drawImage(imagen, posicion_x, posicion_y, width=imagen_ancho, height=imagen_alto)
            c.showPage()
            c.save()

            # Leer los PDFs
            lector_original = PyPDF2.PdfReader(archivo)
            lector_imagen = PyPDF2.PdfReader(temp_pdf)

            # Crear un escritor para el nuevo PDF
            escritor = PyPDF2.PdfWriter()

            # Combinar la primera página con la imagen como marca de agua
            pagina_original = lector_original.pages[1]
            pagina_uno = lector_original.pages[0]
            pagina_imagen = lector_imagen.pages[0]

            # Crear una nueva página que combine la original y la imagen
            pagina_original.merge_page(pagina_imagen)
            escritor.add_page(pagina_uno)
            escritor.add_page(pagina_original)

            # Añadir el resto de las páginas del PDF original
            for i in range(2, len(lector_original.pages)):
                escritor.add_page(lector_original.pages[i])

            # Guardar el nuevo PDF
            with open(os.path.join(aseg_dir, n_archivo), 'wb') as salida:
                escritor.write(salida)

            # Limpiar el archivo temporal
            os.remove(temp_pdf)
       
    else:
       
        # =============================================================================
        # AXA
        # =============================================================================
       
        for i in range(0, len(df_arch['archivos'])):
            #i = 0
            archivo = df_arch.iloc[i, 1]
            n_archivo = df_arch.iloc[i, 0]
            temp_pdf = aseg_dir + '/temp.pdf'
           
            # Tamaño y posición de la imagen para marca de Agua
            imagen_ancho = 300  # ancho de la imagen en puntos
            imagen_alto = 150   # alto de la imagen en puntos
            posicion_x = 180    # posición x de la imagen desde el borde izquierdo en puntos
            posicion_y = 160    # posición y de la imagen desde el borde inferior en puntos
   
            # Crear un nuevo PDF temporal con la imagen usando reportlab
            c = canvas.Canvas(temp_pdf, pagesize=letter)
            # Ajustar el tamaño de la imagen
            imagen = ImageReader(m_agua)
            c.drawImage(imagen, posicion_x, posicion_y, width=imagen_ancho, height=imagen_alto)
            c.showPage()
            c.save()
   
            # Leer los PDFs
            lector_original = PyPDF2.PdfReader(archivo)
            lector_imagen = PyPDF2.PdfReader(temp_pdf)
   
            # Crear un escritor para el nuevo PDF
            escritor = PyPDF2.PdfWriter()
   
            # Combinar la primera página con la imagen como marca de agua
            pagina_original = lector_original.pages[0]
            pagina_imagen = lector_imagen.pages[0]
   
            # Crear una nueva página que combine la original y la imagen
            pagina_original.merge_page(pagina_imagen)
            escritor.add_page(pagina_original)
   
            # Añadir el resto de las páginas del PDF original
            for i in range(1, len(lector_original.pages)):
                escritor.add_page(lector_original.pages[i])
   
            # Guardar el nuevo PDF
            with open(os.path.join(aseg_dir, n_archivo), 'wb') as salida:
                escritor.write(salida)
   
            # Limpiar el archivo temporal
            os.remove(temp_pdf)