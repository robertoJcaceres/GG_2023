from datetime import datetime
from tkinter import Tk, Label, Button, Entry, StringVar, Text, END
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.platypus import Image
import pyodbc
import pandas as pd
from reportlab.lib import colors
from tkinter import messagebox
from reportlab.lib.units import inch
import subprocess
from tkinter import Toplevel
import re
from PyPDF2 import PdfReader, PdfWriter
import os

class App:


    def __init__(self, root):

        self.root = root

        root.title("Generador de Ficha Técnica")

        self.entry_text = StringVar()

        self.extra_field1 = StringVar()

        self.extra_field2 = StringVar()

        self.label = Label(root, text="Ingrese el código del artículo")

        self.label.pack()

        self.entry = Entry(root, textvariable=self.entry_text)

        self.entry.pack()

        # self.extra_entry2 = Entry(root, textvariable=self.extra_field2)

        # self.extra_entry2.pack()

        self.label_output = Text(root, height=10, width=40)

        self.label_output.pack()

        self.button = Button(root, text="Obtener datos", command=self.get_data)

        self.button.pack()

        self.button2 = Button(root, text="Generar PDF", command=self.generate_pdf1)

        self.button2.pack() 

        self.df_FT = None  # Crear una variable de instancia para df_FT
 
    def wrap_text(self, text, max_chars):

        parts = text.split()
 
        lines = []
 
        line = ""

        for part in parts:

            if len(line) + len(part) < max_chars:

                line += " " + part

            else:

                lines.append(line.strip())

                line = part

        lines.append(line.strip())

        return '\n'.join(lines)

    def get_data(self):

        conn = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.2.20;DATABASE=SistradeERP;UID=GaleriaGrafica;PWD=Galeria123')

        query = f"""
WITH CTE AS (
    SELECT 

		 A.Desc_Cliente
		,A.Descripcion_Articulo
		,A.Codigo_Barras_2
		,A.Ref_Cliente
		,A.Cliente
		,A.Codigo_Articulo
		,A.Tipo_Tinta
		,a.Tamano_Estuche
		,A.Papel_Interno_Contracolado
		,A.Certif_Papel
		,A.Papel_Reverso_Contracolado
		,A.Gramaje
		,A.Tipo_Papel
		,A.Peso
		,A.Acabado_1_Barniz
		,A.Acabado_2_Barniz
		,A.Acabado_Peliculado
		,A.Acabado_Poliester
		,A.Acabado_Stamping
		,A.Detalle_1_Barniz
		,A.Detalle_2_Barniz
        ,A.coresAnverso
		,A.Tinta_Anverso_1
		,A.Tinta_Anverso_2
		,A.Tinta_Anverso_3
		,A.Tinta_Anverso_4
		,A.Tinta_Anverso_5
		,A.Tinta_Anverso_6
		,A.Tinta_Anverso_7
		,A.Tinta_Anverso_8
		,A.Tinta_Anverso_9
		,A.Colores_Reverso
		,A.Tinta_Reverso_1
		,A.Tinta_Reverso_2
		,A.Tinta_Reverso_3
		,A.Tinta_Reverso_4
		,A.Tinta_Reverso_5
		,A.Tinta_Reverso_6
		,A.Tinta_Reverso_7
		,A.Tinta_Reverso_8
		,A.Tinta_Reverso_9

		, B.*, C.*,
        ROW_NUMBER() OVER(PARTITION BY A.Desc_Troquel ORDER BY (SELECT NULL)) AS rn
    FROM [00GG_LISTADO_ARTICULOS_PA_GG] AS A
    LEFT JOIN [00GG_Listado_Troqueles_2] AS B ON A.Desc_Troquel = B.troquel_desc 
    LEFT JOIN [00GG_Listado_Troqueles_info_ad] AS C ON B.troquel_codigo = C.t_art
    WHERE A.Codigo_Articulo LIKE '{self.entry_text.get()}'
)
SELECT * FROM CTE
WHERE rn = 1;
  

            """
        self.df_FT = pd.read_sql(query, conn)

        conn.close()

        self.label_output.delete(1.0, END)

        for index, row in self.df_FT.iterrows():
            self.label_output.insert(END, f"\u2022 Cliente: {row['Cliente']}\n")
 
            self.label_output.insert(END, f"\u2022 Descripción del Artículo: {row['Descripcion_Articulo']}\n")
    def generate_pdf1(self):

        # Verificar si se han obtenido datos

        if self.df_FT is None:
            messagebox.showerror("Error", "Debes obtener datos primero.")
            return

        # Crear una ventana secundaria

        self.new_window = Toplevel(self.root)
        self.new_window.title("Campos adicionales")
        self.extra_field1 = StringVar()
        self.extra_field2 = StringVar()
        self.extra_field3 = StringVar()
        self.extra_field8 = StringVar()
        self.extra_field10 = StringVar()
        Label(self.new_window, text="Observaciones producto:").pack()
        Entry(self.new_window, textvariable=self.extra_field1).pack()

        Label(self.new_window, text="Observaciones Impresión:").pack()
        Entry(self.new_window, textvariable=self.extra_field2).pack()

        Label(self.new_window, text="MEDIDAS CAJAS EMBALAJES:").pack()
        Entry(self.new_window, textvariable=self.extra_field3).pack()


        Label(self.new_window, text="ALTURA POR PALET").pack()
        Entry(self.new_window, textvariable=self.extra_field8).pack()


        Label(self.new_window, text="TIPO PLEGADO").pack()
        Entry(self.new_window, textvariable=self.extra_field10).pack()


        Button(self.new_window, text="Aceptar", command=self.generate_pdf).pack()
        Button(self.new_window, text="Cancelar", command=self.new_window.destroy).pack()

    def generate_pdf(self):

        if self.df_FT is None:
            messagebox.showerror("Error", "Debes obtener datos primero.")

            return

        ObservacionesArticulo = self.extra_field1.get()
        ObservacionesImpresion= self.extra_field2.get()
        Tipo_de_Plegado= self.extra_field10.get() #la tiene que poner Alba
        Altura_por_Palet  = self.extra_field8.get()  #la tiene que poner Alba
        M_CajasEmbalajes  = self.extra_field3.get()  #no mi pana


        N_Estuches_Caja   = self.df_FT.iloc[0]['t2_val']
        Capas_Por_Palet   = self.df_FT.iloc[0]['t5_val']
        Cajas_Por_capa    = int(self.df_FT.iloc[0]['t3_val'])/int(self.df_FT.iloc[0]['t5_val'])  #t3_val/t5_val
        No_Estuches_Palet = self.df_FT.iloc[0]['t6_val'] #t6_val
        
        Tipo_de_Palet     = self.df_FT.iloc[0]['t1_val']

        
        Desc_Cliente= self.wrap_text(str(self.df_FT.iloc[0]['Desc_Cliente']), 20)

        current_date = datetime.now().strftime('%Y-%m-%d')
 
        wrapped_description = self.wrap_text(str(self.df_FT.iloc[0]['Descripcion_Articulo']), 10)

        wrapped_barcode = self.wrap_text(str(self.df_FT.iloc[0]['Codigo_Barras_2']), 20)

        Ref_Cliente = self.df_FT.iloc[0]['Ref_Cliente']

        CLIENTE = self.df_FT.iloc[0]['Cliente']

        codigo_articulo = self.df_FT.iloc[0]['Codigo_Articulo']  # Obtiene el valor de la columna 'Codigo_Articulo' del DataFrame
        Tipo_Tinta = self.df_FT.iloc[0]['Tipo_Tinta']  # Obtiene el valor de la columna 'Codigo_Articulo' del DataFrame

        tamano_estuche= self.df_FT.iloc[0]['Tamano_Estuche']  # Obtiene el valor de la columna 'Codigo_Articulo' del DataFrame
        
        Certificado_papel= self.df_FT.iloc[0]['Certif_Papel']  # Obtiene el valor de la columna 'Codigo_Articulo' del DataFrame
        
        papel_interno_Contracolado= self.df_FT.iloc[0]['Papel_Interno_Contracolado']  # Obtiene el valor de la columna 'Codigo_Articulo' del DataFrame
        papel_Reverso_Contracolado= self.df_FT.iloc[0]['Papel_Reverso_Contracolado']  # Obtiene el valor de la columna 'Codigo_Articulo' del DataFrame
 

        Gramaje= str(self.df_FT.iloc[0]['Gramaje'])  # Obtiene el valor de la columna 'Codigo_Articulo' del DataFrame

        Tipo_Papel = str(self.df_FT.iloc[0]['Tipo_Papel']) # Obtiene el valor de la columna 'Codigo_Articulo' del DataFrame

        Peso = str(self.df_FT.iloc[0]['Peso']) # Obtiene el valor de la columna 'Codigo_Articulo' del DataFrame

        # Supongamos que la nueva columna se llame 'Numero_Maqueta' (puedes cambiarlo cuando lo sepas)

        numero_maqueta = self.df_FT.iloc[0].get('ref', 'Aún no disponible')  # Obtiene el valor o muestra 'Aún no disponible'

        material =self.wrap_text(Tipo_Papel +" " + Gramaje , 20)
        material2 =self.wrap_text(Tipo_Papel +" " + Gramaje + " " + papel_interno_Contracolado + " " + papel_Reverso_Contracolado, 20)
        elements = [
            self.df_FT.iloc[0]['Acabado_1_Barniz'] +" " +self.df_FT.iloc[0]['Detalle_1_Barniz'],
            self.df_FT.iloc[0]['Acabado_2_Barniz'] +" " +self.df_FT.iloc[0]['Detalle_2_Barniz'],
            self.df_FT.iloc[0]['Acabado_Peliculado'],
            self.df_FT.iloc[0]['Acabado_Poliester'],
            self.df_FT.iloc[0]['Acabado_Stamping']
                ]

        ACABADO = " + ".join(filter(lambda x: x.strip() != "", elements))

        elements2 = [
            self.df_FT.iloc[0]['coresAnverso'],
               self.df_FT.iloc[0]['Tinta_Anverso_1'],
               self.df_FT.iloc[0]['Tinta_Anverso_2'],
               self.df_FT.iloc[0]['Tinta_Anverso_3'],
               self.df_FT.iloc[0]['Tinta_Anverso_4'],
               self.df_FT.iloc[0]['Tinta_Anverso_5'],
               self.df_FT.iloc[0]['Tinta_Anverso_6'],
               self.df_FT.iloc[0]['Tinta_Anverso_7'],
               self.df_FT.iloc[0]['Tinta_Anverso_8'],
               self.df_FT.iloc[0]['Tinta_Anverso_9']
        ]
        ANVERSO = " + ".join(filter(lambda x: x is not None and x.strip() != "", elements2))
        elements3 = [

            self.df_FT.iloc[0]['Colores_Reverso'],

               self.df_FT.iloc[0]['Tinta_Reverso_1'],

               self.df_FT.iloc[0]['Tinta_Reverso_2'],

               self.df_FT.iloc[0]['Tinta_Reverso_3'],

               self.df_FT.iloc[0]['Tinta_Reverso_4'],

               self.df_FT.iloc[0]['Tinta_Reverso_5'],

               self.df_FT.iloc[0]['Tinta_Reverso_6'],

               self.df_FT.iloc[0]['Tinta_Reverso_7'],

               self.df_FT.iloc[0]['Tinta_Reverso_8'],

               self.df_FT.iloc[0]['Tinta_Reverso_9']

        ]

        REVERSO = " + ".join(filter(lambda x: x is not None and x.strip() != "", elements3))

# Ahora, `final_string` contiene la concatenación deseada sin signos " + " adicionales.

        ACABADO =self.wrap_text(ACABADO,90)

        ANVERSO =self.wrap_text(ANVERSO,90)

        REVERSO=self.wrap_text(REVERSO,90)

        TEXT_OBSERVACION = """Tiempo recomendado de consumo: 
        Estuchado automático: 6 meses. 
        Estuchado manual: 1 año. 
        Envase no reutilizable para contenido de productos alimenticios. 
        Los alimentos no deben entrar en contacto con las superficies impresas. 
        Óptimas condiciones de almacenamiento: 65-70% máximo HR, sin cambios bruscos de temperatura y en ausencia física de agua. 
        Apilar en estanterías y no pallet sobre pallet."""

        desc = self.df_FT.iloc[0]['Descripcion_Articulo']
        desc = desc.replace(" ", "_")  # Reemplaza los espacios en blanco con guiones bajos
        desc = desc.replace("/", "-")  # Reemplaza las barras con guiones
        desc = desc.replace("*", "_")  # Reemplaza las barras con guiones
        desc = desc.replace(".", "_")  # Reemplaza las barras con guiones
        pdf_filename = f"FT_{desc}.pdf"
        pdf = SimpleDocTemplate(
            pdf_filename,
            pagesize=letter, 
            topMargin=0.5*inch       )

        logo = Image("logo.png", 1.3*inch, 0.6*inch)
        
        # Creación del diccionario con los materiales y sus fibras correspondientes

        materiales_fibra = {
            "Folding Madera (GC2)-": "FIBRA VIRGEN",
            "CARTULINA GRAFICA - SBS - 2 CARA": "FIBRA VIRGEN",
            "Estucado Rev. Madera (GT2)": "FIBRA RECICLADA",
            "ONDULADOS - GUITARRA": "FIBRA RECICLADA",
            "OTROS PAPELES - COUCHÉ MATE - 2 CARAS": "FIBRA VIRGEN",
            "ONDULADOS - MICRO": "FIBRA RECICLADA",
            "ONDULADOS - NANOMICRO": "FIBRA RECICLADA",
            "OTROS PAPELES - COMPACTO": "FIBRA RECICLADA",
            "Kraft Blanco (CKB)": "FIBRA VIRGEN",
            "OTROS PAPELES - ESPECIALIDADES": None,
            "CARTULINA GRAFICA - SBS - 1 CARA": "FIBRA VIRGEN",
            "Estucado Rev. Blanco (GT1)": "FIBRA RECICLADA",
            "Kraft Marron (CKB)": "FIBRA VIRGEN",
            "OTROS PAPELES - VERJURADO - BLANCO": "FIBRA VIRGEN",
            "ONDULADOS - MICRO COSMETICA NEGRO": None,
            "CARTON CLIENTE": None,
            "OTROS PAPELES - COUCHÉ BRILLO - 2 CARAS": "FIBRA VIRGEN",
            "Estucado Rev. Gris (GD)": "FIBRA RECICLADA",
            "OTROS PAPELES - PAPEL OFFSET": "FIBRA VIRGEN",
            "OTROS PAPELES": None,
            "ONDULADOS - MICRO COSMETICA BLANCO": None,
            "Folding Blanco (GC1)": "FIBRA VIRGEN",
            "OTROS PAPELES - COUCHÉ BRILLO - 1 CARA": "FIBRA VIRGEN"
        }
        # Función para buscar la fibra de un material dado
        def buscar_fibra(material, diccionario):
            return diccionario.get(material, "Material no encontrado")

        # Ejemplo de uso

        fibra = buscar_fibra(Tipo_Papel, materiales_fibra)



        data = [
            [logo, CLIENTE],
            ["FECHA:", current_date],
            ["DATOS IDENTIFICATIVOS DEL PRODUCTO"],
            ["DESCRIPCIÓN:", self.wrap_text(wrapped_description,60)],
            ["CÓDIGO DE BARRAS:", wrapped_barcode],
            ["REFERENCIA INTERNA:", codigo_articulo, "Nº MAQUETA:", numero_maqueta],
            ["REFERENCIA CLIENTE:", Desc_Cliente],
            ["TAMAÑO (mm)", tamano_estuche, "TOLERANCIA:", " ± 0,5 mm"],
            ["MATERIAL:", material2 , "PESO KG/UD:", Peso + " ± 3%"],
            ["CERTIFICACIÓN",Certificado_papel, "PROVENIENCIAS :",fibra],
            ["TIPO DE PLEGADO:", Tipo_de_Plegado],
            ["OBSERVACIONES:", self.wrap_text(ObservacionesArticulo,30)],
            ["IMPRESIÓN"],
            ["TIPO:", Tipo_Tinta +"-"+ " "+'OFFSET'],
            ["ANVERSO Nº TINTAS:", ANVERSO],
            ["REVERSO Nº TINTAS:", REVERSO],
            ["ACABADO:", ACABADO],
            ["OBSERVACIONES:", self.wrap_text(ObservacionesImpresion,30)],
            ["LOGÍSTICA"],
            ["MEDIDAS CAJAS EMBALAJES:", M_CajasEmbalajes],
            ["Nº ESTUCHES POR CAJA:", N_Estuches_Caja],
            ["CAPAS POR PALET:", Capas_Por_Palet],
            ["CAJAS POR CAPA:", Cajas_Por_capa],
            ["Nº ESTUCHES POR PALET:",No_Estuches_Palet],
            ["ALTURA POR PALET:", Altura_por_Palet],
            ["TIPO DE PALET:", Tipo_de_Palet],
            ["OBSERVACIONES:",self.wrap_text(TEXT_OBSERVACION,80) ]
        ]

        table = Table(data, colWidths=[1.9*inch,  1.5*inch, 1.5*inch])
       # Estilos de fusión de celdas
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 1), (0, 26), colors.lightgrey),  # Todos 
            ('FONTSIZE', (0, 1), (-1, 26),8), #OBSERVACIONE
            ('ALIGN', (0, 1), (0, 26), 'RIGHT'),  # Todos 
             ('VALIGN', (0, 0), (-1, -1),'MIDDLE'), # Todos
            
            ('ALIGN', (0, 0), (1, 0), 'CENTER'),  #Encabezado
            ('SPAN', (1, 0), (3, 0)),  #Encabezado
            ('FONTSIZE', (0, 0), (-1, 0),14), #Encabezado
           
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), #Encabezado

            
            ('SPAN', (1, 1), (3, 1)),  #Cliente
            ('BACKGROUND', (0, 1), (0, 1), colors.lightgrey), #Cliente 
            ('ALIGN', (0, 1), (3, 1), 'CENTER'),  #Cliente 
           
            
            ('SPAN', (0, 2), (-1, 2)),  #Datos Identificativos del Producto
            ('BACKGROUND', (0, 2), (-1, 2), colors.lightgrey), #Datos Identificativos del Producto
            ('ALIGN', (0, 2), (-1, 2), 'CENTER'), #Datos Identificativos del Producto
            ('FONTNAME', (0, 2), (0, 2), 'Helvetica-Bold'),

            
            ('SPAN', (1, 3), (-1, 3)),  # Descripción
            ('ALIGN', (1, 3), (-1, 3), 'CENTER'), # Descripción
            
            ('SPAN', (1, 4), (-1, 4)), #Codigo de barras
            ('ALIGN', (1, 4), (-1, 4), 'CENTER'),  #Codigo de barras    
            
            ('ALIGN', (1, 5), (1, 5), 'CENTER'),  # "REFERENCIA GALERÍA GRÁFICA"
            ('ALIGN', (2, 5), (2, 5), 'RIGHT'),  #"Nº MAQUETA:"
            ('BACKGROUND', (2, 5), (2, 5), colors.lightgrey),
            
            ('SPAN', (1, 6), (-1, 6)),  # REFERENCIA CLIENTE
            ('ALIGN', (1, 6), (1, 6), 'CENTER'), # REFERENCIA CLIENTE
            
            
            ('ALIGN', (0, 7), (0, 7), 'RIGHT'), # TAMAÑO
            ('ALIGN', (1, 7), (1, 7), 'CENTER'), # TAMAÑO
            
            ('ALIGN', (3, 7), (3, 7), 'CENTER'), # TOLERANCIA
            ('ALIGN', (2, 7), (2, 7), 'RIGHT'), # TOLERANCIA
            ('BACKGROUND', (2, 7), (2, 7), colors.lightgrey), # TOLERANCIA
            
            ('ALIGN', (0, 8), (0, 8), 'RIGHT'),  # MATERIAL
            
            ('ALIGN', (2, 8), (2, 8), 'RIGHT'),  # PESO KG/UD
            ('BACKGROUND', (2, 8), (2, 8), colors.lightgrey), # PESO KG/UD
            ('ALIGN', (3, 8), (3, 8), 'CENTER'), # PESO KG/UD
            
            ('ALIGN', (0, 8), (0, 8), 'RIGHT'),  # CERTIFICACION
            
            ('ALIGN', (2, 9), (2, 9), 'RIGHT'),  # PROVENIENCIAS
            ('BACKGROUND', (2, 9), (2, 9), colors.lightgrey), # PROVENIENCIAS
            ('ALIGN', (3, 9), (3, 9), 'CENTER'), # PROVENIENCIAS

            ('SPAN', (1, 10), (-1, 10)),  # TIPO DE PLEGADO
            
            ('SPAN', (1, 11), (-1, 11)),  # OBSERVACIONES
            
            ('SPAN', (0, 12), (-1, 12)),  # IMPRESION
            ('ALIGN', (0, 12), (-1, 12), 'CENTER'),  # IMPRESION
            ('FONTNAME', (0, 12), (-1, 12), 'Helvetica-Bold'), # LOGÍSTICA
             ('BACKGROUND', (0, 12), (-1, 12), colors.lightgrey), #Cliente 
            
             ('SPAN', (1, 13), (-1, 13)), # TIPO
             
             ('SPAN', (1, 14), (-1, 14)), # ANVERSO Nº TINTAS:
             
             ('SPAN', (1, 15), (-1, 15)), # REVERSO Nº TINTAS:
             
             ('SPAN', (1, 16), (-1, 16)), # ACABADO:
             
             ('SPAN', (1, 17), (-1, 17)), # OBSERVACIONES
             
             ('ALIGN', (0, 18), (-1, 18), 'CENTER'),  # LOGÍSTICA
             ('SPAN', (0, 18), (-1, 18)),  # LOGÍSTICA
             ('FONTNAME', (0, 18), (-1, 18), 'Helvetica-Bold'), # LOGÍSTICA
             ('BACKGROUND', (0, 18), (-1, 18), colors.lightgrey), #Cliente 
            

            
            ('SPAN', (1, 19), (-1, 19)), #MEDIDAS CAJAS EMBALAJES
            
            ('SPAN', (1, 20), (-1, 20)), # Nº ESTUCHES POR CAJA
            
            ('SPAN', (1, 21), (-1, 21)), # CAPAS POR PALET
            
            ('SPAN', (1, 22), (-1, 22)), # CAJAS POR CAPA
            
            ('SPAN', (1, 23), (-1, 23)), # Nº ESTUCHES POR PALET
            
            ('SPAN', (1, 24), (-1,24)), # ALTURA POR PALET
            
            ('SPAN', (1, 25), (-1,25)), # TIPO DE PALET
            
            ('SPAN', (1, 26), (-1, 26)), # OBSERVACIONES
            
            ('ALIGN', (0, 1), (0, 1), 'RIGHT'),  # fecha
            ('FONTSIZE', (1, 26), (-1, 26),6), #OBSERVACIONES
            
            ('GRID', (0, 0), (-1, -1), 1, colors.black)  # Aplicar los bordes de nuevo
        ]))



        elems = []

        elems.append(table)

        pdf.build(elems)
        
        # Rutas de los archivos PDF
        pdf_1 = pdf_filename
        pdf_2 = "Certificado.pdf"
        pdf_salida = pdf_filename

        # Crear un objeto PdfFileWriter para la salida
        pdf_writer = PdfWriter()

        # Función para añadir páginas de un PDF al PdfFileWriter
        def anadir_paginas(pdf_path, pdf_writer):
            pdf_reader = PdfReader(pdf_path)
            for page in range(len(pdf_reader.pages)):
                pdf_writer.add_page(pdf_reader.pages[page])

        # Añadir páginas de ambos PDFs
        anadir_paginas(pdf_1, pdf_writer)
        anadir_paginas(pdf_2, pdf_writer)

        # Guardar el PDF combinado
        with open(pdf_salida, 'wb') as out_file:
            pdf_writer.write(out_file)

        self.new_window.destroy()  # Cierra la ventana secundaria
       
        subprocess.Popen([pdf_filename], shell=True)
        
root = Tk()
app = App(root)
root.mainloop()