import openpyxl
from tabulate import tabulate
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Variables globales para encabezados y alineacion de tabulate (Paso 3)
encabezados = ["ID", "Producto", "Categoría", "Precio ($)", "Ventas (Unid)"]
alineacion = ("center", "left", "center", "right", "right")


# Paso 1 y 2: Cargar el archivo de Excel y leer registros dinamicamente
def cargar_datos():
    # 1. Cargar el libro de trabajo existente
    wb = openpyxl.load_workbook("ventas_tecnologia_v2.xlsx", data_only=True)

    # 2. Seleccionar la hoja activa (primera pestaña)
    hoja = wb.active
    print(f"Hoja cargada exitosamente: {hoja.title}")

    # Crear una lista vacia para almacenar las filas
    datos = []

    # Recorrer desde la fila 2 (omitimos encabezados de la fila 1)
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        datos.append(list(fila))

    print(f"Total de registros leidos en memoria: {len(datos)}")
    return datos


# Reto 2: Calcular suma total de unidades vendidas y precio promedio
def calcular_metricas(lista_datos):
    suma_unidades = 0
    suma_precios = 0

    for fila in lista_datos:
        suma_precios += fila[3]
        suma_unidades += fila[4]

    promedio_precio = suma_precios / len(lista_datos)
    return suma_unidades, promedio_precio


# Paso 3 y Reto 2: Mostrar tabla formateada en consola y metricas al final
def mostrar_consola(lista_datos):
    print(f"\n=== REPORTE DE VENTAS EN CONSOLA ({len(lista_datos)} PRODUCTOS) ===")
    print(tabulate(lista_datos, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))

    # Mostrar metricas del Reto 2 al final de la consola
    suma_unidades, promedio_precio = calcular_metricas(lista_datos)
    print(f"\nSuma total de unidades vendidas: {suma_unidades}")
    print(f"Precio promedio de los productos: ${promedio_precio:.2f}")


# Reto 1: Filtrado por criterio de precio (> $100)
def filtrar_por_precio(datos):
    datos_filtrados = []
    for fila in datos:
        if fila[3] > 100:
            datos_filtrados.append(fila)
    return datos_filtrados


# Reto 3: Busqueda dinamica por ID de producto
def buscar_por_id(datos):
    id_buscado = int(input("\nIngrese el numero de ID del producto a buscar: "))
    encontrado = []

    for fila in datos:
        if fila[0] == id_buscado:
            encontrado.append(fila)
            break

    if len(encontrado) > 0:
        print("\nProducto encontrado:")
        print(tabulate(encontrado, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))
    else:
        print(f"\nAdvertencia: El producto con ID {id_buscado} no existe en la base de datos.")


# Paso 4 y Reto 2: Generacion del informe ejecutivo en PDF con ReportLab
def exportar_pdf(lista_datos):
    # 1. Crear documento base
    doc = SimpleDocTemplate("reporte_ejecutivo_ventas.pdf", pagesize=letter)
    estilos = getSampleStyleSheet()

    # 2. Estilos personalizados
    estilo_titulo = ParagraphStyle(
        'EstiloTitulo',
        parent=estilos['Title'],
        fontName='Helvetica-Bold',
        fontSize=18,
        textColor=colors.HexColor("#1A365D"),
        spaceAfter=10
    )
    estilo_cuerpo = ParagraphStyle('EstiloCuerpo', parent=estilos['Normal'], fontSize=9)
    estilo_encabezado = ParagraphStyle(
        'EstiloEncabezado',
        parent=estilos['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        textColor=colors.whitesmoke,
        alignment=1
    )

    # 3. Lista story (carrito de compras)
    story = [
        Paragraph("Informe Ejecutivo de Ventas Tecno", estilo_titulo),
        Paragraph(f"<b>Total de Productos Procesados:</b> {len(lista_datos)}", estilo_cuerpo),
        Spacer(1, 12)
    ]

    # 4. Convertir celdas a Paragraph
    encabezados_pdf = [Paragraph(f"<b>{h}</b>", estilo_encabezado) for h in encabezados]
    filas_pdf = [
        [
            Paragraph(str(r[0]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
            Paragraph(str(r[1]), ParagraphStyle('L', parent=estilo_cuerpo, alignment=0)),
            Paragraph(str(r[2]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
            Paragraph(f"${r[3]:.2f}", ParagraphStyle('R', parent=estilo_cuerpo, alignment=2)),
            Paragraph(str(r[4]), ParagraphStyle('R', parent=estilo_cuerpo, alignment=2))
        ]
        for r in lista_datos
    ]

    # Reto 2: Agregar fila final resumida en la tabla del PDF
    suma_unidades, promedio_precio = calcular_metricas(lista_datos)
    fila_resumen = [
        Paragraph("<b>-</b>", ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
        Paragraph("<b>RESUMEN (PROMEDIO / TOTAL)</b>", ParagraphStyle('L', parent=estilo_cuerpo, alignment=0)),
        Paragraph("<b>-</b>", ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
        Paragraph(f"<b>${promedio_precio:.2f}</b>", ParagraphStyle('R', parent=estilo_cuerpo, alignment=2)),
        Paragraph(f"<b>{suma_unidades}</b>", ParagraphStyle('R', parent=estilo_cuerpo, alignment=2))
    ]
    filas_pdf.append(fila_resumen)

    # 5. Configurar anchos y tabla
    anchos_cols = [40, 160, 100, 80, 80]
    tabla_pdf = Table([encabezados_pdf] + filas_pdf, colWidths=anchos_cols, repeatRows=1)

    # 6. Aplicar estilo visual con TableStyle
    tabla_pdf.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2E8F0"))
    ]))

    story.append(tabla_pdf)

    # 7. Compilar PDF
    doc.build(story)
    print("\n[EXITO] Archivo 'reporte_ejecutivo_ventas.pdf' generado correctamente.")


# Reto 4: Menu interactivo de consola con bucle while True
def menu():
    datos = cargar_datos()
    datos_actuales = datos

    while True:
        print("\n--- MENU PRINCIPAL TECNOVENTAS ---")
        print("1) Ver catalogo completo")
        print("2) Filtrar por precio (> $100)")
        print("3) Buscar por ID")
        print("4) Exportar informe PDF")
        print("5) Salir del programa")

        opcion = input("Seleccione una opcion (1-5): ")

        if opcion == "1":
            datos_actuales = datos
            mostrar_consola(datos_actuales)

        elif opcion == "2":
            datos_actuales = filtrar_por_precio(datos)
            mostrar_consola(datos_actuales)

        elif opcion == "3":
            buscar_por_id(datos)

        elif opcion == "4":
            exportar_pdf(datos_actuales)

        elif opcion == "5":
            print("Saliendo del programa...")
            break

        else:
            print("Opcion no valida, intente de nuevo.")


menu()
