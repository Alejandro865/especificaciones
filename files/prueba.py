from docxtpl import DocxTemplate

def generar_datos_tablas(cantidad_tablas):
    datos_tablas = []
    for i in range(cantidad_tablas):
        # Generar los datos para cada tabla
        # Puedes usar listas, diccionarios u otras estructuras de datos aquí
        datos_tabla = [
            {"celda1": "Dato 1", "celda2": "Dato 2", "celda3": "Dato 3"},
            {"celda1": "Dato 4", "celda2": "Dato 5", "celda3": "Dato 6"},
            # ... Puedes agregar más filas aquí ...
        ]
        datos_tablas.append(datos_tabla)
    return datos_tablas

def crear_documento_con_tablas(cantidad_tablas):
    # Cargar la plantilla del documento
    doc = DocxTemplate('.\Input\Template\plantilla.docx')

    # Generar los datos para las tablas
    datos_tablas = generar_datos_tablas(cantidad_tablas)

    # Reemplazar el marcador {{tabla}} con los datos generados
    for datos_tabla in datos_tablas:
        # Añadir una tabla al documento
        doc.render(datos_tabla)  # Renderizar la tabla con los datos
        doc.add_page_break()  # Añadir un salto de página después de cada tabla

    # Guardar el documento generado
    doc.save("documento_con_tablas.docx")

if __name__ == "__main__":
    cantidad_tablas_deseadas = 5  # Puedes cambiar este valor según tus necesidades
    crear_documento_con_tablas(cantidad_tablas_deseadas)
