# Definir el tamaño del pliego disponible en el almacén y la cantidad disponible
tamano_almacen = [1500, 1050]
cantidad_almacen = 10

# Definir las dimensiones del pliego necesario para la impresión y la cantidad necesaria
tamano_pliego = [1060, 532]
cantidad_necesaria = 5

# Comprobar si el pliego cumple con las dimensiones y si hay suficiente cantidad disponible
if tamano_pliego[0] <= tamano_almacen[0] and tamano_pliego[1] <= tamano_almacen[1]:
    if cantidad_necesaria <= cantidad_almacen:
        print("Se puede utilizar el pliego disponible en el almacén para la impresión.")
    else:
        print("No hay suficiente cantidad de pliegos disponibles en el almacén para la impresión.")
else:
    print("El pliego disponible en el almacén no cumple con las dimensiones requeridas para la impresión.")

# Calcular el exceso o la falta de espacio en cada dimensión
exceso_ancho = tamano_pliego[0] - tamano_almacen[0]
exceso_alto = tamano_pliego[1] - tamano_almacen[1]

if exceso_ancho > 0:
    print(f"El pliego es {exceso_ancho} unidades más ancho que el pliego disponible en el almacén.")
elif exceso_ancho < 0:
    print(f"El pliego es {abs(exceso_ancho)} unidades más estrecho que el pliego disponible en el almacén.")
else:
    print("El pliego tiene el mismo ancho que el pliego disponible en el almacén.")

if exceso_alto > 0:
    print(f"El pliego es {exceso_alto} unidades más alto que el pliego disponible en el almacén.")
elif exceso_alto < 0:
    print(f"El pliego es {abs(exceso_alto)} unidades más bajo que el pliego disponible en el almacén.")
else:
    print("El pliego tiene la misma altura que el pliego disponible en el almacén.")
