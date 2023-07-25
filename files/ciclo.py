def ciclo_una_vez():
    # Inicializamos una variable para contar el número de veces que se ejecuta el ciclo
    contador = 0

    while True:  # Ciclo infinito
        # Tu código dentro del ciclo
        print("Este es el ciclo número:", contador + 1)

        contador += 1  # Incrementamos el contador

        if contador == 1:  # Si el contador es igual a 1
            break  # Salimos del ciclo

ciclo_una_vez()
