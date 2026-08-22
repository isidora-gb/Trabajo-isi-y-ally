def fibonacci(posicion):
	if posicion <= 1:
		return posicion
	return fibonacci(posicion - 1) + fibonacci(posicion - 2)


terminos = [fibonacci(posicion) for posicion in range(15)]

print("Primeros 15 términos:", terminos)
print("Suma:", sum(terminos))
