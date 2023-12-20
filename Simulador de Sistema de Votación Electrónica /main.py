import hashlib
import re

# Función para crear un hash de votos
def create_vote_hash(vote):
    # Crear un hash del voto usando SHA-256
    hash_object = hashlib.sha256(vote.encode())
    return hash_object.hexdigest()

# Función para validar el formato del DNI
def validate_dni(dni):
    # Verificar si el DNI tiene el formato correcto: 8 números seguidos de una letra
    return bool(re.match(r'^\d{8}[A-Za-z]$', dni))

# Simular una votación
def main():
    candidates = ["VOX", "PSOE", "SUMAR"]

    print("Bienvenido a la votación electrónica.")
    dni = input("Por favor, ingrese su número de DNI: ")

    if validate_dni(dni):
        print("Por favor, elija su candidato:")
        for i, candidate in enumerate(candidates, start=1):
            print(f"{i}. {candidate}")

        vote = input("Ingrese el número del candidato elegido: ")

        try:
            vote = int(vote)
            if 1 <= vote <= len(candidates):
                selected_candidate = candidates[vote - 1]
                hashed_vote = create_vote_hash(selected_candidate + dni)  # Se agrega el DNI al voto antes de hacer el hash
                print("Su voto ha sido registrado de manera segura.")
                print(f"Hash de su voto: {hashed_vote}")
            else:
                print("Por favor, ingrese un número válido correspondiente al candidato.")
        except ValueError:
            print("Por favor, ingrese un número válido para el candidato.")
    else:
        print("El número de DNI ingresado no es válido o no tiene el formato adecuado (8 números seguidos de una letra).")

if __name__ == "__main__":
    main()
