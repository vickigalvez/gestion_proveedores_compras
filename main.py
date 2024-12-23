from config.db import create_connection, close_connection

def main():
    # 1. Establecer conexión a la BD
    connection = create_connection()

    if connection:
        print('conexion creada')
        try:
            # 2. Ejemplo: Ejecutar una consulta
            with connection.cursor() as cursor:
                if cursor:
                    print('cursor creado')
                    cursor.close()
        except Exception as e:
            print(f"Error al ejecutar la consulta: {e}")
        finally:
            close_connection(connection)

if __name__ == "__main__":
    main()