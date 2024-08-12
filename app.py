import pyodbc
import pandas as pd

# Ruta al archivo de Excel
excel_file_path = 'ListadoInstitucionesEducativasPublicas-2023-03-31.xlsx'

# Leer el archivo de Excel
df = pd.read_excel(excel_file_path)

# Configuración de conexión
server = 'DESKTOP-K2U11LG'  
database = 'dbqaliwarma'

# Verificar los primeros registros y los tipos de datos del DataFrame
print(df.head())
print(df.dtypes)

# Convertir tipos de datos si es necesario
df['CodigoIIEEQW'] = pd.to_numeric(df['CodigoIIEEQW'], errors='coerce')  # Convertir a numérico (float o int)
df['proveedor'] = df['proveedor'].astype(str)  # Asegurar que proveedor sea string

# Eliminar filas con valores nulos en las columnas que vamos a insertar
df = df.dropna(subset=['CodigoIIEEQW', 'proveedor'])

try:
    # Conectar a la base de datos
    with pyodbc.connect(f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;') as conexion:
        print("Conexión exitosa")
        
        with conexion.cursor() as cursor:
            # Iterar sobre el DataFrame e insertar cada fila en la tabla SQL
            for index, row in df.iterrows():
                id_modalidad = row['id_modalidad']
                ModalidadAtencion = row['ModalidadAtencion']  # Asegúrate de que este nombre coincida con la columna de tu archivo Excel
                CodigoIIEEQW = row['CodigoIIEEQW']
                
                try:
                    # Consulta SQL para insertar datos
                    insert_query = """
                    INSERT INTO Modalidad_atencion(id_modalidad, Modalidad, codigo_qali)
                    VALUES (?, ?, ?)
                    """
                    
                    cursor.execute(insert_query, (id_modalidad, ModalidadAtencion, CodigoIIEEQW))
                
                except Exception as e:
                    # Manejar el error e imprimir información sobre la fila problemática
                    print(f"Error al insertar la fila {index}: {e}")
            
            # Confirmar los cambios
            conexion.commit()
            print("Datos insertados correctamente")

except Exception as e:
    print(f"Error al conectar o insertar datos: {e}")
