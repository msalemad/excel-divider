import os
import pandas as pd
from tqdm import tqdm
import time


def get_brand_and_product(cell_value):
    """
    Divide el contenido de una celda de texto en una marca y un nombre de producto.
    La marca corresponde a la primera palabra antes del primer espacio, y el producto
    es el resto del texto.
    """
    words = str(cell_value).strip().split(maxsplit=1)
    brand = words[0] if words else "Unknown"
    product = words[1] if len(words) > 1 else ""
    return brand, product


def split_excel_by_brand(input_file, output_folder):
    start_time = time.time()
    if not os.path.exists(input_file):
        print(f"Error: El archivo '{input_file}' no existe.")
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    try:
        print(f"Cargando archivo: {input_file}... Esto puede tardar un momento.")
        df = pd.read_excel(input_file, header=1)
        print(f"Archivo cargado con {len(df)} filas.")
    except Exception as e:
        print(f"Error al cargar el archivo: {e}")
        return

    if df.shape[1] < 3:
        print("Error: El archivo no tiene al menos tres columnas.")
        return

    column_index = 2  # Índice de la tercera columna (columna C)
    column_name = df.columns[column_index]

    brand_data = {}
    print("Dividiendo las filas por marcas...")
    for index, row in tqdm(df.iterrows(), total=len(df)):
        brand, product = get_brand_and_product(row[column_name])
        if brand not in brand_data:
            brand_data[brand] = []
        new_row = row.copy()
        new_row[column_name] = product  # Actualiza la columna con solo el nombre del producto
        brand_data[brand].append(new_row)

    print("Guardando los archivos por marca...")
    total_brands = len(brand_data)
    for i, (brand, rows) in enumerate(brand_data.items(), 1):
        output_path = os.path.join(output_folder, f"{brand}.xlsx")
        brand_df = pd.DataFrame(rows)
        brand_df.to_excel(output_path, index=False)
        tqdm.write(f"[{i}/{total_brands}] Archivo generado para la marca: '{brand}'")

    elapsed_time = time.time() - start_time
    print(f"\nOperación completada en {elapsed_time:.2f} segundos.")
    print(f"Se generaron {total_brands} archivos en '{output_folder}'")

    print("Presione Enter para cerrar o escriba 'reiniciar' para reiniciar.")
    user_input = input().strip().lower()
    if user_input == 'reiniciar':
        main()
    else:
        input("Presione Enter para finalizar y ver todas las líneas...")


def main():
    print("=== Script para dividir Excel por marcas ===")
    input_folder = "input"
    output_folder = "output"

    if not os.path.exists(input_folder):
        os.makedirs(input_folder)
        print(f"La carpeta '{input_folder}' fue creada. Coloque el archivo fuente dentro de esta carpeta y vuelva a ejecutar.")
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    input_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx')]
    # Este bloque maneja la situación cuando no hay archivos en la carpeta de entrada.
    # Proporciona un mensaje claro para mejorar la experiencia del usuario.
    if not input_files:
        print(f"No se encontraron archivos .xlsx en la carpeta '{input_folder}'.")
        return

    print("Archivos encontrados:")
    for i, file in enumerate(input_files, 1):
        print(f"[{i}] {file}")

    file_choice = int(input("Seleccione el número del archivo a procesar: ")) - 1
    if file_choice < 0 or file_choice >= len(input_files):
        print("Opción inválida.")
        return

    input_file = os.path.join(input_folder, input_files[file_choice])
    split_excel_by_brand(input_file, output_folder)


if __name__ == "__main__":
    main()