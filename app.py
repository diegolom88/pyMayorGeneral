import os
import pandas as pd

def combine_excel_files(folder_path):
    # List to store all dataframes
    all_data = []

    # Sort files in ascending order
    for filename in sorted(os.listdir(folder_path)):
        if filename.endswith('.xlsx'):
            print(filename)
            # Create full file path
            file_path = os.path.join(folder_path, filename)

            # Read the Excel file
            df = pd.read_excel(file_path)

            # add file name to df
            df['Archivo'] = filename

            # Append to our list
            all_data.append(df)

    # Combine all dataframes
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        # Format the "Fecha" column to "yyyy-mm-dd"
        combined_df['Fecha'] = pd.to_datetime(combined_df['Fecha']).dt.strftime('%Y-%m-%d')
        print("RETURNING COMBINED DF")
        return combined_df
    else:
        return None

def add_maquinaria_categorization(df):
    # open maquinaria categorization file in sheet "CategorizacionMaquinaria"
    maquinaria_categorization_df = pd.read_excel("/mnt/c/Users/pc/Desktop/ba-files/DYCUSA/Datos/CategorizacionMaquinaria.xlsx", sheet_name="CategorizacionMaquinaria")

    # remove duplicates in column Conc
    maquinaria_categorization_df = maquinaria_categorization_df.drop_duplicates(subset=["Conc"])

    # Create a mapping dictionary from Conc to Tipo and Categoria
    maquinaria_categorization_dict = maquinaria_categorization_df.set_index("Conc").to_dict(orient="index")

    # Initialize blank df to store new rows
    new_maquinaria_categorization_df = pd.DataFrame(columns=["Conc"])

    # Categorize the combined mayor with maquinaria data
    for index, row in df.iterrows():
        if row["Conc"] in maquinaria_categorization_dict:
            foundRow = maquinaria_categorization_dict[row["Conc"]]
            df.at[index, "TipoMaquinaria"] = foundRow.get("Tipo", None)
            df.at[index, "CategoriaMaquinaria"] = foundRow.get("Categoria", None)
        else:
            # Add a new row to the new_maquinaria_categorization_df dataframe
            new_maquinaria_categorization_df = pd.concat([new_maquinaria_categorization_df, pd.DataFrame([{"Conc": row["Conc"]}])], ignore_index=True)

    # Save the new_maquinaria_categorization_df dataframe to a new Excel file
    new_maquinaria_categorization_df.to_excel("/mnt/c/Users/pc/Desktop/ba-files/DYCUSA/Datos/CategorizacionMaquinaria.xlsx", sheet_name="CategorizacionMaquinaria", index=False)

    return df

# Correr el programa
if __name__ == "__main__":
    folder_path = "/mnt/c/Users/pc/Desktop/ba-files/DYCUSA/Mayores/MayoresSis"
    output_folder = "/mnt/c/Users/pc/Desktop/ba-files/DYCUSA/Mayores/MayorAcumDYCUSA"
    combined_df = combine_excel_files(folder_path)
    result = add_maquinaria_categorization(combined_df)

    if result is not None:
        # Create output folder if it doesn't exist
        os.makedirs(output_folder, exist_ok=True)

        # Save combined data to a new Excel file in the specified output folder
        output_path = os.path.join(output_folder, "MayorAcumDYCUSA.xlsx")
        print("SAVING COMBINED DATA TO EXCEL")
        result.to_excel(output_path, index=False, sheet_name="MayorAcumDYCUSA", engine="openpyxl")

        # Auto-adjust column widths
        from openpyxl import load_workbook
        wb = load_workbook(output_path)
        ws = wb["MayorAcumDYCUSA"]

        # Auto-adjust column widths based on content
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
            ws.column_dimensions[column_letter].width = adjusted_width
        wb.save(output_path)
        wb.close()
        print("FINISHED SAVING COMBINED DATA TO EXCEL")
        print(f"Files combined successfully! Saved to: {output_path}")
    else:
        print("No Excel files found in the specified folder.")
