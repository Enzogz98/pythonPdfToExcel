import aspose.pdf as ap
import os
import pandas as pd

def convertir(pdf_path,excel_path):
    pdf_document=ap.Document(pdf_path)
    excel_save_options=ap.ExcelSaveOptions()
    excel_save_options.format=ap.ExcelSaveOptions.ExcelFormat.XLSX
    excel_save_options.minimize_the_number_of_worksheets=True

    
    pdf_document.save(excel_path, excel_save_options)

    

def filtrar_datos(excel_path):
    df=pd.read_excel(excel_path, header=17)
    df_filtrado=df[(df['A']=='Param') | (df['A']=='Param')]
    output_filtered_excel_path = excel_path.replace('.xlsx', '_filtered.xlsx')
    df_filtrado.to_excel(output_filtered_excel_path, index=False)


def combinacion(source_folder, output_older,excel_folder,keywords):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    if not os.path.exists(excel_folder):
        os.makedirs(excel_folder)

    pdf_editor=ap.facades.PdfFileEditor()

    for keyword in keywords:
        files_to_merge=[]
        for filename in os.listdir(source_folder):
            if keyword in filename.upper() and filename.endswith('.pdf'):
                files_to_merge.append(os.path.join(source_folder,filename))

        if files_to_merge:
            output_pdf_path=os.path.join(output_folder, f"{keyword.replace(' ','_')}_combined.pdf")
            output_excel_path=os.path.join(excel_folder, f"{keyword.replace(' ','_')}_combined.xlsx")

            pdf_editor.concatenate(files_to_merge, output_pdf_path)

            convertir(output_pdf_path, output_excel_path)
            filtrar_datos(output_excel_path)
            
            print('archivo combinado para "{keyword}" convertido a excel y guardado')
        else:
            print('no se encontraron archivos')
   
source_folder=''
output_folder=''
excel_folder=''
keywords=['key_values']
combinacion(source_folder,output_folder,excel_folder,keywords)


