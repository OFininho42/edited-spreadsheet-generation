import os
import openpyxl as opxl

def try_open_file(path, file):

    load_file_status = False

    complete_path_file = os.path.join(path, file)

    try:

        uploaded_file = opxl.load_workbook(complete_path_file)

        load_file_status = uploaded_file

    except:

        load_file_status = load_file_status

    return load_file_status


def copy_spreadsheet_content(wb, sheet1, sheet2, save_path):

    if sheet1 not in wb.sheetnames:

        print(f"Erro: A aba de origem '{sheet1}' não existe no arquivo.")

        return

    if sheet2 in wb.sheetnames:

        del wb[sheet2]

    ws_origem = wb[sheet1]

    ws_destino = wb.copy_worksheet(ws_origem)

    ws_destino.title = sheet2

    wb.save(save_path)

    print(f"Conteúdo de '{sheet1}' copiado com sucesso para '{sheet2}' em '{save_path}'!")


file_path = input("O caminho:\n")

file_used = input("Nome do arquivo:\n")

origin_sheet = input("Planilha origem:\n")

destination_sheet = input("Planilha destino:\n")

complete_path = os.path.join(file_path, file_used)

spreadsheets_found = []

if try_open_file(file_path, file_used) != False:

    workbook = try_open_file(file_path, file_used)

    print("Caminho ou arquivo encontrado:")

    for spreadsheets in workbook.sheetnames:

        spreadsheets_found.append(spreadsheets) 

else: 

    print("Caminho ou arquivo não encontrado:")   


while True:

    if origin_sheet in spreadsheets_found and destination_sheet in spreadsheets_found:

        copy_spreadsheet_content(workbook, origin_sheet, destination_sheet, complete_path)
        
        break

    else: 

        print("Açaão não foi executada.")

        break