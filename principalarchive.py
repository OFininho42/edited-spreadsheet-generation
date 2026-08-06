import os
from pathlib import Path
import openpyxl as opxl

class ExcelHandler:

    def __init__(self, folder_path: str, file_name: str):
        self.file_path = Path(folder_path) / file_name
        self.wb = self._load_workbook()

    def _load_workbook(self):
        # Tenta carregar a pasta de trabalho de forma segura.
        if not self.file_path.exists():
            print(f"❌ Erro: O arquivo '{self.file_path}' não foi encontrado.")
            return None
        try:
            return opxl.load_workbook(self.file_path)
        except Exception as e:
            print(f"❌ Erro ao abrir o arquivo Excel: {e}")
            return None

    @property
    def is_valid(self) -> bool:
        # Retorna se o arquivo foi carregado com sucesso.
        return self.wb is not None

    def get_sheet_names(self) -> list[str]:
        # Retorna os nomes das abas disponíveis no arquivo.
        return self.wb.sheetnames if self.is_valid else []

    def copy_sheet(self, origin: str, destination: str) -> bool:
        # Copia o conteúdo de uma aba para outra no mesmo arquivo.
        if not self.is_valid:
            return False

        if origin not in self.wb.sheetnames:
            print(f"❌ A aba de origem '{origin}' não existe.")
            return False

        # Se a aba destino já existir, remove antes de recriar
        if destination in self.wb.sheetnames:
            del self.wb[destination]

        ws_origin = self.wb[origin]
        ws_dest = self.wb.copy_worksheet(ws_origin)
        ws_dest.title = destination

        print(f"✔️ Conteúdo de '{origin}' copiado para '{destination}'.")
        return True

    def clean_sheet(self, sheet_name: str) -> bool:
        # Desmescla todas as células e remove a quebra automática de texto.
        if not self.is_valid or sheet_name not in self.wb.sheetnames:
            print(
                f"❌ Aba '{sheet_name}' não encontrada para tratamento."
            )
            return False

        ws = self.wb[sheet_name]

        # 1. Desmesclar células
        merged_ranges = list(ws.merged_cells.ranges)
        for cell_range in merged_ranges:
            ws.unmerge_cells(str(cell_range))

        # 2. Remover quebra automática de texto (wrap_text = False)
        for row in ws.iter_rows():
            for cell in row:
                if cell.alignment and cell.alignment.wrap_text:
                    ws.cell(row=cell.row, column=cell.column).alignment = (
                        cell.alignment.copy(wrap_text=False)
                    )

        print(f"🧹 Aba '{sheet_name}' desmesclada e sem quebra de texto.")
        return True

    def save(self):
        # Salva as alterações no arquivo original.
        if self.is_valid:
            self.wb.save(self.file_path)
            print(f"💾 Arquivo salvo com sucesso em: '{self.file_path}'")


# ==========================================
# EXECUÇÃO DO SCRIPT
# ==========================================

if __name__ == "__main__":
    folder_path = input("O caminho da pasta:\n").strip()
    file_name = input("Nome do arquivo (ex: dados.xlsx):\n").strip()

    # Instancia e tenta carregar o arquivo
    excel = ExcelHandler(folder_path, file_name)

    if excel.is_valid:
        print("\n✅ Arquivo encontrado!")
        print(f"Abas disponíveis: {excel.get_sheet_names()}\n")

        origin_sheet = input("Aba de origem:\n").strip()
        destination_sheet = input("Aba de destino:\n").strip()

        # Executa o processo completo
        if excel.copy_sheet(origin_sheet, destination_sheet):
            excel.clean_sheet(destination_sheet)
            excel.save()