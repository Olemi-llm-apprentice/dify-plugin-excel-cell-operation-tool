from typing import Any, Generator
import openpyxl
from io import BytesIO

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

class ExcelCellEditorTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        try:
            # Excelファイルの取得
            excel_file = tool_parameters.get("excel_file")
            if not excel_file:
                yield ToolInvokeMessage(
                    type="text",
                    message={"text": "Excelファイルが提供されていません。"}
                )
                return

            # 操作モードの取得
            # operation = tool_parameters.get("operation", "read")
            
            # Excelファイルの読み込み
            wb = openpyxl.load_workbook(BytesIO(excel_file))
            ws = wb.active

            # if operation == "read":
            #     # セル内容の読み取り
            cell_data = {}
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value:
                        cell_data[cell.coordinate] = cell.value
            
            yield ToolInvokeMessage(
                type="text",
                message={"text": str(cell_data)}
            )
            
            # elif operation == "edit":
            #     # セル内容の編集
            #     edits = tool_parameters.get("edits", {})
            #     for cell_ref, new_value in edits.items():
            #         ws[cell_ref] = new_value
                
            #     # 編集結果を保存
            #     output = BytesIO()
            #     wb.save(output)
            #     output.seek(0)
                
            #     yield self.create_blob_message(
            #         blob=output.getvalue(),
            #         meta={
            #             "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            #             "filename": "edited_file.xlsx"
            #         }
            #     )
            
            # else:
            #     yield ToolInvokeMessage(
            #         type="text",
            #         message={"text": f"無効な操作モード: {operation}"}
            #     )

        except Exception as e:
            yield ToolInvokeMessage(
                type="text",
                message={"text": f"エラーが発生しました: {str(e)}"}
            )
