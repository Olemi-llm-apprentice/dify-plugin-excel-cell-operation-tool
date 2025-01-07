from typing import Any, Generator
import openpyxl
import requests
from io import BytesIO
from urllib.parse import urlparse

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

def get_url_from_file_data(file_data: Any) -> str:
    """DifyのファイルデータからURLを抽出する"""
    if hasattr(file_data, 'url'):
        return file_data.url
    elif isinstance(file_data, dict) and 'url' in file_data:
        return file_data['url']
    return ''

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
            
            # ファイルのURLを取得
            file_url = get_url_from_file_data(excel_file)
            
            if not file_url:
                yield ToolInvokeMessage(
                    type="text",
                    message={"text": "ファイルのURLが見つかりません。"}
                )
                return

            # ファイルをダウンロード
            try:
                response = requests.get(file_url)
                response.raise_for_status()
                file_bytes = response.content
                
                # デバッグ用ログ（コンソール出力）
                print(f"ファイルURL: {file_url}")
                print(f"ファイルサイズ: {len(file_bytes)} bytes")
            except Exception as e:
                yield ToolInvokeMessage(
                    type="text",
                    message={"text": f"ファイルのダウンロードに失敗しました: {str(e)}"}
                )
                return

            # Excelファイルの読み込み
            try:
                wb = openpyxl.load_workbook(BytesIO(file_bytes))
                ws = wb.active
            except Exception as e:
                yield ToolInvokeMessage(
                    type="text", 
                    message={"text": f"Excelファイル読み込みエラー: {str(e)}"}
                )
                return

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
