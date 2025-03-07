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

def get_blob_from_file_data(file_data: Any) -> bytes:
    """Difyのファイルデータからblobを抽出する"""
    if hasattr(file_data, 'blob'):
        return file_data.blob
    elif isinstance(file_data, dict) and 'blob' in file_data:
        return file_data['blob']
    return None

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

            # ファイルのblobデータを取得
            file_blob = get_blob_from_file_data(excel_file)
            
            if not file_blob:
                yield ToolInvokeMessage(
                    type="text",
                    message={"text": "ファイルのblobデータが見つかりません。"}
                )
                return

            # Excelファイルの読み込み
            try:
                wb = openpyxl.load_workbook(BytesIO(file_blob))
                ws = wb.active
            except Exception as e:
                yield ToolInvokeMessage(
                    type="text", 
                    message={"text": f"Excelファイル読み込みエラー: {str(e)}"}
                )
                return

            # セル内容の読み取り
            cell_data = {}
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value:
                        cell_data[cell.coordinate] = cell.value
            
            yield ToolInvokeMessage(
                type="text",
                message={"text": str(cell_data)}
            )

        except Exception as e:
            yield ToolInvokeMessage(
                type="text",
                message={"text": f"エラーが発生しました: {str(e)}"}
            )
