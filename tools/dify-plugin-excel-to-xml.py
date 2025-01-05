from collections.abc import Generator
from typing import Any
import pandas as pd
from io import BytesIO
import json
import requests
from urllib.parse import urlparse

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

class DifyPluginExcelToXMLTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        # ファイルURLを取得
        file_url = tool_parameters.get("file_url", "")
        if not file_url:
            yield ToolInvokeMessage(
                type="text",
                message={"text": "ファイルURLが提供されていません。"}
            )
            return

        try:
            # URLの検証
            parsed_url = urlparse(file_url)
            if not parsed_url.scheme or not parsed_url.netloc:
                raise ValueError("無効なURLです")

            # ファイルをダウンロード
            response = requests.get(file_url)
            response.raise_for_status()
            
            # ExcelファイルをPandasで読み込む
            excel_data = BytesIO(response.content)
            df = pd.read_excel(excel_data)
            
            # XMLに変換
            xml_data = df.to_xml(index=False, root_name="data", row_name="row")
            
            # XML形式のテキストを返す
            yield ToolInvokeMessage(
                type="text",
                message={"text": xml_data}
            )

        except requests.exceptions.RequestException as e:
            yield ToolInvokeMessage(
                type="text",
                message={"text": f"ファイルのダウンロードに失敗しました: {str(e)}"}
            )
        except Exception as e:
            yield ToolInvokeMessage(
                type="text",
                message={"text": f"エラーが発生しました: {str(e)}"}
            )
            return 