from collections.abc import Generator
from typing import Any
import pandas as pd
import tempfile
import os
from pathlib import Path
from io import StringIO
import json

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

class DifyPluginMarkdownToExcelTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        # マークダウンテキストを取得
        markdown_text = tool_parameters.get("markdown_text", "")
        if not markdown_text:
            yield ToolInvokeMessage(
                type="text",
                message={"text": "マークダウンテキストが提供されていません。"}
            )
            return

        try:
            # JSONとして解析可能な場合は、"text"フィールドを取得
            try:
                json_data = json.loads(markdown_text)
                if isinstance(json_data, dict) and "text" in json_data:
                    markdown_text = json_data["text"]
            except json.JSONDecodeError:
                pass

            # マークダウンの表を行ごとに分割
            lines = markdown_text.strip().split('\n')
            if len(lines) < 2:
                yield ToolInvokeMessage(
                    type="text",
                    message={"text": "有効なマークダウン表が見つかりません。"}
                )
                return

            # ヘッダー行とデータ行を分離
            header_row = [col.strip() for col in lines[0].split('|') if col.strip()]
            data_rows = []
            
            for line in lines[2:]:  # 区切り行をスキップ
                if line.strip():
                    cols = [col.strip() for col in line.split('|') if col]
                    if cols:
                        data_rows.append(cols)

            # DataFrameを作成
            df = pd.DataFrame(data_rows, columns=header_row)
            
            # 空の列を削除
            df = df.replace('nan', '')
            df = df.replace('', pd.NA)
            df = df.dropna(axis=1, how='all')
            df = df.fillna('')

            # 一時ファイルを作成
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                excel_path = tmp_file.name
                
            # ExcelファイルとしてDataFrameを保存
            df.to_excel(excel_path, index=False, engine='openpyxl')
            
            # ファイルのURLを返す
            file_url = f"file://{excel_path}"
            yield ToolInvokeMessage(
                type="text",
                message={"text": file_url}
            )

        except Exception as e:
            yield ToolInvokeMessage(
                type="text",
                message={"text": f"エラーが発生しました: {str(e)}"}
            )
            return
