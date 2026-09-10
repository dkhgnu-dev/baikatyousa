from fastapi import FastAPI, Form, UploadFile, File, Request, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
import os
import google.generativeai as genai
from PIL import Image
import io
import json
import pandas as pd
from typing import List

app = FastAPI()
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    return templates.TemplateResponse(request=request, name="index.html")

@app.post("/api/process")
async def process_images(
    files: List[UploadFile] = File(...)
):
    try:
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            raise HTTPException(status_code=500, detail="サーバー側にAPIキーが設定されていません。管理者に連絡してください。")
            
        genai.configure(api_key=api_key)
        
        # 最新のモデル（gemini-3.5-flash）を使用
        model_name = 'gemini-3.5-flash'
        
        # 念のため利用可能か確認し、無ければ最新のflashモデルを探す
        available_models = list(genai.list_models())
        if not any(model_name in m.name for m in available_models):
            for m in available_models:
                if 'flash' in m.name and 'generateContent' in m.supported_generation_methods:
                    model_name = m.name.replace('models/', '')
                    break
                    
        model = genai.GenerativeModel(model_name)
        
        prompt = """
        この画像は売価調査票です。画像内の表データを抽出し、以下の列定義に従ってJSONの配列形式で出力してください。
        抽出する列:
        ["NO", "商品コード", "商品名称", "規格", "自社売価", "競合店A", "競合店B"]
        
        注意事項:
        - 該当データがないセルは空文字("")にしてください。
        - JSON配列のみを出力してください。Markdownのコードブロック(```json ... ```)は付けないでください。
        
        出力形式の例:
        [
            {
                "NO": "1",
                "商品コード": "123456",
                "商品名称": "テスト商品",
                "規格": "100g",
                "自社売価": "100",
                "競合店A": "98",
                "競合店B": "105"
            }
        ]
        """
        
        all_data = []
        import time
        
        for file in files:
            image_bytes = await file.read()
            image = Image.open(io.BytesIO(image_bytes))
            
            # 429エラー（リクエスト過多）対策のためのリトライ処理
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    response = model.generate_content([prompt, image])
                    break # 成功したらループを抜ける
                except Exception as e:
                    if "429" in str(e) and attempt < max_retries - 1:
                        time.sleep(8) # 制限に引っかかったら8秒待ってから再試行
                    else:
                        raise e
            
            # 短時間での送りすぎを防ぐため、1枚終わるごとに少し待機
            time.sleep(1.5)
            
            response_text = response.text.strip()
            
            # Clean up markdown if model returned it
            if response_text.startswith("```json"):
                response_text = response_text[7:]
            if response_text.startswith("```"):
                response_text = response_text[3:]
            if response_text.endswith("```"):
                response_text = response_text[:-3]
                
            try:
                data = json.loads(response_text.strip())
                if isinstance(data, list):
                    all_data.extend(data)
                elif isinstance(data, dict):
                    all_data.append(data)
            except json.JSONDecodeError:
                pass

        if not all_data:
            raise HTTPException(status_code=400, detail="データを抽出できませんでした。")

        # Create DataFrame
        df = pd.DataFrame(all_data)
        expected_columns = ["NO", "商品コード", "商品名称", "規格", "自社売価", "競合店A", "競合店B"]
        for col in expected_columns:
            if col not in df.columns:
                df[col] = ""
        df = df[expected_columns]
        
        # Save to Excel with Formatting
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='月間売価調査')
            
            workbook = writer.book
            worksheet = writer.sheets['月間売価調査']
            
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            
            # 書式の定義
            header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") # 薄い青色
            header_font = Font(bold=True)
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                               top=Side(style='thin'), bottom=Side(style='thin'))
            center_alignment = Alignment(horizontal="center", vertical="center")
            left_alignment = Alignment(horizontal="left", vertical="center")
            right_alignment = Alignment(horizontal="right", vertical="center")
            
            # 列幅の設定
            column_widths = {
                "A": 5,   # NO
                "B": 15,  # 商品コード
                "C": 35,  # 商品名称
                "D": 15,  # 規格
                "E": 12,  # 自社売価
                "F": 12,  # 競合店A
                "G": 12   # 競合店B
            }
            for col, width in column_widths.items():
                worksheet.column_dimensions[col].width = width

            # 1行目（見出し）の固定
            worksheet.freeze_panes = "A2"

            # 全セルの書式設定
            for row_idx, row in enumerate(worksheet.iter_rows(min_row=1, max_row=len(df)+1, min_col=1, max_col=7), 1):
                for col_idx, cell in enumerate(row, 1):
                    # 罫線を引く
                    cell.border = thin_border
                    
                    if row_idx == 1:
                        # 見出し行の装飾
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = center_alignment
                    else:
                        # データ行の装飾
                        if col_idx in [1, 2]: # NO, 商品コード
                            cell.alignment = center_alignment
                        elif col_idx in [3, 4]: # 商品名称, 規格
                            cell.alignment = left_alignment
                        elif col_idx in [5, 6, 7]: # 売価の列
                            cell.alignment = right_alignment
                            # 数字の場合はカンマ区切りにする
                            try:
                                if str(cell.value).replace(',', '').isdigit():
                                    cell.value = int(str(cell.value).replace(',', ''))
                                    cell.number_format = '#,##0'
                            except:
                                pass
        
        output.seek(0)
        
        headers = {
            'Content-Disposition': 'attachment; filename="survey_result.xlsx"'
        }
        return StreamingResponse(
            output, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers
        )

    except Exception as e:
        available_str = ""
        try:
            available = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            if available:
                available_str = f" | 利用可能なモデル: {', '.join(available)}"
        except:
            pass
        raise HTTPException(status_code=500, detail=f"APIエラー: {str(e)}{available_str}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
