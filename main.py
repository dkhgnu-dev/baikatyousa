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
        
        # 利用可能なモデルの中から flash モデルを自動検索
        model_name = 'gemini-1.5-flash' # デフォルト
        for m in genai.list_models():
            if 'gemini-1.5-flash' in m.name:
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
        
        for file in files:
            image_bytes = await file.read()
            image = Image.open(io.BytesIO(image_bytes))
            
            response = model.generate_content([prompt, image])
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
                # If one fails, we continue with others or you can choose to abort
                print("JSON Decode Error for a file. Response was:", response_text)
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
        
        # Save to Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='月間売価調査')
        
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
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
