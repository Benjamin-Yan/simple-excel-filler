# 簡易 Excel 填寫工具（自用版）

## 懶人包
> 1. 安裝 Python
> 2. 下載程式壓縮檔
> 3. 參考這部 [demo 影片](https://youtu.be/Z9ull5nU5Ok)的說明
> 4. Done!

## 安裝說明（Windows, 只需要做一次）

> [!NOTE]
> 程式已部屬在 [render](https://simple-excel-filler.onrender.com) 網站，可以移步至該處而不用進行安裝。

1. 下載原始碼
   到 https://github.com/Benjamin-Yan/simple-excel-filler/releases/tag/V1.0.1 下載 `simple-excel-filler.zip` 並解壓縮到任意資料夾中。
2. 安裝 Python
   如果電腦沒有下載 python，請至以下網址下載並安裝 Python：https://www.python.org/downloads/
   安裝時請務必勾選 **「Add Python to PATH」** 的選項，這樣系統才找得到 Python。

   參考影片連結: https://youtu.be/czeEyrCm-bQ (看下載python就好)
3. 安裝所需的工具
   點兩下 `create_venv.bat`，等待它創建虛擬環境並安裝需要的 python 套件。

## 使用說明

> [!NOTE]
> - 目前限定使用東華IP使用本程式。
> - 除非你想要下載原始碼並設定環境變數，否則可以跳過 `執行小工具` 這一part.

### 執行小工具
按兩下 `start_app.bat`，等待網頁跳出即可。(預設使用 Chrome 開啟，也可手動連結到這個網址 `http://127.0.0.1:5000/`)

### 功能介紹
- 預設: 自動將 `姓名、學號、電話、身分證` 等欄位填入表單中。
- 進階: 可以自己選擇要填的欄位。

### 操作流程

可以參考這部 [demo 影片](https://youtu.be/Z9ull5nU5Ok) 的示範。

1. 選擇 Excel 檔案 (包含要讀取的檔案以及輸出用的模板檔案)
   - 如果是預設功能，則到此步驟即可完成。
2. 選擇要取出的欄位，以及要填寫幾筆資料 (例如填到第30位工讀生)
3. 選擇對應的目標欄位，請注意格式，如: `3F`。
4. 結束填寫時，把命令提示符(黑底的視窗)關掉即可。

## 注意事項
- 請把要讀資料的資料表放在最前面(預設是讀第一張工作表)
- 取出的欄位中，預設第一個欄位的值會成為工作表的名稱，例如: 取出的第一欄是姓名，則用姓名命名工作表
- 要填寫的數量預設從第一筆開始，到所填寫的數值為止。

---

## Folder Structure
```
simple-excel-filler/
│
├── app.py                     # Main Flask app
├── shared.py
├── utils/
│   ├── __init__.py
│   ├── set_fixed_cell_value.py
│   └── set_selected_value.py
│
├── templates/
│   ├── index.html
│   └── select.html
│
├── static/
│   ├── css/
│   │   └── style.css
│   └── js/
│       └── select.js
│
├── requirements.txt
├── start_app.bat              # Batch file to run the app
└── README.md
```

