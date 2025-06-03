# 簡易 Excel 填寫工具（自用版）

## 安裝說明（Windows）

### 1. 安裝 Python
請至以下網址下載並安裝 Python：https://www.python.org/downloads/
安裝時請務必勾選 **「Add Python to PATH」** 的選項，這樣系統才找得到 Python。

### 2. 下載這個資料夾
把壓縮檔先解壓縮。
然後打開「命令提示字元」：按下 `Win + R` → 輸入 `cmd` → 按 Enter。

### 3. 移動到資料夾
輸入這行指令，移動到剛剛解壓縮的資料夾（以下是範例路徑，請依照實際位置調整）：
```cmd
cd %USERPROFILE%\Downloads\simple-excel-filler
```

### 4. 安裝所需的工具（只需要做一次）
輸入這行指令來安裝小工具需要的套件：
```cmd
pip install -r requirements.txt
```

---

## 使用說明

### 執行小工具
按兩下 `start_app.bat` 啟動程式後，會直接打開你的瀏覽器（Chrome），並連結到這個網址 `http://127.0.0.1:5000/` 就可以看到主畫面啦！

### 功能介紹
- 預設: 自動將 `姓名、學號、電話、身分證` 等欄位填入表單中。
- 進階: 可以讓使用者自己選擇要填哪些欄位。

### 操作流程
1. 選擇 Excel 檔案 (包含要讀取的檔案以及輸出用的模板檔案)
   - 如果是預設功能，則到此步驟即可完成。
2. 選擇要取出的欄位，以及要填寫幾筆資料 (例如填到第30位工讀生)
3. 選擇對應的目標欄位，請注意格式。

## 注意事項
- 請把要讀資料的資料表放在最前面(預設是讀第一張工作表)
- 取出的欄位中，預設第一個欄位的值會成為工作表的名稱，例如: 取出的第一欄是姓名，則用姓名命名工作表
- 要填寫的數量預設從第一筆開始，到所填寫的數值為止。


------------------------------------------------------------------------------------------------------

***自己留存用***

# Simple Excel Filler

## Setup Instructions (Windows)

### 1. Install Python
Download and install from: https://www.python.org/downloads/
Make sure you check the box that says **"Add Python to PATH"** during install.

### 2. Download this folder
Unzip the folder, and open the Command Prompt (press Win + R → type `cmd` → Enter)

### 3. Navigate to the folder (example)
```cmd
cd %USERPROFILE%\Downloads\simple-excel-filler
```

### 4. Install required packages
```cmd
pip install -r requirements.txt
```

### 5. Run the app
```cmd
python app.py
```

### 6. Open in browser
Visit: `http://127.0.0.1:5000/`

---

## Usage

### Function
- Auto-fill the `姓名, 學號, 電話, 身分證` into sheets.
- Auto-fill user-chosen data into sheets.

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

