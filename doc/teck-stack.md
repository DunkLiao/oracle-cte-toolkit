# 技術棧與套件說明

本專案是一組以 Python 開發的 Windows 桌面工具，用於連線 Oracle、執行 SQL/CTE 查詢、讀取欄位 metadata，並將結果匯出為 Excel。若未來要模組化產生同類工具，可把「設定載入、GUI、資料庫連線、SQL 處理、Excel 匯出」拆成可重用元件。

## 核心技術棧

| 類別 | 技術 | 用途 |
| --- | --- | --- |
| 程式語言 | Python 3.9+ | 主程式、GUI、資料處理與 Oracle 存取 |
| 桌面介面 | Tkinter / ttk | 建立表單、按鈕、檔案選擇、訊息視窗與執行狀態 |
| 資料庫 | Oracle Database | 查詢來源資料、解析 CTE 欄位 metadata |
| 設定格式 | JSON | 使用 `config.json` 保存本機連線與輸出路徑 |
| 輸出格式 | Excel `.xlsx` | 交付查詢結果與欄位說明 |

## Python 標準庫

- `os` / `sys`: 判斷執行檔位置、組合 `config.json` 路徑，並支援 PyInstaller 類型的 frozen 執行環境。
- `json`: 讀寫 `config.json`，保存 Oracle 帳號、DSN、SQL 資料夾與輸出路徑。
- `threading`: 將耗時查詢放到背景執行，避免 Tkinter GUI 凍結。
- `tkinter`: 提供桌面操作介面，包含 `ttk`、`filedialog`、`messagebox`、`font`。

## 第三方套件

- `oracledb`: 連線 Oracle。預設可使用 Thin Mode；若目標資料庫版本較舊，可能需要安裝 Oracle Instant Client 並改用 Thick Mode。
- `pandas`: 將查詢結果整理成 DataFrame，方便後續轉換與輸出。
- `openpyxl`: 建立與寫入 `.xlsx` 檔案；`oracle_query_tool.py` 也使用 `dataframe_to_rows` 將 DataFrame 寫入工作表。
- `chardet`: 偵測 `.sql` 檔案編碼，支援 UTF-8、BIG5、CP950 等常見來源。

安裝指令：

```powershell
pip install oracledb pandas openpyxl chardet
```

## 現有工具模組

- `oracle_query_tool.py`: 批次讀取 SQL 檔、偵測編碼、連線 Oracle、執行查詢並匯出 Excel。
- `export_cte_metadata.py`: 透過 `SELECT * FROM (...) WHERE 1=0` 取得 `cursor.description`，匯出 CTE 欄位 metadata。
- `config.example.json`: 可複製為 `config.json`，作為本機工具設定入口。

## 建議的模組化方向

未來產生同類工具時，可優先抽出以下模組：

- `config_loader.py`: 統一處理設定檔路徑、預設值、讀寫與敏感資訊檢查。
- `oracle_client.py`: 封裝 `oracledb.connect`、連線測試、錯誤轉換與 Thin/Thick Mode 切換。
- `sql_reader.py`: 封裝 SQL 檔案搜尋、編碼偵測、內容讀取與基本驗證。
- `excel_exporter.py`: 統一 DataFrame、metadata 與多工作表 Excel 輸出格式。
- `tkinter_shell.py`: 提供共用 GUI layout、欄位輸入、背景執行、訊息提示與進度狀態。

## 開發與驗證

基本語法檢查：

```powershell
python -m py_compile oracle_query_tool.py export_cte_metadata.py
```

本機執行前，先由範本建立設定檔：

```powershell
Copy-Item config.example.json config.json
```

`config.json` 可能包含帳密與內部主機資訊，必須維持在 `.gitignore` 中，不可提交到版本庫。
