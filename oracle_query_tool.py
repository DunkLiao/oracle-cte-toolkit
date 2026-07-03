import os
import threading
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import font as tkfont

import pandas as pd

from config_loader import build_config
from config_loader import build_dsn as format_dsn
from config_loader import default_config
from config_loader import get_config_path
from config_loader import load_config as read_config
from config_loader import parse_dsn as split_dsn
from config_loader import save_config as write_config
from excel_exporter import export_query_dataframes
from oracle_client import connect as connect_oracle
from sql_reader import detect_encoding as detect_file_encoding
from sql_reader import read_sql_file


DEFAULT_CONFIG_PATH = get_config_path(__file__)


class OracleQueryApp:

    def __init__(self, root):

        self.root = root
        self.root.title("Oracle 查詢工具")
        self.root.geometry("1050x820")
        self.root.minsize(950, 740)

        self.connection = None

        # 主題

        style = ttk.Style()

        try:
            style.theme_use("vista")
        except:
            style.theme_use("clam")

        style.configure(
            "TLabel",
            font=("Microsoft JhengHei UI", 10)
        )
        style.configure(
            "TButton",
            font=("Microsoft JhengHei UI", 10),
            padding=6
        )
        style.configure(
            "TEntry",
            padding=4
        )
        style.configure(
            "TLabelframe.Label",
            font=("Microsoft JhengHei UI", 11, "bold"),
            foreground="#1F4E79"
        )
        style.configure(
            "Run.TButton",
            font=("Microsoft JhengHei UI", 11, "bold"),
            foreground="white",
            background="#2E7D32",
            padding=10
        )
        style.map(
            "Run.TButton",
            background=[("active", "#1B5E20")]
        )
        style.configure(
            "Conn.TButton",
            font=("Microsoft JhengHei UI", 10, "bold"),
            foreground="white",
            background="#1565C0",
            padding=8
        )
        style.map(
            "Conn.TButton",
            background=[("active", "#0D47A1")]
        )
        style.configure(
            "Warn.TButton",
            font=("Microsoft JhengHei UI", 10, "bold"),
            foreground="white",
            background="#C62828",
            padding=8
        )
        style.map(
            "Warn.TButton",
            background=[("active", "#8E0000")]
        )
        style.configure(
            "Save.TButton",
            font=("Microsoft JhengHei UI", 10, "bold"),
            padding=8
        )

        # 變數

        self.config_file_path = tk.StringVar(
            value=DEFAULT_CONFIG_PATH
        )
        self.sql_folder_path = tk.StringVar()
        self.output_excel_path = tk.StringVar()

        self.show_password = tk.BooleanVar(value=False)

        self.build_ui()

        # 啟動時：檢查 / 建立範例 → 載入

        self.check_default_config()

        if os.path.exists(self.config_file_path.get()):
            self.load_config_from_file(silent=False)

        # 關閉時自動存檔

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.on_close
        )

    # ==============================
    # UI
    # ==============================

    def build_ui(self):

        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        main.columnconfigure(0, weight=1)
        main.rowconfigure(4, weight=1)

        # ---------- 配置檔案 ----------

        cfg_frame = ttk.LabelFrame(
            main,
            text=" 配置檔案 ",
            padding=12
        )
        cfg_frame.grid(
            row=0, column=0, sticky="ew", pady=(0, 10)
        )
        cfg_frame.columnconfigure(1, weight=1)

        ttk.Label(cfg_frame, text="Config 路徑：").grid(
            row=0, column=0, sticky="w", padx=(0, 6), pady=4
        )
        ttk.Entry(
            cfg_frame,
            textvariable=self.config_file_path
        ).grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=4)

        ttk.Button(
            cfg_frame,
            text="📁 瀏覽",
            command=self.browse_config_file
        ).grid(row=0, column=2, padx=(0, 4), pady=4)

        ttk.Button(
            cfg_frame,
            text="🔄 載入",
            command=lambda: self.load_config_from_file(silent=False)
        ).grid(row=0, column=3, padx=(0, 4), pady=4)

        ttk.Button(
            cfg_frame,
            text="💾 儲存",
            command=lambda: self.save_config(silent=False)
        ).grid(row=0, column=4, pady=4)

        # ---------- Oracle 連線 ----------

        conn_frame = ttk.LabelFrame(
            main,
            text=" Oracle 連線設定 ",
            padding=12
        )
        conn_frame.grid(
            row=1, column=0, sticky="ew", pady=(0, 10)
        )
        conn_frame.columnconfigure(1, weight=1)
        conn_frame.columnconfigure(3, weight=1)

        # Host

        ttk.Label(conn_frame, text="Host：").grid(
            row=0, column=0, sticky="w", padx=(0, 6), pady=4
        )
        self.host = ttk.Entry(conn_frame)
        self.host.grid(
            row=0, column=1, sticky="ew", padx=(0, 20), pady=4
        )

        # Port

        ttk.Label(conn_frame, text="Port：").grid(
            row=0, column=2, sticky="w", padx=(0, 6), pady=4
        )
        self.port = ttk.Entry(conn_frame, width=10)
        self.port.insert(0, "1521")
        self.port.grid(
            row=0, column=3, sticky="w", pady=4
        )

        # Service

        ttk.Label(conn_frame, text="Service Name：").grid(
            row=1, column=0, sticky="w", padx=(0, 6), pady=4
        )
        self.service = ttk.Entry(conn_frame)
        self.service.grid(
            row=1, column=1, columnspan=3,
            sticky="ew", pady=4
        )

        # User

        ttk.Label(conn_frame, text="User：").grid(
            row=2, column=0, sticky="w", padx=(0, 6), pady=4
        )
        self.user = ttk.Entry(conn_frame)
        self.user.grid(
            row=2, column=1, sticky="ew", padx=(0, 20), pady=4
        )

        # Password

        ttk.Label(conn_frame, text="Password：").grid(
            row=2, column=2, sticky="w", padx=(0, 6), pady=4
        )

        pwd_frame = ttk.Frame(conn_frame)
        pwd_frame.grid(row=2, column=3, sticky="ew", pady=4)
        pwd_frame.columnconfigure(0, weight=1)

        self.password = ttk.Entry(pwd_frame, show="●")
        self.password.grid(row=0, column=0, sticky="ew")

        ttk.Checkbutton(
            pwd_frame,
            text="顯示",
            variable=self.show_password,
            command=self.toggle_password
        ).grid(row=0, column=1, padx=(6, 0))

        # 連線狀態 + 按鈕

        conn_btn_frame = ttk.Frame(conn_frame)
        conn_btn_frame.grid(
            row=3, column=0, columnspan=4,
            sticky="ew", pady=(8, 0)
        )

        self.conn_btn = ttk.Button(
            conn_btn_frame,
            text="🔌 連接資料庫",
            style="Conn.TButton",
            command=self.connect_db
        )
        self.conn_btn.pack(side="left")

        self.close_btn = ttk.Button(
            conn_btn_frame,
            text="⏻ 關閉連接",
            style="Warn.TButton",
            command=self.close_connection
        )
        self.close_btn.pack(side="left", padx=(8, 0))

        self.conn_status_var = tk.StringVar(value="● 未連線")

        self.conn_status_label = ttk.Label(
            conn_btn_frame,
            textvariable=self.conn_status_var,
            foreground="#B71C1C",
            font=("Microsoft JhengHei UI", 10, "bold")
        )
        self.conn_status_label.pack(side="left", padx=(15, 0))

        # ---------- 路徑設定 ----------

        path_frame = ttk.LabelFrame(
            main,
            text=" 路徑設定 ",
            padding=12
        )
        path_frame.grid(
            row=2, column=0, sticky="ew", pady=(0, 10)
        )
        path_frame.columnconfigure(1, weight=1)

        ttk.Label(path_frame, text="SQL 資料夾：").grid(
            row=0, column=0, sticky="w", padx=(0, 6), pady=4
        )
        ttk.Entry(
            path_frame,
            textvariable=self.sql_folder_path
        ).grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=4)

        ttk.Button(
            path_frame,
            text="📁 瀏覽",
            command=self.browse_sql_folder
        ).grid(row=0, column=2, pady=4)

        ttk.Label(path_frame, text="Excel 輸出：").grid(
            row=1, column=0, sticky="w", padx=(0, 6), pady=4
        )
        ttk.Entry(
            path_frame,
            textvariable=self.output_excel_path
        ).grid(row=1, column=1, sticky="ew", padx=(0, 6), pady=4)

        ttk.Button(
            path_frame,
            text="💾 瀏覽",
            command=self.browse_output_file
        ).grid(row=1, column=2, pady=4)

        # ---------- 執行區 ----------

        run_frame = ttk.LabelFrame(
            main,
            text=" 執行 ",
            padding=12
        )
        run_frame.grid(
            row=3, column=0, sticky="ew", pady=(0, 10)
        )
        run_frame.columnconfigure(0, weight=1)

        btn_bar = ttk.Frame(run_frame)
        btn_bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        self.run_button = ttk.Button(
            btn_bar,
            text="▶  執行查詢",
            style="Run.TButton",
            command=self.run_thread
        )
        self.run_button.pack(side="left")

        ttk.Button(
            btn_bar,
            text="🗑 清除 Log",
            style="Save.TButton",
            command=self.clear_log
        ).pack(side="left", padx=(10, 0))

        # 進度條

        pb_frame = ttk.Frame(run_frame)
        pb_frame.grid(row=1, column=0, sticky="ew")
        pb_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(
            pb_frame,
            orient="horizontal",
            mode="determinate"
        )
        self.progress.grid(row=0, column=0, sticky="ew")

        self.progress_label = ttk.Label(
            pb_frame,
            text="0 / 0",
            width=12,
            anchor="e"
        )
        self.progress_label.grid(row=0, column=1, padx=(10, 0))

        # ---------- Log ----------

        log_frame = ttk.LabelFrame(
            main,
            text=" 執行 Log ",
            padding=8
        )
        log_frame.grid(row=4, column=0, sticky="nsew")

        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        mono = tkfont.Font(family="Consolas", size=10)

        self.log_text = tk.Text(
            log_frame,
            font=mono,
            bg="#1E1E1E",
            fg="#D4D4D4",
            insertbackground="white",
            wrap="none",
            relief="flat"
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        self.log_text.tag_config(
            "INFO", foreground="#4FC3F7"
        )
        self.log_text.tag_config(
            "OK", foreground="#81C784"
        )
        self.log_text.tag_config(
            "ERR", foreground="#E57373"
        )
        self.log_text.tag_config(
            "WARN", foreground="#FFB74D"
        )

        vsb = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.log_text.yview
        )
        vsb.grid(row=0, column=1, sticky="ns")

        hsb = ttk.Scrollbar(
            log_frame,
            orient="horizontal",
            command=self.log_text.xview
        )
        hsb.grid(row=1, column=0, sticky="ew")

        self.log_text.configure(
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set
        )

        # ---------- 狀態列 ----------

        self.status = tk.StringVar(value="就緒")

        status_bar = ttk.Frame(self.root, relief="sunken")
        status_bar.pack(fill="x", side="bottom")

        ttk.Label(
            status_bar,
            textvariable=self.status,
            anchor="w",
            padding=(10, 4)
        ).pack(fill="x")

    # ==============================
    # 工具
    # ==============================

    def log(self, msg, tag="INFO"):

        if self.log_text:
            self.log_text.insert(tk.END, msg + "\n", tag)
            self.log_text.see(tk.END)
            self.root.update_idletasks()

    def clear_log(self):
        self.log_text.delete("1.0", tk.END)

    def set_status(self, text):
        self.status.set(text)
        self.root.update_idletasks()

    def toggle_password(self):

        if self.show_password.get():
            self.password.config(show="")
        else:
            self.password.config(show="●")

    def set_entry(self, entry, value):

        entry.delete(0, tk.END)
        if value is not None:
            entry.insert(0, str(value))

    def set_conn_status(self, connected):

        if connected:
            self.conn_status_var.set("● 已連線")
            self.conn_status_label.config(foreground="#2E7D32")
        else:
            self.conn_status_var.set("● 未連線")
            self.conn_status_label.config(foreground="#B71C1C")

    # ==============================
    # DSN
    # ==============================

    def parse_dsn(self, dsn):
        return split_dsn(dsn)

    def build_dsn(self):
        return format_dsn(
            self.host.get(),
            self.port.get(),
            self.service.get(),
        )

    # ==============================
    # 編碼偵測
    # ==============================

    def detect_encoding(self, file_path):
        return detect_file_encoding(file_path)

    # ==============================
    # 檔案選擇
    # ==============================

    def browse_config_file(self):

        f = filedialog.askopenfilename(
            title="選擇配置檔案",
            filetypes=[
                ("JSON", "*.json"),
                ("所有檔案", "*.*")
            ]
        )
        if f:
            self.config_file_path.set(f)

    def browse_sql_folder(self):

        d = filedialog.askdirectory(title="選擇 SQL 資料夾")
        if d:
            self.sql_folder_path.set(d)

    def browse_output_file(self):

        f = filedialog.asksaveasfilename(
            title="Excel 輸出位置",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel", "*.xlsx"),
                ("所有檔案", "*.*")
            ]
        )
        if f:
            self.output_excel_path.set(f)

    # ==============================
    # Config
    # ==============================

    def check_default_config(self):

        path = self.config_file_path.get()

        if not os.path.exists(path):
            self.create_example_config(path)

    def create_example_config(self, config_path):

        example = default_config(config_path)

        try:

            write_config(config_path, example)

            sql_folder = example["sql_folder_path"]

            if not os.path.exists(sql_folder):

                os.makedirs(sql_folder, exist_ok=True)

                sample = os.path.join(
                    sql_folder, "example.sql"
                )
                with open(
                    sample, "w", encoding="utf-8"
                ) as f:
                    f.write(
                        "-- 範例 SQL\n"
                        "SELECT * FROM dual"
                    )

        except Exception as e:
            print(f"建立範例配置失敗: {e}")

    def load_config_from_file(self, silent=True):

        config_path = self.config_file_path.get()

        if not config_path:
            if not silent:
                messagebox.showerror(
                    "錯誤", "請先選擇配置檔案"
                )
            return None

        if not os.path.exists(config_path):
            self.log(
                f"[Config] 檔案不存在：{config_path}",
                "WARN"
            )
            return None

        try:

            enc = self.detect_encoding(config_path)

            self.log(
                f"[Config] 編碼：{enc}",
                "INFO"
            )

            cfg = read_config(config_path)

            db = cfg.get("database", {}) or {}

            username = db.get("username", "")
            password = db.get("password", "")
            dsn = db.get("dsn", "")

            host, port, service = self.parse_dsn(dsn)

            self.set_entry(self.host, host)
            self.set_entry(self.port, port or "1521")
            self.set_entry(self.service, service)
            self.set_entry(self.user, username)
            self.set_entry(self.password, password)

            self.sql_folder_path.set(
                cfg.get("sql_folder_path", "")
            )

            self.output_excel_path.set(
                cfg.get("output_excel_path", "")
            )

            self.log(
                f"[Config] 已載入：{config_path}",
                "OK"
            )

            return cfg

        except Exception as e:

            self.log(
                f"[Config] 載入失敗：{e}",
                "ERR"
            )

            if not silent:
                messagebox.showerror(
                    "錯誤",
                    f"讀取配置檔案時發生錯誤: {e}"
                )
            return None

    def save_config(self, silent=True):

        cfg = build_config(
            username=self.user.get().strip(),
            password=self.password.get(),
            dsn=self.build_dsn(),
            sql_folder_path=self.sql_folder_path.get().strip(),
            output_excel_path=self.output_excel_path.get().strip(),
        )

        path = self.config_file_path.get()

        try:

            write_config(path, cfg)

            if not silent:
                self.log(
                    f"[Config] 已儲存：{path}",
                    "OK"
                )
                messagebox.showinfo(
                    "完成",
                    f"設定已儲存\n\n{path}"
                )

        except Exception as e:

            self.log(
                f"[Config] 儲存失敗：{e}",
                "ERR"
            )
            if not silent:
                messagebox.showerror("錯誤", str(e))

    def on_close(self):

        try:
            self.save_config(silent=True)
        except:
            pass

        try:
            if self.connection:
                self.connection.close()
        except:
            pass

        self.root.destroy()

    # ==============================
    # DB 連線
    # ==============================

    def connect_db(self):

        try:

            dsn = self.build_dsn()

            self.log(
                f"[Connect] {dsn}", "INFO"
            )

            self.connection = connect_oracle(
                self.user.get().strip(),
                self.password.get(),
                dsn,
            )

            self.log(
                "[Connect] 連線成功", "OK"
            )
            self.set_conn_status(True)
            self.set_status("已連線")

            messagebox.showinfo(
                "成功",
                "已成功連接到資料庫"
            )

        except Exception as e:

            self.connection = None
            self.set_conn_status(False)

            self.log(
                f"[Connect] 失敗：{e}", "ERR"
            )

            messagebox.showerror(
                "錯誤",
                f"連接資料庫時發生錯誤: {e}"
            )

    def close_connection(self):

        if self.connection:

            try:
                self.connection.close()
            except:
                pass

            self.connection = None
            self.set_conn_status(False)
            self.set_status("連線已關閉")

            self.log(
                "[Connect] 已關閉連線", "OK"
            )

            messagebox.showinfo(
                "資訊", "資料庫連接已關閉"
            )

        else:
            self.log(
                "[Connect] 沒有開啟中的連線", "WARN"
            )

    # ==============================
    # 讀取 SQL
    # ==============================

    def read_sql_files(self):

        folder = self.sql_folder_path.get().strip()

        if not folder or not os.path.isdir(folder):
            messagebox.showerror(
                "錯誤", "請先選擇正確的 SQL 資料夾"
            )
            return None

        sql_files = {}

        for file in sorted(os.listdir(folder)):

            if not file.lower().endswith((".sql", ".txt")):
                continue

            path = os.path.join(folder, file)

            try:

                enc = self.detect_encoding(path)

                self.log(
                    f"[Read] {file} ({enc})",
                    "INFO"
                )

                content = read_sql_file(path)

                sheet = os.path.splitext(file)[0]
                sql_files[sheet] = content

            except Exception as e:

                self.log(
                    f"[Read] {file} 失敗：{e}",
                    "ERR"
                )

        if not sql_files:
            self.log(
                f"[Read] 找不到 SQL 檔",
                "WARN"
            )
            messagebox.showwarning(
                "警告",
                f"在 {folder} 中找不到 SQL 檔案"
            )

        return sql_files

    # ==============================
    # 執行
    # ==============================

    def run_thread(self):

        if not self.connection:
            messagebox.showerror(
                "錯誤", "請先連接到資料庫"
            )
            return

        if not self.output_excel_path.get().strip():
            messagebox.showerror(
                "錯誤", "請先選擇輸出 Excel 檔案路徑"
            )
            return

        # 執行前先自動存檔

        self.save_config(silent=True)

        self.run_button.config(state="disabled")

        threading.Thread(
            target=self.execute_queries,
            daemon=True
        ).start()

    def execute_queries(self):

        try:

            output_path = self.output_excel_path.get().strip()

            sql_files = self.read_sql_files()
            if not sql_files:
                return

            dataframes = {}

            total = len(sql_files)

            self.progress["value"] = 0
            self.progress["maximum"] = total
            self.progress_label.config(text=f"0 / {total}")

            for idx, (sheet_name, sql_query) in enumerate(
                sql_files.items(), start=1
            ):

                self.set_status(
                    f"執行中 ({idx}/{total}) {sheet_name}"
                )

                self.log(
                    f"[{idx}/{total}] {sheet_name}",
                    "INFO"
                )

                try:

                    df = pd.read_sql(
                        sql_query,
                        self.connection
                    )

                    dataframes[sheet_name] = df
                    valid_name = sheet_name

                    self.log(
                        f"    ✓ {len(df)} 筆 → 工作頁：{valid_name}",
                        "OK"
                    )

                except Exception as e:

                    self.log(
                        f"    ✗ 錯誤：{e}",
                        "ERR"
                    )

                self.progress["value"] = idx
                self.progress_label.config(
                    text=f"{idx} / {total}"
                )

            self.set_status("寫入 Excel...")

            export_query_dataframes(dataframes, output_path)

            self.log("", "INFO")
            self.log("=" * 60, "OK")
            self.log("✓ 完成", "OK")
            self.log(f"輸出檔：{output_path}", "OK")
            self.log("=" * 60, "OK")

            self.set_status("完成")

            messagebox.showinfo(
                "成功",
                f"已將所有查詢結果儲存到\n\n{output_path}"
            )

        except Exception as e:

            self.log(
                f"[ERROR] {e}", "ERR"
            )
            self.set_status("發生錯誤")

            messagebox.showerror("錯誤", str(e))

        finally:

            self.run_button.config(state="normal")


# ==============================
# main
# ==============================

def main():

    root = tk.Tk()
    app = OracleQueryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
