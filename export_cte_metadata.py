import os
import threading
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import font as tkfont

from config_loader import build_config
from config_loader import build_dsn as format_dsn
from config_loader import get_config_path as shared_get_config_path
from config_loader import load_config as read_config
from config_loader import parse_dsn as split_dsn
from config_loader import save_config as write_config
from excel_exporter import export_metadata_rows
from oracle_client import connect as connect_oracle
from oracle_client import oracle_type_from_description
from oracle_client import wrap_metadata_sql
from sql_reader import read_sql_file


CONFIG_PATH = shared_get_config_path(__file__)


class MetadataExporter:

    def __init__(self, root):

        self.root = root
        self.root.title("Oracle CTE Metadata Exporter")
        self.root.geometry("1000x800")
        self.root.minsize(900, 720)

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
            "Save.TButton",
            font=("Microsoft JhengHei UI", 10, "bold"),
            padding=8
        )

        self.show_password = tk.BooleanVar(value=False)

        self.build_ui()

        # 啟動時載入設定

        self.load_config()

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
        main.rowconfigure(3, weight=1)

        # ===== Oracle 連線設定 =====

        conn_frame = ttk.LabelFrame(
            main,
            text=" Oracle 連線設定 ",
            padding=12
        )
        conn_frame.grid(
            row=0,
            column=0,
            sticky="ew",
            pady=(0, 10)
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

        # Service Name

        ttk.Label(conn_frame, text="Service Name：").grid(
            row=1, column=0, sticky="w", padx=(0, 6), pady=4
        )
        self.service = ttk.Entry(conn_frame)
        self.service.grid(
            row=1, column=1, columnspan=3, sticky="ew", pady=4
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

        # ===== 路徑設定 =====

        path_frame = ttk.LabelFrame(
            main,
            text=" 路徑設定 ",
            padding=12
        )
        path_frame.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(0, 10)
        )

        path_frame.columnconfigure(1, weight=1)

        # SQL Folder

        ttk.Label(path_frame, text="SQL 資料夾：").grid(
            row=0, column=0, sticky="w", padx=(0, 6), pady=4
        )
        self.sql_folder = ttk.Entry(path_frame)
        self.sql_folder.grid(
            row=0, column=1, sticky="ew", padx=(0, 6), pady=4
        )
        ttk.Button(
            path_frame,
            text="📁 瀏覽",
            command=self.select_sql_folder
        ).grid(row=0, column=2, pady=4)

        # Output

        ttk.Label(path_frame, text="Excel 輸出：").grid(
            row=1, column=0, sticky="w", padx=(0, 6), pady=4
        )
        self.output_file = ttk.Entry(path_frame)
        self.output_file.grid(
            row=1, column=1, sticky="ew", padx=(0, 6), pady=4
        )
        ttk.Button(
            path_frame,
            text="💾 瀏覽",
            command=self.select_output_file
        ).grid(row=1, column=2, pady=4)

        # ===== 執行區 =====

        run_frame = ttk.LabelFrame(
            main,
            text=" 執行 ",
            padding=12
        )
        run_frame.grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(0, 10)
        )

        run_frame.columnconfigure(0, weight=1)

        btn_bar = ttk.Frame(run_frame)
        btn_bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        self.run_button = ttk.Button(
            btn_bar,
            text="▶  開始分析",
            style="Run.TButton",
            command=self.run_thread
        )
        self.run_button.pack(side="left")

        ttk.Button(
            btn_bar,
            text="💾 儲存設定",
            style="Save.TButton",
            command=self.save_config_manual
        ).pack(side="left", padx=(10, 0))

        ttk.Button(
            btn_bar,
            text="🔄 重新載入設定",
            style="Save.TButton",
            command=self.load_config
        ).pack(side="left", padx=(10, 0))

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

        # ===== Log =====

        log_frame = ttk.LabelFrame(
            main,
            text=" 執行 Log ",
            padding=8
        )
        log_frame.grid(
            row=3,
            column=0,
            sticky="nsew"
        )

        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        mono = tkfont.Font(family="Consolas", size=10)

        self.log = tk.Text(
            log_frame,
            font=mono,
            bg="#1E1E1E",
            fg="#D4D4D4",
            insertbackground="white",
            wrap="none",
            relief="flat"
        )
        self.log.grid(row=0, column=0, sticky="nsew")

        self.log.tag_config("INFO", foreground="#4FC3F7")
        self.log.tag_config("OK", foreground="#81C784")
        self.log.tag_config("ERR", foreground="#E57373")
        self.log.tag_config("WARN", foreground="#FFB74D")

        vsb = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.log.yview
        )
        vsb.grid(row=0, column=1, sticky="ns")

        hsb = ttk.Scrollbar(
            log_frame,
            orient="horizontal",
            command=self.log.xview
        )
        hsb.grid(row=1, column=0, sticky="ew")

        self.log.configure(
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set
        )

        # ===== 狀態列 =====

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
    # Config 讀寫
    # ==============================

    def parse_dsn(self, dsn):
        """
        將 host:port/service 拆解
        """

        return split_dsn(dsn)

    def build_dsn(self):

        return format_dsn(
            self.host.get(),
            self.port.get(),
            self.service.get(),
        )

    def set_entry(self, entry, value):

        entry.delete(0, tk.END)

        if value is not None:
            entry.insert(0, str(value))

    def load_config(self):

        if not os.path.exists(CONFIG_PATH):
            self.write_log(
                f"[Config] 未找到設定檔：{CONFIG_PATH}",
                "WARN"
            )
            return

        try:

            cfg = read_config(CONFIG_PATH)

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

            self.set_entry(
                self.sql_folder,
                cfg.get("sql_folder_path", "")
            )

            self.set_entry(
                self.output_file,
                cfg.get("output_excel_path", "")
            )

            self.write_log(
                f"[Config] 已載入：{CONFIG_PATH}",
                "OK"
            )

        except Exception as e:

            self.write_log(
                f"[Config] 載入失敗：{e}",
                "ERR"
            )

    def save_config(self, silent=True):

        cfg = build_config(
            username=self.user.get().strip(),
            password=self.password.get(),
            dsn=self.build_dsn(),
            sql_folder_path=self.sql_folder.get().strip(),
            output_excel_path=self.output_file.get().strip(),
        )

        try:

            write_config(CONFIG_PATH, cfg)

            if not silent:

                self.write_log(
                    f"[Config] 已儲存：{CONFIG_PATH}",
                    "OK"
                )

                messagebox.showinfo(
                    "完成",
                    f"設定已儲存\n\n{CONFIG_PATH}"
                )

        except Exception as e:

            self.write_log(
                f"[Config] 儲存失敗：{e}",
                "ERR"
            )

            if not silent:
                messagebox.showerror("錯誤", str(e))

    def save_config_manual(self):
        self.save_config(silent=False)

    def on_close(self):

        try:
            self.save_config(silent=True)
        except:
            pass

        self.root.destroy()

    # ==============================
    # 密碼顯示
    # ==============================

    def toggle_password(self):

        if self.show_password.get():
            self.password.config(show="")
        else:
            self.password.config(show="●")

    # ==============================
    # 檔案
    # ==============================

    def select_sql_folder(self):

        folder = filedialog.askdirectory(title="選擇 SQL 資料夾")

        if folder:
            self.sql_folder.delete(0, tk.END)
            self.sql_folder.insert(0, folder)

    def select_output_file(self):

        file = filedialog.asksaveasfilename(
            title="Excel 輸出位置",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )

        if file:
            self.output_file.delete(0, tk.END)
            self.output_file.insert(0, file)

    # ==============================
    # Log
    # ==============================

    def write_log(self, text, tag="INFO"):

        self.log.insert(tk.END, text + "\n", tag)
        self.log.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):

        self.log.delete("1.0", tk.END)

    def set_status(self, text):

        self.status.set(text)
        self.root.update_idletasks()

    # ==============================
    # Oracle 型態
    # ==============================

    def oracle_type(self, col):
        return oracle_type_from_description(col)

    # ==============================
    # SQL 包裝
    # ==============================

    def wrap_sql(self, sql_text):
        return wrap_metadata_sql(sql_text)

    # ==============================
    # 執行
    # ==============================

    def run_thread(self):

        if not self.host.get().strip():
            messagebox.showwarning("提示", "請輸入 Host")
            return

        if not self.sql_folder.get().strip():
            messagebox.showwarning("提示", "請選擇 SQL 資料夾")
            return

        if not self.output_file.get().strip():
            messagebox.showwarning("提示", "請選擇 Excel 輸出檔")
            return

        # 執行前自動存檔

        self.save_config(silent=True)

        self.run_button.config(state="disabled")

        threading.Thread(
            target=self.execute,
            daemon=True
        ).start()

    def execute(self):

        try:

            dsn = self.build_dsn()

            self.set_status("連線中...")
            self.write_log(f"[Connect] {dsn}", "INFO")

            conn = connect_oracle(
                self.user.get().strip(),
                self.password.get(),
                dsn,
            )

            cur = conn.cursor()

            self.write_log("[Connect] 連線成功", "OK")

            sql_folder = self.sql_folder.get().strip()
            output_file = self.output_file.get().strip()

            sql_files = [
                x for x in os.listdir(sql_folder)
                if x.lower().endswith(".sql")
            ]

            total = len(sql_files)

            self.progress["value"] = 0
            self.progress["maximum"] = total
            self.progress_label.config(text=f"0 / {total}")

            self.write_log(
                f"[Scan] 找到 {total} 個 SQL 檔",
                "INFO"
            )

            all_rows = []

            for idx, file_name in enumerate(sql_files, start=1):

                self.set_status(
                    f"處理中 ({idx}/{total}) {file_name}"
                )

                self.write_log(
                    f"[{idx}/{total}] {file_name}",
                    "INFO"
                )

                path = os.path.join(sql_folder, file_name)

                try:
                    sql_text = read_sql_file(path)
                except Exception as ex:
                    self.write_log(
                        f"    無法讀取 SQL 檔案: {ex}",
                        "ERR"
                    )
                    continue

                if not sql_text:
                    self.write_log(
                        f"    ✗ 無法讀取檔案",
                        "ERR"
                    )
                    continue

                try:

                    sql = self.wrap_sql(sql_text)
                    cur.execute(sql)

                    for seq, col in enumerate(
                        cur.description, start=1
                    ):

                        all_rows.append({
                            "SQL_FILE": file_name,
                            "COLUMN_ID": seq,
                            "COLUMN_NAME": col[0],
                            "ORACLE_TYPE": self.oracle_type(col),
                            "LENGTH": col[3],
                            "PRECISION": col[4],
                            "SCALE": col[5],
                            "NULLABLE": col[6]
                        })

                    self.write_log(
                        f"    ✓ {len(cur.description)} 個欄位",
                        "OK"
                    )

                except Exception as ex:

                    self.write_log(
                        f"    ✗ 錯誤: {ex}",
                        "ERR"
                    )

                    all_rows.append({
                        "SQL_FILE": file_name,
                        "COLUMN_ID": "",
                        "COLUMN_NAME": "",
                        "ORACLE_TYPE": "",
                        "LENGTH": "",
                        "PRECISION": "",
                        "SCALE": "",
                        "NULLABLE": "",
                        "ERROR": str(ex)
                    })

                self.progress["value"] = idx
                self.progress_label.config(
                    text=f"{idx} / {total}"
                )

            self.set_status("寫入 Excel...")

            export_metadata_rows(all_rows, output_file)

            cur.close()
            conn.close()

            self.write_log("", "INFO")
            self.write_log("=" * 60, "OK")
            self.write_log("✓ 完成", "OK")
            self.write_log(f"輸出檔: {output_file}", "OK")
            self.write_log("=" * 60, "OK")

            self.set_status("完成")

            messagebox.showinfo(
                "完成",
                f"Metadata 已成功匯出\n\n{output_file}"
            )

        except Exception as e:

            self.write_log(f"[ERROR] {e}", "ERR")
            self.set_status("發生錯誤")

            messagebox.showerror("錯誤", str(e))

        finally:

            self.run_button.config(state="normal")


if __name__ == "__main__":

    root = tk.Tk()
    app = MetadataExporter(root)
    root.mainloop()
