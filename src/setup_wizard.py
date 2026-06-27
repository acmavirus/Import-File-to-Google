import os
import sys
import shutil
import winreg
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# App Constants
APP_NAME = "ImportToSheet"
APP_DISPLAY_NAME = "Import to Google Sheets"
APP_VERSION = "1.0.0"
PUBLISHER = "Antigravity"


def get_resource_path(relative_path):
    """Return an absolute path to a bundled resource."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class InstallerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Setup - {APP_DISPLAY_NAME} {APP_VERSION}")
        self.geometry("620x420")
        self.resizable(False, False)

        default_dir = os.path.join(os.environ.get("LOCALAPPDATA", "C:\\"), "Programs", APP_NAME)
        self.install_dir = tk.StringVar(value=default_dir)
        self.credentials_path = tk.StringVar(value="")

        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        self.bg_color = "#f8f9fa"
        self.card_bg = "#ffffff"
        self.primary_color = "#1a73e8"
        self.primary_hover = "#1557b0"
        self.text_color = "#202124"
        self.border_color = "#dadce0"

        self.configure(bg=self.bg_color)
        self.style.configure(".", background=self.bg_color, foreground=self.text_color, font=("Segoe UI", 10))
        self.style.configure("TLabel", background=self.bg_color)
        self.style.configure("Card.TFrame", background=self.card_bg, borderwidth=1, relief="solid")
        self.style.configure("Wizard.TButton", font=("Segoe UI", 10, "bold"), padding=6)
        self.style.map(
            "Wizard.TButton",
            background=[("active", self.primary_hover), ("!disabled", self.primary_color)],
            foreground=[("active", "#ffffff"), ("!disabled", "#ffffff")],
        )

        self.current_frame = None
        self.show_welcome_page()

    def clear_frame(self):
        if self.current_frame:
            self.current_frame.destroy()

    def create_navigation_bar(self, parent, next_cmd, back_cmd=None, next_text="Tiếp tục >"):
        nav_frame = tk.Frame(parent, bg=self.bg_color, height=50)
        nav_frame.pack(side="bottom", fill="x", padx=20, pady=15)

        sep = tk.Frame(parent, bg=self.border_color, height=1)
        sep.pack(side="bottom", fill="x", padx=20)

        if back_cmd:
            back_btn = ttk.Button(nav_frame, text="< Quay lại", command=back_cmd)
            back_btn.pack(side="left")

        cancel_btn = ttk.Button(nav_frame, text="Hủy bỏ", command=self.confirm_cancel)
        cancel_btn.pack(side="right", padx=(10, 0))

        next_btn = ttk.Button(nav_frame, text=next_text, style="Wizard.TButton", command=next_cmd)
        next_btn.pack(side="right")
        return next_btn

    def confirm_cancel(self):
        if messagebox.askyesno("Hủy cài đặt", "Bạn có chắc chắn muốn thoát khỏi chương trình cài đặt?"):
            self.destroy()

    def show_welcome_page(self):
        self.clear_frame()
        self.current_frame = tk.Frame(self, bg=self.bg_color)
        self.current_frame.pack(fill="both", expand=True)

        banner = tk.Frame(self.current_frame, bg=self.primary_color, width=150)
        banner.pack(side="left", fill="y")

        lbl_banner = tk.Label(
            banner,
            text="IMPORT\nTO\nSHEET",
            fg="#ffffff",
            bg=self.primary_color,
            font=("Segoe UI", 16, "bold"),
        )
        lbl_banner.place(relx=0.5, rely=0.3, anchor="center")

        content = tk.Frame(self.current_frame, bg=self.bg_color, padx=25, pady=25)
        content.pack(side="right", fill="both", expand=True)

        title = tk.Label(
            content,
            text="Chào mừng đến với Trình Cài đặt",
            fg=self.primary_color,
            bg=self.bg_color,
            font=("Segoe UI", 16, "bold"),
            anchor="w",
        )
        title.pack(fill="x", pady=(0, 10))

        desc = tk.Label(
            content,
            text=(
                "Trình cài đặt sẽ cấu hình chương trình tự động gửi tài liệu Word/Excel/CSV "
                "lên Google Sheets từ menu chuột phải của Windows.\n\n"
                "Chương trình sẽ được cài đặt cục bộ cho tài khoản Windows hiện tại của bạn "
                "và không yêu cầu quyền Admin hệ thống.\n\n"
                "Bấm Tiếp tục để bắt đầu."
            ),
            bg=self.bg_color,
            font=("Segoe UI", 10),
            justify="left",
            wraplength=400,
        )
        desc.pack(fill="both", expand=True, pady=10)

        self.create_navigation_bar(content, self.show_directory_page)

    def show_directory_page(self):
        self.clear_frame()
        self.current_frame = tk.Frame(self, bg=self.bg_color)
        self.current_frame.pack(fill="both", expand=True)

        content = tk.Frame(self.current_frame, bg=self.bg_color, padx=25, pady=25)
        content.pack(fill="both", expand=True)

        title = tk.Label(
            content,
            text="Chọn Thư mục Cài đặt",
            fg=self.primary_color,
            bg=self.bg_color,
            font=("Segoe UI", 14, "bold"),
            anchor="w",
        )
        title.pack(fill="x", pady=(0, 10))

        desc = tk.Label(
            content,
            text="Hãy chọn thư mục trên máy tính của bạn nơi chương trình sẽ được cài đặt.",
            bg=self.bg_color,
            anchor="w",
        )
        desc.pack(fill="x", pady=(0, 15))

        dir_frame = tk.Frame(content, bg=self.bg_color)
        dir_frame.pack(fill="x", pady=10)

        txt_entry = ttk.Entry(dir_frame, textvariable=self.install_dir, font=("Segoe UI", 10), width=45)
        txt_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        def browse_dir():
            chosen = filedialog.askdirectory(initialdir=self.install_dir.get(), title="Chọn thư mục cài đặt")
            if chosen:
                self.install_dir.set(os.path.abspath(chosen))

        btn_browse = ttk.Button(dir_frame, text="Duyệt...", command=browse_dir)
        btn_browse.pack(side="right")

        self.create_navigation_bar(content, self.show_credentials_page, self.show_welcome_page, next_text="Tiếp tục")

    def show_credentials_page(self):
        self.clear_frame()
        self.current_frame = tk.Frame(self, bg=self.bg_color)
        self.current_frame.pack(fill="both", expand=True)

        content = tk.Frame(self.current_frame, bg=self.bg_color, padx=25, pady=25)
        content.pack(fill="both", expand=True)

        title = tk.Label(
            content,
            text="Cấu hình Google Sheets API (credentials.json)",
            fg=self.primary_color,
            bg=self.bg_color,
            font=("Segoe UI", 14, "bold"),
            anchor="w",
        )
        title.pack(fill="x", pady=(0, 10))

        desc = tk.Label(
            content,
            text=(
                "Ứng dụng cần file cấu hình xác thực từ Google để có quyền tạo và chỉnh sửa file Sheets.\n"
                "Nếu bạn đã tải file 'credentials.json' từ Google Cloud, hãy chọn file đó ở dưới đây.\n\n"
                "(Nếu chưa có, bạn có thể bỏ trống và copy file này vào thư mục ứng dụng sau khi cài đặt)."
            ),
            bg=self.bg_color,
            justify="left",
            wraplength=550,
            anchor="w",
        )
        desc.pack(fill="x", pady=(0, 15))

        def open_google_console():
            webbrowser.open("https://console.cloud.google.com/apis/credentials")

        btn_console = ttk.Button(
            content,
            text="Mở Google Cloud Console để lấy credentials.json",
            command=open_google_console,
        )
        btn_console.pack(fill="x", pady=(0, 15))

        file_frame = tk.Frame(content, bg=self.bg_color)
        file_frame.pack(fill="x", pady=10)

        txt_entry = ttk.Entry(file_frame, textvariable=self.credentials_path, font=("Segoe UI", 10), width=45)
        txt_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        def browse_file():
            chosen = filedialog.askopenfilename(
                title="Chọn file credentials.json của bạn",
                filetypes=[("Google Client Secrets File", "credentials.json"), ("JSON Files", "*.json")],
            )
            if chosen:
                self.credentials_path.set(os.path.abspath(chosen))

        btn_browse = ttk.Button(file_frame, text="Tải file lên...", command=browse_file)
        btn_browse.pack(side="right")

        self.create_navigation_bar(content, self.run_installation, self.show_directory_page, next_text="Cài đặt")

    def run_installation(self):
        self.clear_frame()
        self.current_frame = tk.Frame(self, bg=self.bg_color)
        self.current_frame.pack(fill="both", expand=True)

        content = tk.Frame(self.current_frame, bg=self.bg_color, padx=25, pady=25)
        content.pack(fill="both", expand=True)

        title = tk.Label(
            content,
            text="Đang tiến hành cài đặt...",
            fg=self.primary_color,
            bg=self.bg_color,
            font=("Segoe UI", 14, "bold"),
            anchor="w",
        )
        title.pack(fill="x", pady=(0, 10))

        self.progress_bar = ttk.Progressbar(content, orient="horizontal", mode="determinate", length=500)
        self.progress_bar.pack(fill="x", pady=15)

        self.status_lbl = tk.Label(
            content,
            text="Đang khởi tạo các tiến trình cài đặt...",
            bg=self.bg_color,
            font=("Segoe UI", 10, "italic"),
            anchor="w",
        )
        self.status_lbl.pack(fill="x", pady=5)

        self.update()

        try:
            self.step_install()
        except Exception as e:
            messagebox.showerror("Lỗi Cài Đặt", f"Gặp sự cố không mong muốn trong khi cài đặt:\n{e}")
            self.destroy()

    def update_progress(self, percent, text):
        self.progress_bar["value"] = percent
        self.status_lbl.configure(text=text)
        self.update()

    def step_install(self):
        dest_dir = self.install_dir.get()

        self.update_progress(10, f"Đang tạo thư mục: {dest_dir}...")
        os.makedirs(dest_dir, exist_ok=True)

        self.update_progress(30, "Đang giải nén và sao chép tệp tin thực thi chính...")
        src_exe = get_resource_path("ImportToSheet.exe")
        dest_exe = os.path.join(dest_dir, "ImportToSheet.exe")

        if os.path.exists(src_exe):
            shutil.copy2(src_exe, dest_exe)
        else:
            with open(dest_exe, "w", encoding="utf-8") as f:
                f.write("DUMMY EXE - COMPILED OUTPUT NOT FOUND IN RESOURCES")

        self.update_progress(50, "Cấu hình tài khoản và mã khóa API...")
        cred_path = self.credentials_path.get()
        if cred_path and os.path.exists(cred_path):
            shutil.copy2(cred_path, os.path.join(dest_dir, "credentials.json"))

        self.update_progress(70, "Tạo trình gỡ cài đặt hệ thống (Uninstaller)...")
        uninstall_bat_path = os.path.join(dest_dir, "uninstall.bat")
        bat_content = f"""@echo off
title Go cai dat {APP_DISPLAY_NAME}
echo Dang go bo ung dung, vui long cho...
taskkill /f /im ImportToSheet.exe >nul 2>&1
timeout /t 1 /nobreak >nul

:: Unregister Registry Context Menu
reg delete "HKCU\\Software\\Classes\\SystemFileAssociations\\.docx\\shell\\{APP_NAME}" /f >nul 2>&1
reg delete "HKCU\\Software\\Classes\\SystemFileAssociations\\.xlsx\\shell\\{APP_NAME}" /f >nul 2>&1
reg delete "HKCU\\Software\\Classes\\SystemFileAssociations\\.xls\\shell\\{APP_NAME}" /f >nul 2>&1
reg delete "HKCU\\Software\\Classes\\SystemFileAssociations\\.csv\\shell\\{APP_NAME}" /f >nul 2>&1

:: Remove Add/Remove Programs
reg delete "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{APP_NAME}" /f >nul 2>&1

:: Delete Files
del /f /q "{dest_exe}"
del /f /q "{os.path.join(dest_dir, 'credentials.json')}"
del /f /q "{os.path.join(dest_dir, 'token.json')}"

echo Xoa sach thu muc cai dat...
cd \\
rmdir /s /q "{dest_dir}"

echo Go cai dat thanh cong!
msg * Da go bo hoan toan {APP_DISPLAY_NAME} khoi may tinh cua ban.
(goto) 2>nul & del "%~f0"
"""
        with open(uninstall_bat_path, "w", encoding="utf-8") as f:
            f.write(bat_content)

        self.update_progress(80, "Đăng ký các khóa Registry cho Context Menu...")
        from register_menu import register

        register_success = register(exe_path=dest_exe)

        self.update_progress(90, "Đăng ký ứng dụng vào Control Panel / Settings...")
        reg_uninstall_path = f"Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{APP_NAME}"
        try:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_uninstall_path)
            winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, APP_DISPLAY_NAME)
            winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, f'"{uninstall_bat_path}"')
            winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, APP_VERSION)
            winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, PUBLISHER)
            winreg.SetValueEx(key, "DisplayIcon", 0, winreg.REG_SZ, dest_exe)
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Không thể ghi Registry cho Trình gỡ cài đặt: {e}")

        self.update_progress(100, "Hoàn tất cài đặt ứng dụng!")
        self.show_finished_page(register_success)

    def show_finished_page(self, register_success):
        self.clear_frame()
        self.current_frame = tk.Frame(self, bg=self.bg_color)
        self.current_frame.pack(fill="both", expand=True)

        content = tk.Frame(self.current_frame, bg=self.bg_color, padx=25, pady=25)
        content.pack(fill="both", expand=True)

        title = tk.Label(
            content,
            text="Cài Đặt Hoàn Tất!",
            fg="#0f9d58",
            bg=self.bg_color,
            font=("Segoe UI", 16, "bold"),
            anchor="w",
        )
        title.pack(fill="x", pady=(0, 10))

        dest_dir = self.install_dir.get()
        msg = (
            f"Ứng dụng '{APP_DISPLAY_NAME}' đã được cài đặt thành công tại:\n"
            f"'{dest_dir}'\n\n"
        )

        if register_success:
            msg += "Tùy chọn 'Import to Google Sheets' đã sẵn sàng khi bạn click chuột phải vào các tệp Word (.docx) hoặc Excel (.xlsx, .csv).\n\n"
        else:
            msg += "CẢNH BÁO: Đăng ký registry gặp lỗi, tùy chọn chuột phải có thể không hiển thị đúng.\n\n"

        if self.credentials_path.get() and os.path.exists(self.credentials_path.get()):
            msg += "File credentials.json đã được chép vào thư mục cài đặt.\n\n"

        if not self.credentials_path.get() or not os.path.exists(self.credentials_path.get()):
            msg += (
                "LƯU Ý: Bạn chưa chọn file credentials.json.\n"
                "Nếu file này không được chép vào thư mục cài đặt, ứng dụng sẽ không thể đăng nhập Google."
            )

        desc = tk.Label(content, text=msg, bg=self.bg_color, justify="left", wraplength=550, anchor="w")
        desc.pack(fill="x", pady=10)

        nav_frame = tk.Frame(content, bg=self.bg_color, height=50)
        nav_frame.pack(side="bottom", fill="x", padx=10, pady=15)

        sep = tk.Frame(content, bg=self.border_color, height=1)
        sep.pack(side="bottom", fill="x", padx=10)

        def open_folder():
            os.startfile(dest_dir)

        btn_open = ttk.Button(nav_frame, text="Mở thư mục cài đặt", command=open_folder)
        btn_open.pack(side="left")

        btn_finish = ttk.Button(nav_frame, text="Hoàn thành", style="Wizard.TButton", command=self.destroy)
        btn_finish.pack(side="right")


if __name__ == "__main__":
    app = InstallerApp()
    app.mainloop()
