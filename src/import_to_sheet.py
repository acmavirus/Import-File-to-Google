import os
import sys
import webbrowser
import ctypes
import traceback

# Support Unicode Vietnamese characters in Windows Console
if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        pass

# Handle dependencies imports safely and show message if missing
try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from docx import Document
    import pandas as pd
except ImportError as e:
    ctypes.windll.user32.MessageBoxW(
        0,
        f"Lỗi: Thiếu thư viện Python cần thiết.\nChi tiết: {e}\n\nVui lòng cài đặt bằng cách chạy file setup hoặc lệnh 'pip install -r requirements.txt'.",
        "Lỗi Thư Viện",
        0x10
    )
    sys.exit(1)

# Google API Scopes
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]

def show_message(title, text, icon_type=0):
    # 0 = Info, 0x10 = Error, 0x30 = Warning, 0x40 = Info
    return ctypes.windll.user32.MessageBoxW(0, text, title, icon_type)

def get_credentials():
    creds = None
    if getattr(sys, 'frozen', False):
        app_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))
        
    token_path = os.path.join(app_dir, 'token.json')
    credentials_path = os.path.join(app_dir, 'credentials.json')

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                creds = None
        
        if not creds:
            if not os.path.exists(credentials_path):
                msg = (
                    "Không tìm thấy file cấu hình xác thực Google (credentials.json) trong thư mục ứng dụng.\n\n"
                    "Vui lòng:\n"
                    "1. Tạo dự án trên Google Cloud Console.\n"
                    "2. Tạo OAuth 2.0 Client ID (chọn Application type là Desktop App).\n"
                    "3. Tải file JSON cấu hình về, đổi tên thành 'credentials.json' và copy vào thư mục:\n"
                    f"'{app_dir}'\n\n"
                    "Sau đó bấm chuột phải thử lại."
                )
                show_message("Thiếu cấu hình Google Credentials", msg, 0x10)
                sys.exit(1)
                
            print("Đang mở trình duyệt để xác thực tài khoản Google...")
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            
            # Check client type: if it is a Web Client, match redirect URIs
            import json
            try:
                with open(credentials_path, 'r', encoding='utf-8') as f:
                    cred_data = json.load(f)
                if 'web' in cred_data:
                    redirect_uris = cred_data['web'].get('redirect_uris', [])
                    local_redirect = None
                    for uri in redirect_uris:
                        if 'localhost:8900' in uri:
                            local_redirect = uri
                            break
                    if local_redirect:
                        flow.redirect_uri = local_redirect
                        print("Phát hiện Web Client credentials. Sử dụng callback port 8900...")
                        creds = flow.run_local_server(port=8900)
                    else:
                        creds = flow.run_local_server(port=0)
                else:
                    creds = flow.run_local_server(port=0)
            except Exception as e:
                print(f"Lỗi khi đọc file credentials: {e}")
                creds = flow.run_local_server(port=0)
            
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
            
    return creds

def parse_docx(file_path):
    doc = Document(file_path)
    sheets_data = {}
    
    # 1. Parse tables
    has_tables = False
    for i, table in enumerate(doc.tables):
        has_tables = True
        table_rows = []
        for row in table.rows:
            row_cells = [cell.text.strip() for cell in row.cells]
            table_rows.append(row_cells)
        if table_rows:
            sheets_data[f"Table {i+1}"] = table_rows
            
    # 2. Parse paragraphs (text content)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    if paragraphs:
        # Format paragraphs as a single-column table
        paragraphs_data = [[p] for p in paragraphs]
        if has_tables:
            sheets_data["Text Content"] = paragraphs_data
        else:
            sheets_data["Sheet1"] = paragraphs_data
            
    if not sheets_data:
        sheets_data["Sheet1"] = [["File Word này không có nội dung văn bản hoặc bảng biểu."]]
        
    return sheets_data

def parse_spreadsheet(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    sheets_data = {}
    
    if ext == '.csv':
        df = pd.read_csv(file_path)
        df = df.fillna('')
        sheets_data['CSV Data'] = [df.columns.tolist()] + df.values.tolist()
    else:
        # Excel (.xlsx, .xls)
        xl = pd.ExcelFile(file_path)
        for sheet_name in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet_name)
            df = df.fillna('')
            # Limit sheet name length to 30 chars for Google Sheets compatibility
            safe_name = sheet_name[:30]
            sheets_data[safe_name] = [df.columns.tolist()] + df.values.tolist()
            
    return sheets_data

def main():
    if len(sys.argv) < 2:
        show_message("Lỗi Tham Số", "Ứng dụng được gọi mà không có đường dẫn file. Vui lòng sử dụng qua menu chuột phải của Windows.", 0x10)
        sys.exit(1)
        
    file_path = sys.argv[1]
    if not os.path.exists(file_path):
        show_message("Lỗi File", f"Không tìm thấy file: {file_path}", 0x10)
        sys.exit(1)
        
    filename = os.path.basename(file_path)
    ext = os.path.splitext(file_path)[1].lower()
    
    print(f"Đang chuẩn bị import file: {filename}...")
    
    # 1. Parse file content
    try:
        if ext == '.docx':
            sheets_data = parse_docx(file_path)
        elif ext in ['.xlsx', '.xls', '.csv']:
            sheets_data = parse_spreadsheet(file_path)
        else:
            show_message("Định Dạng Không Hỗ Trợ", f"Không hỗ trợ định dạng file: {ext}", 0x10)
            sys.exit(1)
    except Exception as e:
        show_message("Lỗi Đọc File", f"Có lỗi xảy ra khi đọc file:\n{e}", 0x10)
        traceback.print_exc()
        sys.exit(1)
        
    # 2. Authenticate Google API
    try:
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)
    except Exception as e:
        show_message("Lỗi Xác Thực", f"Không thể kết nối đến Google API:\n{e}", 0x10)
        traceback.print_exc()
        sys.exit(1)
        
    # 3. Create Google Spreadsheet
    try:
        print("Đang tạo Google Spreadsheet mới...")
        sheet_titles = list(sheets_data.keys())
        
        # Prepare body with sheets definitions
        sheets_def = [{'properties': {'title': title}} for title in sheet_titles]
        
        body = {
            'properties': {
                'title': f"[Imported] {os.path.splitext(filename)[0]}"
            },
            'sheets': sheets_def
        }
        
        spreadsheet = service.spreadsheets().create(body=body, fields='spreadsheetId,spreadsheetUrl').execute()
        spreadsheet_id = spreadsheet.get('spreadsheetId')
        spreadsheet_url = spreadsheet.get('spreadsheetUrl')
        
        # 4. Populate Spreadsheet content
        print("Đang tải dữ liệu lên Google Sheets...")
        data = []
        for sheet_name, values in sheets_data.items():
            data.append({
                'range': f"'{sheet_name}'!A1",
                'values': values
            })
            
        body_update = {
            'valueInputOption': 'USER_ENTERED',
            'data': data
        }
        
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body=body_update).execute()
        print(f"Import thành công! Link Google Sheet: {spreadsheet_url}")
        
        # Auto-open in browser
        webbrowser.open(spreadsheet_url)
        
        # Show success toast/notification
        show_message("Import Thành Công", f"Đã import thành công file '{filename}' vào Google Sheets và đang mở trên trình duyệt của bạn.", 0x40)
        
    except Exception as e:
        show_message("Lỗi Upload Google Sheets", f"Gặp lỗi khi tạo/ghi vào Google Sheets:\n{e}", 0x10)
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        show_message("Lỗi Hệ Thống", f"Đã xảy ra lỗi không mong muốn:\n{e}", 0x10)
        traceback.print_exc()
        input("\nNhấn Enter để đóng cửa sổ...")
