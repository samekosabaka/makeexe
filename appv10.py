import tkinter as tk
import hmac, base64, struct, hashlib, time, os, sys, json
import winreg

# --- CẤU HÌNH ĐƯỜNG DẪN WINAUTH ---
APP_DIR = r"C:\Users\thietbingoaiquang\Desktop\app"
CONFIG_FILE = os.path.join(APP_DIR, "config.json")

# Biến toàn cục
SECRET = ""
LOCKED = False
AUTO_START = False

def load_config():
    global SECRET, LOCKED, AUTO_START
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                SECRET = data.get("secret", "")
                LOCKED = data.get("locked", False)
                AUTO_START = data.get("auto_start", False)
        except Exception:
            pass

def save_config():
    os.makedirs(APP_DIR, exist_ok=True)
    data = {
        "secret": SECRET,
        "locked": LOCKED,
        "auto_start": AUTO_START
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f)

# Xử lý tự động khởi động bằng Windows Registry
def set_autostart(enable):
    # Xác định đường dẫn thực tế (File .exe hoặc .py)
    if getattr(sys, 'frozen', False):
        exe_path = sys.executable
    else:
        exe_path = os.path.abspath(__file__)
        
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    app_name = "DesktopOTPWidget"
    
    try:
        # Mở Registry Key
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
        if enable:
            winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, f'"{exe_path}"')
        else:
            try:
                winreg.DeleteValue(key, app_name)
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
    except Exception:
        pass

# Tải cấu hình khi vừa bật
load_config()

def get_totp(secret):
    if not secret:
        return "NO KEY"
    try:
        secret = secret.replace(' ', '').upper()
        secret += '=' * ((8 - len(secret) % 8) % 8)
        key = base64.b32decode(secret)
        msg = struct.pack(">Q", int(time.time() // 30))
        h = hmac.new(key, msg, hashlib.sha1).digest()
        o = h[19] & 15
        h = (struct.unpack(">I", h[o:o+4])[0] & 0x7fffffff) % 1000000
        return f"{h:06d}"
    except Exception:
        return "ERROR!"

def update_otp():
    otp = get_totp(SECRET)
    remaining = 30 - int(time.time() % 30)
    
    label_otp.config(text=otp)
    if otp in ["NO KEY", "ERROR!"]:
        label_time.config(text="Vui lòng nhập Key")
    else:
        label_time.config(text=f"Đổi sau: {remaining}s")
    
    # Khóa dưới cùng
    root.lower()
    root.after(1000, update_otp)

# --- XỬ LÝ NÚT KHÓA ---
def toggle_lock(event=None):
    global LOCKED
    LOCKED = not LOCKED
    # Cập nhật icon
    btn_lock.config(text="🔒" if LOCKED else "🔓", fg="#ff4444" if LOCKED else "#aaaaaa")
    save_config()

# --- XỬ LÝ CỬA SỔ CÀI ĐẶT ---
def open_settings(event=None):
    settings_win = tk.Toplevel(root)
    settings_win.title("Cài Đặt OTP")
    settings_win.geometry("300x150")
    settings_win.attributes('-topmost', True)
    
    tk.Label(settings_win, text="Nhập Secret Key 2FA của bạn:").pack(pady=(10, 0))
    entry_secret = tk.Entry(settings_win, width=35)
    entry_secret.insert(0, SECRET)
    entry_secret.pack(pady=5)
    
    # Checkbox Khởi động cùng Windows
    var_autostart = tk.BooleanVar(value=AUTO_START)
    chk_autostart = tk.Checkbutton(settings_win, text="Khởi động cùng Windows", variable=var_autostart)
    chk_autostart.pack(pady=5)
    
    def save_and_close():
        global SECRET, AUTO_START
        SECRET = entry_secret.get().strip()
        AUTO_START = var_autostart.get()
        
        save_config()
        set_autostart(AUTO_START) # Ghi vào Registry
        settings_win.destroy()
        
        # Cập nhật UI ngay lập tức
        label_otp.config(text=get_totp(SECRET))
        
    tk.Button(settings_win, text="Lưu Cấu Hình", command=save_and_close, bg="#00ff00", fg="black").pack(pady=5)

# --- XỬ LÝ KÉO THẢ (CÓ KIỂM TRA KHÓA) ---
def start_move(event):
    if LOCKED: return
    root.x = event.x
    root.y = event.y

def stop_move(event):
    if LOCKED: return
    root.x = None
    root.y = None

def do_move(event):
    if LOCKED: return
    deltax = event.x - root.x
    deltay = event.y - root.y
    x = root.winfo_x() + deltax
    y = root.winfo_y() + deltay
    root.geometry(f"+{x}+{y}")

# --- THIẾT LẬP WIDGET DESKTOP ---
root = tk.Tk()
# Mở rộng chiều ngang một chút để chứa 2 icon
root.geometry("180x65+100+100") 
root.overrideredirect(True) 

# Trong suốt nền
root.configure(bg='black')
root.wm_attributes("-transparentcolor", "black")

# Khóa cứng lớp dưới cùng
root.lower()
root.bind("<FocusIn>", lambda e: root.lower()) 

# Gán sự kiện kéo thả và nhấp đúp để thoát
root.bind("<ButtonPress-1>", start_move)
root.bind("<ButtonRelease-1>", stop_move)
root.bind("<B1-Motion>", do_move)
root.bind("<Double-Button-1>", lambda e: root.destroy())

# --- GIAO DIỆN CHỮ VÀ NÚT ---
# Nút Khóa (Thay đổi màu sắc và biểu tượng dựa trên trạng thái hiện tại)
btn_lock = tk.Label(root, text="🔒" if LOCKED else "🔓", font=("Segoe UI", 11), fg="#ff4444" if LOCKED else "#aaaaaa", bg="black", cursor="hand2")
btn_lock.place(x=135, y=2)
btn_lock.bind("<Button-1>", toggle_lock)

# Nút Cài đặt
btn_settings = tk.Label(root, text="⚙", font=("Segoe UI", 12), fg="#aaaaaa", bg="black", cursor="hand2")
btn_settings.place(x=155, y=0)
btn_settings.bind("<Button-1>", open_settings)

label_otp = tk.Label(root, text="------", font=("Consolas", 26, "bold"), fg="#00ff00", bg="black")
label_otp.pack(pady=0)

label_time = tk.Label(root, text="Đổi sau: --s", font=("Consolas", 10, "bold"), fg="#ffffff", bg="black")
label_time.pack(pady=0)

update_otp()
root.mainloop()
