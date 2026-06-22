import tkinter as tk
import hmac, base64, struct, hashlib, time, os, sys, json

# --- CẤU HÌNH ĐƯỜNG DẪN WINAUTH ---
APP_DIR = r"C:\Users\thietbingoaiquang\Desktop\app"
CONFIG_FILE = os.path.join(APP_DIR, "config.json")

# --- HỆ MÀU DẠ QUANG DỄ NHÌN ---
COLORS = [
    "#00ff00", # Xanh lá (Mặc định)
    "#00ffff", # Xanh lơ (Cyan)
    "#ffff00", # Vàng
    "#ff9900", # Cam
    "#ff00ff", # Hồng (Magenta)
    "#ffffff", # Trắng
    "#00bfff", # Xanh dương nhạt
    "#ff3333"  # Đỏ tươi
]

# Biến toàn cục
SECRET = ""
LOCKED = False
AUTO_START = False
COLOR_INDEX = 0
UNLOCK_PASSWORD = "2026" 
IS_MINIMIZED = False # Trạng thái đang thu nhỏ

def load_config():
    global SECRET, LOCKED, AUTO_START, COLOR_INDEX
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                SECRET = data.get("secret", "")
                LOCKED = data.get("locked", False)
                AUTO_START = data.get("auto_start", False)
                COLOR_INDEX = data.get("color_index", 0)
        except Exception:
            pass

def save_config():
    os.makedirs(APP_DIR, exist_ok=True)
    data = {
        "secret": SECRET,
        "locked": LOCKED,
        "auto_start": AUTO_START,
        "color_index": COLOR_INDEX
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f)

def set_autostart(enable):
    if getattr(sys, 'frozen', False):
        exe_path = sys.executable
    else:
        exe_path = os.path.abspath(__file__)
        
    startup_dir = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    bat_path = os.path.join(startup_dir, "DesktopOTPWidget.bat")
    
    try:
        if enable:
            with open(bat_path, "w", encoding="utf-8") as f:
                f.write(f'@echo off\nstart "" "{exe_path}"')
        else:
            if os.path.exists(bat_path):
                os.remove(bat_path)
    except Exception:
        pass

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
        root.title("OTP: NO KEY")
    else:
        label_time.config(text=f"Đổi sau: {remaining}s")
        # [THỦ THUẬT]: Đẩy OTP và thời gian trực tiếp lên tiêu đề Taskbar
        root.title(f"{otp} ({remaining}s)")
    
    # Chỉ ép nằm dưới cùng khi đang hiển thị trên Desktop
    if not IS_MINIMIZED:
        root.lower()
        
    root.after(1000, update_otp)

# --- XỬ LÝ ẨN XUỐNG TASKBAR ---
def minimize_app(event=None):
    if LOCKED: return
    global IS_MINIMIZED
    IS_MINIMIZED = True
    # Trả lại viền cửa sổ để Windows nhận diện đưa xuống Taskbar
    root.overrideredirect(False)
    root.iconify()

def on_restore(event):
    global IS_MINIMIZED
    # Nếu đang thu nhỏ và người dùng click vào icon taskbar để mở lại
    if IS_MINIMIZED and str(event.widget) == ".":
        IS_MINIMIZED = False
        # Xóa viền và biến nền thành trong suốt trở lại
        root.overrideredirect(True)
        root.wm_attributes("-transparentcolor", "black")
        root.lower()

# --- XỬ LÝ HỆ MÀU SẮC ---
def cycle_color(event=None):
    if LOCKED: return
    global COLOR_INDEX
    COLOR_INDEX = (COLOR_INDEX + 1) % len(COLORS)
    label_otp.config(fg=COLORS[COLOR_INDEX])
    save_config()

# --- XỬ LÝ KHÓA BẢO MẬT ---
def update_lock_ui():
    btn_lock.config(text="🔒" if LOCKED else "🔓", fg="#ff4444" if LOCKED else "#aaaaaa")

def toggle_lock(event=None):
    global LOCKED
    if not LOCKED:
        LOCKED = True
        update_lock_ui()
        save_config()
    else:
        prompt_unlock()

def prompt_unlock():
    unlock_win = tk.Toplevel(root)
    unlock_win.title("Mở Khóa Widget")
    unlock_win.geometry("220x110")
    unlock_win.attributes('-topmost', True)
    
    tk.Label(unlock_win, text="Nhập mật khẩu để mở khóa:").pack(pady=5)
    
    entry_pw = tk.Entry(unlock_win, show="*", width=20, justify='center')
    entry_pw.pack(pady=5)
    entry_pw.focus()
    
    lbl_error = tk.Label(unlock_win, text="", fg="red")
    lbl_error.pack()
    
    def check_pw(event=None):
        global LOCKED
        if entry_pw.get() == UNLOCK_PASSWORD:
            LOCKED = False
            update_lock_ui()
            save_config()
            unlock_win.destroy()
        else:
            lbl_error.config(text="Sai mật khẩu!")
            entry_pw.delete(0, tk.END)
            
    tk.Button(unlock_win, text="Mở Khóa", command=check_pw, bg="#aaaaaa", fg="black").place(x=80, y=80)
    unlock_win.bind('<Return>', check_pw)

# --- XỬ LÝ CỬA SỔ CÀI ĐẶT ---
def open_settings(event=None):
    if LOCKED: return 
        
    settings_win = tk.Toplevel(root)
    settings_win.title("Cài Đặt OTP")
    settings_win.geometry("300x150")
    settings_win.attributes('-topmost', True)
    
    tk.Label(settings_win, text="Nhập Secret Key 2FA của bạn:").pack(pady=(10, 0))
    entry_secret = tk.Entry(settings_win, width=35)
    entry_secret.insert(0, SECRET)
    entry_secret.pack(pady=5)
    
    var_autostart = tk.BooleanVar(value=AUTO_START)
    chk_autostart = tk.Checkbutton(settings_win, text="Khởi động cùng Windows", variable=var_autostart)
    chk_autostart.pack(pady=5)
    
    def save_and_close():
        global SECRET, AUTO_START
        SECRET = entry_secret.get().strip()
        AUTO_START = var_autostart.get()
        
        save_config()
        set_autostart(AUTO_START)
        settings_win.destroy()
        label_otp.config(text=get_totp(SECRET))
        
    tk.Button(settings_win, text="Lưu Cấu Hình", command=save_and_close, bg="#00ff00", fg="black").pack(pady=5)

# --- XỬ LÝ KÉO THẢ & THOÁT ---
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

def exit_app(event=None):
    if not LOCKED:
        root.destroy()

# --- THIẾT LẬP WIDGET DESKTOP ---
root = tk.Tk()
# Nới rộng widget để chứa 4 nút thoải mái hơn
root.geometry("200x70+100+100") 
root.overrideredirect(True) 

root.configure(bg='black')
root.wm_attributes("-transparentcolor", "black")
root.lower()

# Lắng nghe sự kiện khôi phục từ Taskbar
root.bind("<Map>", on_restore)
root.bind("<FocusIn>", lambda e: root.lower() if not IS_MINIMIZED else None) 

root.bind("<ButtonPress-1>", start_move)
root.bind("<ButtonRelease-1>", stop_move)
root.bind("<B1-Motion>", do_move)
root.bind("<Double-Button-1>", exit_app)

# --- BỐ TRÍ CỤM NÚT BẤM (Hitbox lớn) ---
# Tăng width, padx, pady để làm ô bấm to ra

# Dòng trên: Cài đặt (⚙) và Đổi màu (🎨)
btn_settings = tk.Label(root, text="⚙", font=("Segoe UI", 12), fg="#aaaaaa", bg="black", cursor="hand2", width=2, pady=2)
btn_settings.place(x=140, y=5)
btn_settings.bind("<Button-1>", open_settings)

btn_color = tk.Label(root, text="🎨", font=("Segoe UI", 11), fg="#aaaaaa", bg="black", cursor="hand2", width=2, pady=2)
btn_color.place(x=170, y=5)
btn_color.bind("<Button-1>", cycle_color)

# Dòng dưới: Ổ khóa (🔒) và Thu nhỏ (_)
btn_lock = tk.Label(root, text="🔒" if LOCKED else "🔓", font=("Segoe UI", 11), fg="#ff4444" if LOCKED else "#aaaaaa", bg="black", cursor="hand2", width=2, pady=2)
btn_lock.place(x=140, y=35)
btn_lock.bind("<Button-1>", toggle_lock)

btn_minimize = tk.Label(root, text=" ─ ", font=("Segoe UI", 10, "bold"), fg="#aaaaaa", bg="black", cursor="hand2", width=2, pady=2)
btn_minimize.place(x=170, y=35)
btn_minimize.bind("<Button-1>", minimize_app)

# Căn lề số OTP và gán màu đã lưu
label_otp = tk.Label(root, text="------", font=("Consolas", 26, "bold"), fg=COLORS[COLOR_INDEX], bg="black")
label_otp.place(x=0, y=0)

label_time = tk.Label(root, text="Đổi sau: --s", font=("Consolas", 10, "bold"), fg="#ffffff", bg="black")
label_time.place(x=5, y=45)

update_otp()
root.mainloop()
