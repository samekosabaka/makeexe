import tkinter as tk
import hmac, base64, struct, hashlib, time, os, sys, json, threading
from PIL import Image, ImageDraw
import pystray

# --- CẤU HÌNH ĐƯỜNG DẪN ---
APP_DIR = r"C:\Users\thietbingoaiquang\Desktop\app"
CONFIG_FILE = os.path.join(APP_DIR, "config.json")

# --- HỆ MÀU DẠ QUANG DỄ NHÌN ---
COLORS = ["#00ff00", "#00ffff", "#ffff00", "#ff9900", "#ff00ff", "#ffffff", "#00bfff", "#ff3333"]

# --- FONT CHỮ MỚI ---
OTP_FONT = ("Segoe UI", 28, "bold")

# Biến toàn cục
SECRET = ""
LOCKED = False
AUTO_START = False
COLOR_INDEX = 0
UNLOCK_PASSWORD = "2026" 
IS_MINIMIZED = False
tray_icon = None
last_remaining = -1

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
    data = {"secret": SECRET, "locked": LOCKED, "auto_start": AUTO_START, "color_index": COLOR_INDEX}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f)

def set_autostart(enable):
    exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
    startup_dir = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    bat_path = os.path.join(startup_dir, "DesktopOTPWidget.bat")
    try:
        if enable:
            with open(bat_path, "w", encoding="utf-8") as f:
                f.write(f'@echo off\nstart "" "{exe_path}"')
        elif os.path.exists(bat_path):
            os.remove(bat_path)
    except Exception:
        pass

load_config()

# --- VÒNG LẶP GIAO DIỆN CHÍNH (Chạy mỗi 100ms để Animation mượt) ---
def update_ui_loop():
    global last_remaining, tray_icon
    
    if not SECRET:
        label_otp.config(text="NO KEY", font=("Segoe UI", 20, "bold"))
        progress_canvas.itemconfig(progress_bar, fill="black")
        root.after(100, update_ui_loop)
        return

    # 1. Tính toán OTP
    try:
        secret_cl = SECRET.replace(' ', '').upper()
        secret_cl += '=' * ((8 - len(secret_cl) % 8) % 8)
        key = base64.b32decode(secret_cl)
        
        current_time = time.time()
        msg = struct.pack(">Q", int(current_time // 30))
        h = hmac.new(key, msg, hashlib.sha1).digest()
        o = h[19] & 15
        code = (struct.unpack(">I", h[o:o+4])[0] & 0x7fffffff) % 1000000
        otp_raw = f"{code:06d}"
        
        # Thêm khoảng trắng ở giữa để siêu dễ đọc (VD: 123 456)
        otp_display = f"{otp_raw[:3]} {otp_raw[3:]}"
    except Exception:
        otp_display = "ERROR!"
        otp_raw = "ERROR"

    label_otp.config(text=otp_display, font=OTP_FONT)

    # 2. Tính toán thời gian thực
    rem_exact = 30.0 - (current_time % 30.0)
    rem_int = int(rem_exact)

    # 3. Cập nhật Icon Taskbar (Chỉ cập nhật mỗi giây 1 lần để không bị quá tải)
    if rem_int != last_remaining:
        last_remaining = rem_int
        if IS_MINIMIZED and tray_icon is not None:
            try:
                tray_icon.title = f"Mã: {otp_raw} ({rem_int}s)"
            except: pass

    # 4. Hiệu ứng Thanh Tiến Trình (Progress Bar Animation)
    if otp_raw != "ERROR":
        bar_width = max(0, int(130 * (rem_exact / 30.0))) # Dài tối đa 130px
        
        if rem_exact <= 5.0:
            # Nhấp nháy màu đỏ mỗi 0.25 giây
            ms = int(current_time * 4)
            bar_color = "#ff3333" if ms % 2 == 0 else "black"
        else:
            # Màu bình thường theo lựa chọn
            bar_color = COLORS[COLOR_INDEX]
            
        progress_canvas.coords(progress_bar, 0, 0, bar_width, 4)
        progress_canvas.itemconfig(progress_bar, fill=bar_color)

    if not IS_MINIMIZED:
        root.lower()
        
    root.after(100, update_ui_loop)

# --- XỬ LÝ KHAY HỆ THỐNG (SYSTEM TRAY) ---
def create_tray_image():
    image = Image.new('RGB', (64, 64), color='black')
    dc = ImageDraw.Draw(image)
    dc.rectangle((16, 16, 48, 48), fill=COLORS[COLOR_INDEX])
    return image

def restore_from_tray(icon, item):
    global IS_MINIMIZED, tray_icon, LOCKED
    icon.stop()
    tray_icon = None
    IS_MINIMIZED = False
    LOCKED = True 
    root.after(0, update_lock_ui)
    root.after(0, save_config)
    root.after(0, root.deiconify)
    root.after(0, root.lower)

def prompt_exit_from_tray(icon, item):
    root.after(0, show_tray_exit_prompt)

def show_tray_exit_prompt():
    exit_win = tk.Toplevel(root)
    exit_win.title("Xác nhận Thoát")
    exit_win.geometry("220x110")
    exit_win.attributes('-topmost', True)
    
    tk.Label(exit_win, text="Nhập mật khẩu để thoát hoàn toàn:").pack(pady=5)
    entry_pw = tk.Entry(exit_win, show="*", width=20, justify='center')
    entry_pw.pack(pady=5)
    entry_pw.focus()
    
    lbl_error = tk.Label(exit_win, text="", fg="red")
    lbl_error.pack()
    
    def check_pw(event=None):
        if entry_pw.get() == UNLOCK_PASSWORD:
            if tray_icon is not None:
                tray_icon.stop()
            root.destroy()
        else:
            lbl_error.config(text="Sai mật khẩu!")
            entry_pw.delete(0, tk.END)
            
    tk.Button(exit_win, text="Thoát App", command=check_pw, bg="#ff4444", fg="white").place(x=75, y=80)
    exit_win.bind('<Return>', check_pw)

def minimize_app(event=None):
    if LOCKED: return
    global IS_MINIMIZED, tray_icon
    IS_MINIMIZED = True
    root.withdraw()
    
    menu = pystray.Menu(
        pystray.MenuItem('Mở lại Widget (Double Click)', restore_from_tray, default=True),
        pystray.MenuItem('Thoát hoàn toàn', prompt_exit_from_tray)
    )
    tray_icon = pystray.Icon("OTP_Widget", create_tray_image(), "Đang tạo mã...", menu)
    threading.Thread(target=tray_icon.run, daemon=True).start()

# --- XỬ LÝ MÀU SẮC & KHÓA ---
def cycle_color(event=None):
    if LOCKED: return
    global COLOR_INDEX
    COLOR_INDEX = (COLOR_INDEX + 1) % len(COLORS)
    label_otp.config(fg=COLORS[COLOR_INDEX])
    save_config()

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

# --- XỬ LÝ CÀI ĐẶT ---
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

# --- THIẾT LẬP WIDGET ---
root = tk.Tk()
root.geometry("200x70+100+100") 
root.overrideredirect(True) 

root.configure(bg='black')
root.wm_attributes("-transparentcolor", "black")
root.lower()
root.bind("<FocusIn>", lambda e: root.lower() if not IS_MINIMIZED else None) 

root.bind("<ButtonPress-1>", start_move)
root.bind("<ButtonRelease-1>", stop_move)
root.bind("<B1-Motion>", do_move)

# Kéo chữ số lên 1 xíu cho nằm giữa widget
label_otp = tk.Label(root, text="--- ---", font=OTP_FONT, fg=COLORS[COLOR_INDEX], bg="black")
label_otp.place(x=0, y=-2)

# VẼ THANH TIẾN TRÌNH DẠ QUANG (Thay cho chữ thời gian)
progress_canvas = tk.Canvas(root, width=130, height=4, bg="black", highlightthickness=0)
progress_canvas.place(x=5, y=52)
progress_bar = progress_canvas.create_rectangle(0, 0, 130, 4, fill=COLORS[COLOR_INDEX], outline="")

# Gán sự kiện chuột
label_otp.bind("<Double-Button-1>", exit_app)
progress_canvas.bind("<Double-Button-1>", exit_app)
label_otp.bind("<ButtonPress-1>", start_move)
progress_canvas.bind("<ButtonPress-1>", start_move)

# --- BỐ TRÍ NÚT BẤM ---
btn_settings = tk.Label(root, text="⚙", font=("Segoe UI", 12), fg="#aaaaaa", bg="black", cursor="hand2", width=2, pady=2)
btn_settings.place(x=140, y=5)
btn_settings.bind("<Button-1>", open_settings)

btn_color = tk.Label(root, text="🎨", font=("Segoe UI", 11), fg="#aaaaaa", bg="black", cursor="hand2", width=2, pady=2)
btn_color.place(x=170, y=5)
btn_color.bind("<Button-1>", cycle_color)

btn_lock = tk.Label(root, text="🔒" if LOCKED else "🔓", font=("Segoe UI", 11), fg="#ff4444" if LOCKED else "#aaaaaa", bg="black", cursor="hand2", width=2, pady=2)
btn_lock.place(x=140, y=35)
btn_lock.bind("<Button-1>", toggle_lock)

btn_minimize = tk.Label(root, text=" ─ ", font=("Segoe UI", 10, "bold"), fg="#aaaaaa", bg="black", cursor="hand2", width=2, pady=2)
btn_minimize.place(x=170, y=35)
btn_minimize.bind("<Button-1>", minimize_app)

update_ui_loop() # Kích hoạt vòng lặp hiệu ứng 100ms
root.mainloop()
