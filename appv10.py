import tkinter as tk
import hmac, base64, struct, hashlib, time

# THAY CHUỖI SECRET KEY CỦA BẠN VÀO ĐÂY
SECRET = "JBSWY3DPEHPK3PXP" 

def get_totp(secret):
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
        return "ERROR"

def update_otp():
    otp = get_totp(SECRET)
    remaining = 30 - int(time.time() % 30)
    label_otp.config(text=otp)
    label_time.config(text=f"Đổi sau: {remaining}s")
    
    # Liên tục ép Widget nằm dưới cùng (Dính vào hình nền)
    root.lower()
    root.after(1000, update_otp)

# --- XỬ LÝ KÉO THẢ ---
def start_move(event):
    root.x = event.x
    root.y = event.y

def stop_move(event):
    root.x = None
    root.y = None

def do_move(event):
    deltax = event.x - root.x
    deltay = event.y - root.y
    x = root.winfo_x() + deltax
    y = root.winfo_y() + deltay
    root.geometry(f"+{x}+{y}")

# --- THIẾT LẬP WIDGET DESKTOP ---
root = tk.Tk()
# Chỉnh vị trí xuất hiện ban đầu (Cách viền trái 100px, viền trên 100px)
root.geometry("140x60+100+100") 
root.overrideredirect(True) # Xóa viền cửa sổ

# [THỦ THUẬT 1]: Biến màu đen thành TRONG SUỐT HOÀN TOÀN
root.configure(bg='black')
root.wm_attributes("-transparentcolor", "black")

# [THỦ THUẬT 2]: Khóa cứng ở lớp dưới cùng (Desktop)
root.lower()
# Nếu lỡ click trúng, lập tức đẩy nó xuống lại
root.bind("<FocusIn>", lambda e: root.lower()) 

# Gán sự kiện kéo thả và thoát
root.bind("<ButtonPress-1>", start_move)
root.bind("<ButtonRelease-1>", stop_move)
root.bind("<B1-Motion>", do_move)
root.bind("<Double-Button-1>", lambda e: root.destroy())

# Giao diện chữ (Nền đen ở đây sẽ bị Windows biến thành trong suốt)
label_otp = tk.Label(root, text="------", font=("Consolas", 26, "bold"), fg="#00ff00", bg="black")
label_otp.pack(pady=0)

label_time = tk.Label(root, text="Đổi sau: --s", font=("Consolas", 10, "bold"), fg="#ffffff", bg="black")
label_time.pack(pady=0)

update_otp()
root.mainloop()
