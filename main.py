import pydirectinput
import pyautogui
import time
import numpy as np
import cv2
import ctypes
from ctypes import wintypes
import threading
import os
import customtkinter as ctk
import random

user32 = ctypes.windll.user32
pydirectinput.FAILSAFE = False

# ====================================================
# SYSTEM CONFIG (พิกัดหน้าจอและการดรอปรูป)
# ====================================================
CONFIG_START_X_PCT = 0.119665
CONFIG_START_Y_PCT = 0.311653
CONFIG_STEP_X_PCT = 0.108341
CONFIG_STEP_Y_PCT = 0.260840
CONFIG_CROP_SCALE = 0.60     

# ====================================================
# SPEED ENGINE (ตั้งค่าความเร็วของโปรแกรม ปรับแต่งได้ตามใจชอบ)
# ====================================================
DELAY_MOUSE_MOVE = 0.25   # ความเร็วตอนลากเมาส์ (วินาที) ยิ่งตัวเลขน้อยยิ่งวาร์ปเร็ว
DELAY_CARD_FLIP = 0.4     # ระยะเวลารออนิเมชั่นไพ่พลิกหงายหน้าให้เสร็จ (วินาที)
DELAY_PAIR_MATCH = 1.0    # ระยะเวลารอกราฟิกตอนจับคู่สำเร็จให้แตกสลายหายไป (วินาที)
DELAY_MISMATCH = 1.0      # ระยะเวลารอให้ไพ่คว่ำหน้ากลับไปสนิท ตอนที่คลิกผิด (วินาที)
DELAY_AFTER_CLICK = 0.15  # หน่วงเวลาสั้นๆ หลังเมาส์คลิกซ้ายเสร็จ (วินาที)

class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

def get_game_window(keyword="seven knights"):
    hwnd_found = None
    def enum_windows_callback(hwnd, lParam):
        nonlocal hwnd_found
        if user32.IsWindowVisible(hwnd):
            length = user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buff = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buff, length + 1)
                if keyword.lower() in buff.value.lower():
                    hwnd_found = hwnd
                    return False
        return True
    
    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows(WNDENUMPROC(enum_windows_callback), 0)
    
    if not hwnd_found:
        return None
        
    rect = wintypes.RECT()
    user32.GetClientRect(hwnd_found, ctypes.byref(rect))
    if rect.right - rect.left == 0:
        return None
        
    pt = POINT(rect.left, rect.top)
    user32.ClientToScreen(hwnd_found, ctypes.byref(pt))
    return hwnd_found, pt.x, pt.y, rect.right - rect.left, rect.bottom - rect.top

def get_card_centers(client_left, client_top, client_width, client_height):
    target_aspect = 16 / 9
    current_aspect = client_width / client_height
    
    offset_x = 0
    offset_y = 0
    base_rw = client_width
    base_rh = client_height
    
    if current_aspect > target_aspect + 0.01:
        new_width = client_height * target_aspect
        offset_x = (client_width - new_width) / 2
        base_rw = new_width
    elif current_aspect < target_aspect - 0.01:
        new_height = client_width / target_aspect
        offset_y = (client_height - new_height) / 2
        base_rh = new_height
        
    centers = []
    for r in range(3):
        for c in range(8):
            cx = client_left + offset_x + (CONFIG_START_X_PCT + c * CONFIG_STEP_X_PCT) * base_rw
            cy = client_top + offset_y + (CONFIG_START_Y_PCT + r * CONFIG_STEP_Y_PCT) * base_rh
            centers.append((int(cx), int(cy)))
            
    crop_size = int(base_rw * CONFIG_STEP_X_PCT * CONFIG_CROP_SCALE)
    return centers, crop_size

def mse(imageA, imageB):
    grayA = cv2.cvtColor(imageA, cv2.COLOR_BGR2GRAY)
    grayB = cv2.cvtColor(imageB, cv2.COLOR_BGR2GRAY)
    blurA = cv2.GaussianBlur(grayA, (5, 5), 0)
    blurB = cv2.GaussianBlur(grayB, (5, 5), 0)
    err = np.sum((blurA.astype("float") - blurB.astype("float")) ** 2)
    err /= float(imageA.shape[0] * imageA.shape[1] * imageA.shape[2]) if len(imageA.shape) == 3 else float(imageA.shape[0] * imageA.shape[1])
    return err

def is_key_pressed(vk):
    return (user32.GetAsyncKeyState(vk) & 0x8000) != 0

def human_click(x, y):
    # ปรับความเร็วลากเมาส์ตามตั้งค่า
    pyautogui.moveTo(x, y, duration=DELAY_MOUSE_MOVE, tween=pyautogui.easeOutQuad)
    time.sleep(0.05) 
    
    user32.mouse_event(0x0002, 0, 0, 0, 0) # LEFTDOWN
    time.sleep(0.05) 
    user32.mouse_event(0x0004, 0, 0, 0, 0) # LEFTUP
    
    time.sleep(DELAY_AFTER_CLICK)

# --- Modern UI Configuration ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class BotApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("7K Card Bot")
        self.geometry("400x380")
        self.attributes("-topmost", True)
        self.resizable(False, False)

        self.grid_columnconfigure(0, weight=1)

        # Title Header
        self.title_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.title_frame.grid(row=0, column=0, padx=20, pady=(15, 5), sticky="ew")

        self.lbl_title = ctk.CTkLabel(self.title_frame, text="🛡️ SEVEN KNIGHTS: REBIRTH 🛡️", font=ctk.CTkFont(family="Impact", size=22, weight="bold"), text_color="#FBC02D")
        self.lbl_title.pack()
        
        self.lbl_subtitle = ctk.CTkLabel(self.title_frame, text="Intelligent Auto Memory Matcher", font=ctk.CTkFont(size=12, slant="italic"), text_color="gray70")
        self.lbl_subtitle.pack()

        # Status Board
        self.status_frame = ctk.CTkFrame(self, corner_radius=15, border_width=2, border_color="#3B8ED0")
        self.status_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        self.lbl_game = ctk.CTkLabel(self.status_frame, text="🔍 Game Status: Searching...", font=ctk.CTkFont(size=15, weight="bold"), text_color="#FF6B6B")
        self.lbl_game.pack(pady=(15, 2))
        
        self.lbl_status = ctk.CTkLabel(self.status_frame, text="Ready (Ensure game is visible) | [F10] to Exit", font=ctk.CTkFont(size=13))
        self.lbl_status.pack(pady=(2, 15))

        # Main Controls
        self.btn_scan = ctk.CTkButton(self, text="📸 1. Record Cards (F8)", font=ctk.CTkFont(size=15, weight="bold"), height=45, corner_radius=8, command=self.on_scan)
        self.btn_scan.grid(row=2, column=0, padx=20, pady=(10, 5), sticky="ew")

        self.btn_click = ctk.CTkButton(self, text="⚔️ 2. Auto Match (F9)", font=ctk.CTkFont(size=15, weight="bold"), height=45, corner_radius=8, command=self.on_click, state="disabled", fg_color="#4CAF50", hover_color="#45a049")
        self.btn_click.grid(row=3, column=0, padx=20, pady=(5, 5), sticky="ew")

        # Credit Footer
        self.footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.footer_frame.grid(row=4, column=0, padx=20, pady=(0, 10), sticky="ew")
        
        self.lbl_credit = ctk.CTkLabel(self.footer_frame, text="Created By Apotoxin", font=ctk.CTkFont(family="Arial", size=11, weight="bold", slant="italic"), text_color="#FFA726")
        self.lbl_credit.pack(side="right")

        # Bot State
        self.pairs = []
        self.game_hwnd = None
        self.game_rect = None
        self.last_f8 = 0
        self.last_f9 = 0
        self.scanning = False

        self.update_game_status()
        self.check_hotkeys()

    def update_game_status(self):
        win_info = get_game_window("seven knights")
        if win_info:
            hwnd, cx, cy, cw, ch = win_info
            self.lbl_game.configure(text=f"🎮 Game Detected ({cw}x{ch})", text_color="#4CAF50")
            self.game_hwnd = hwnd
            self.game_rect = (cx, cy, cw, ch)
            if self.lbl_status.cget("text") == "No game window found":
                self.lbl_status.configure(text="Ready (Ensure game is visible) | [F10] to Exit")
            if not self.scanning:
                self.btn_scan.configure(state="normal")
        else:
            self.lbl_game.configure(text="❌ Game Not Found", text_color="#FF6B6B")
            self.lbl_status.configure(text="No game window found")
            self.game_hwnd = None
            self.game_rect = None
            if not self.scanning:
                self.btn_scan.configure(state="disabled")
            
        self.after(1000, self.update_game_status)

    def check_hotkeys(self):
        now = time.time()
        if is_key_pressed(0x77) and now - self.last_f8 > 1:
            self.last_f8 = now
            if self.btn_scan.cget("state") == "normal":
                self.on_scan()
                
        if is_key_pressed(0x78) and now - self.last_f9 > 1:
            self.last_f9 = now
            if self.btn_click.cget("state") == "normal":
                self.on_click()
                
        # F10 Detection
        if is_key_pressed(0x79):
            self.destroy()
            os._exit(0)
                
        self.after(30, self.check_hotkeys)

    def on_scan(self):
        if not self.game_rect:
            return
            
        if getattr(self, 'scanning', False):
            self.scanning = False
            return
            
        self.scanning = True
        self.btn_scan.configure(text="⏹️ Stop Recording (F8)", fg_color="#FFA726", hover_color="#FF9800", text_color="black")
        self.lbl_status.configure(text="Recording... (Press F8 when all flipping ends)", text_color="#42A5F5")
        self.btn_click.configure(state="disabled")
        self.update()
        
        def scan_worker():
            cx, cy, cw, ch = self.game_rect
            centers, crop_size = get_card_centers(cx, cy, cw, ch)
            self.centers = centers
            self.crop_size = crop_size
            half = crop_size // 2
            actual_crop_size = half * 2
            
            scr = pyautogui.screenshot()
            base_img = cv2.cvtColor(np.array(scr), cv2.COLOR_RGB2BGR)
            
            baseline_crops = []
            best_crops = []
            best_diffs = [0] * 24
            
            try:
                for (ptx, pty) in centers:
                    top = max(0, pty - half)
                    bottom = pty + half
                    left = max(0, ptx - half)
                    right = ptx + half
                    
                    if top < 0 or bottom > base_img.shape[0] or left < 0 or right > base_img.shape[1]:
                        def show_err():
                            self.lbl_status.configure(text="Error: Game window out of bounds", text_color="#FF6B6B")
                            self.btn_scan.configure(text="📸 1. Record Cards (F8)", fg_color=["#3B8ED0", "#1F6AA5"], text_color=["gray10", "#DCE4EE"], state="normal")
                        self.after(0, show_err)
                        self.scanning = False
                        return
                        
                    crop = base_img[top:bottom, left:right]
                    baseline_crops.append(crop)
                    best_crops.append(crop)
                
                while getattr(self, 'scanning', False):
                    scr = pyautogui.screenshot()
                    frame = cv2.cvtColor(np.array(scr), cv2.COLOR_RGB2BGR)
                    for i, (ptx, pty) in enumerate(centers):
                        top = max(0, pty - half)
                        bottom = pty + half
                        left = max(0, ptx - half)
                        right = ptx + half
                        crop = frame[top:bottom, left:right]
                        diff = mse(crop, baseline_crops[i])
                        if diff > best_diffs[i]:
                            best_diffs[i] = diff
                            best_crops[i] = crop
                    time.sleep(0.01)
            except Exception as e:
                print("Error during scan:", e)
                
            matched = set()
            local_pairs = []
            for i in range(24):
                if i in matched: continue
                best_match = -1
                best_mse = float('inf')
                for j in range(i + 1, 24):
                    if j in matched: continue
                    err = mse(best_crops[i], best_crops[j])
                    if err < best_mse:
                        best_mse = err
                        best_match = j
                if best_match != -1:
                    local_pairs.append((i, best_match)) # เก็บเป็น Index ตัวเลขใบไพ่แทนอักษรพิกัด เพื่อให้เช็คซ้ำง่ายๆ
                    matched.add(i)
                    matched.add(best_match)
                    
            self.pairs = local_pairs
            self.baseline_crops = baseline_crops # ลายหลังไพ่เก็บไว้เทียบ
            self.best_crops = best_crops # ลายหน้าไพ่
            
            def finish_scan():
                self.btn_scan.configure(text="🔄 1. Re-Record (F8)", fg_color=["#3B8ED0", "#1F6AA5"], text_color=["gray10", "#DCE4EE"])
                self.lbl_status.configure(text=f"Success! {len(local_pairs)} pairs found. Ready.", text_color="#4CAF50")
                self.btn_click.configure(state="normal")
                
            self.after(0, finish_scan)
        threading.Thread(target=scan_worker, daemon=True).start()

    def on_click(self):
        if not self.pairs:
            return
        cv2.destroyAllWindows()
        self.lbl_status.configure(text="Matching pairs... Please wait.", text_color="#42A5F5")
        self.btn_scan.configure(state="disabled")
        self.btn_click.configure(state="disabled")
        
        def clicker_thread():
            if self.game_hwnd:
                try: user32.SetForegroundWindow(self.game_hwnd)
                except: pass
                time.sleep(0.2)
                
            pairs_to_click = list(self.pairs)
            retry_count = 0
            
            def is_card_face_up(idx):
                scr = pyautogui.screenshot()
                frame = cv2.cvtColor(np.array(scr), cv2.COLOR_RGB2BGR)
                half = self.crop_size // 2
                px, py = self.centers[idx]
                c = frame[max(0, py-half):py+half, max(0, px-half):px+half]
                diff = mse(c, self.baseline_crops[idx])
                return diff > 100.0
                
            while pairs_to_click and retry_count < 5:
                random.shuffle(pairs_to_click) # สุ่มลำดับคู่ไพ่ ให้เหมือนคนเล่น
                
                # สุ่มสลับว่าในแต่ละคู่ จะเปิดใบไหนก่อน (แก้ปัญหาบอทชอบเปิดจากซ้ายไปขวาเสมอ)
                for i in range(len(pairs_to_click)):
                    if random.random() > 0.5:
                        pairs_to_click[i] = (pairs_to_click[i][1], pairs_to_click[i][0])
                        
                self.lbl_status.configure(text=f"Matching {len(pairs_to_click)} pairs (Attempt {retry_count+1})...", text_color="#42A5F5")
                
                for idx1, idx2 in pairs_to_click:
                    p1 = self.centers[idx1]
                    p2 = self.centers[idx2]
                    
                    # สุ่มพิกัดให้เหลื่อมจากจุดกึ่งกลาง (Jitter) ระดับ Pixel
                    jitter = self.crop_size // 3
                    
                    # 1. พยายามเปิดใบแรก
                    success1 = False
                    for attempts1 in range(2):
                        if attempts1 > 0:
                            cx, cy = pyautogui.position()
                            pyautogui.moveTo(cx + 15, cy + 15, duration=0.1)
                            
                        rx1 = p1[0] + random.randint(-jitter, jitter)
                        ry1 = p1[1] + random.randint(-jitter, jitter)
                        human_click(rx1, ry1)
                        time.sleep(DELAY_CARD_FLIP)
                        if is_card_face_up(idx1):
                            success1 = True
                            break
                            
                    if not success1:
                        continue # ข้ามคู่นี้ไปก่อนเดี๋ยวระบบ auto recovery รวบยอดตอนท้ายมันจะหยิบมาทำใหม่
                        
                    time.sleep(DELAY_CARD_FLIP) # รอให้ใบที่ 1 กางสุดๆ นิ่งๆ แบบ 100%
                    
                    # --- ระบบตรวจค่ายืนยันหน้าไพ่ (Dynamic Verification) ---
                    valid_candidates = set()
                    for pair in pairs_to_click:
                        valid_candidates.add(pair[0])
                        valid_candidates.add(pair[1])
                        
                    scr = pyautogui.screenshot()
                    frame = cv2.cvtColor(np.array(scr), cv2.COLOR_RGB2BGR)
                    half = self.crop_size // 2
                    current_face1 = frame[max(0, p1[1]-half):p1[1]+half, max(0, p1[0]-half):p1[0]+half]
                    
                    forced_skip = False
                    diff_expected = mse(current_face1, self.best_crops[idx2])
                    
                    # ถ้าภาพไพ่ใบที่เปิดขึ้นมา ดันไม่ตรงกับข้อมูลใบแฝด (เกมหลอก หรือระบบแสกนพลาด)
                    if diff_expected > 1500.0:  
                        best_diff = float('inf')
                        dynamic_idx2 = -1
                        for j in valid_candidates:
                            if j == idx1: continue
                            diff = mse(current_face1, self.best_crops[j])
                            if diff < best_diff:
                                best_diff = diff
                                dynamic_idx2 = j
                                
                        # ถ้าเจอแฝดที่แท้จริง ก็สลับเป้าหมายไปหาแฝดแท้เลย!
                        if dynamic_idx2 != -1 and best_diff < 1500.0:
                            idx2 = dynamic_idx2
                            p2 = self.centers[idx2]
                        else:
                            forced_skip = True # หาแฝดแท้ไม่เจอ ยอมแพ้
                            
                    if forced_skip:
                        # ยอมเปิดไพ่มั่วๆ ไป 1 ใบ (ผิด 1 ครั้ง) เพื่อให้ใบแรกคว่ำลงไป
                        for dummy in valid_candidates:
                            if dummy != idx1:
                                dx = self.centers[dummy][0] + random.randint(-jitter, jitter)
                                dy = self.centers[dummy][1] + random.randint(-jitter, jitter)
                                human_click(dx, dy)
                                time.sleep(DELAY_MISMATCH)
                                break
                        continue
                    # ----------------------------------------------------
                    
                    # 2. ลองเปิดใบแฝด 
                    success2 = False
                    for attempts2 in range(2):
                        if attempts2 > 0:
                            cx, cy = pyautogui.position()
                            pyautogui.moveTo(cx + 15, cy + 15, duration=0.1)
                            
                        rx2 = p2[0] + random.randint(-jitter, jitter)
                        ry2 = p2[1] + random.randint(-jitter, jitter)
                        human_click(rx2, ry2)
                        time.sleep(DELAY_CARD_FLIP)
                        if is_card_face_up(idx2):
                            success2 = True
                            break
                            
                    if not success2:
                        time.sleep(DELAY_MISMATCH) # ใบแรกหงายเก้อ รอให้มันคว่ำหน้ากลับไปก่อนค่อยข้ามเดี๋ยวเกมรวน
                        continue
                        
                    time.sleep(DELAY_PAIR_MATCH) 
                
                # --- Auto Recovery & Orphan Collector ---
                self.lbl_status.configure(text="Verifying remaining cards...", text_color="#FFA726")
                time.sleep(1.5) # รอพวกที่ไม่จับคู่คว่ำกลับไปสนิทๆ ก่อน
                scr = pyautogui.screenshot()
                frame = cv2.cvtColor(np.array(scr), cv2.COLOR_RGB2BGR)
                half = self.crop_size // 2
                
                # หาว่ามีไพ่ใบไหนบนกระดานที่ยัง "คว่ำหน้า" อยู่ทั้งหมดบนลาน
                remaining_facedown = []
                for i in range(24):
                    px, py = self.centers[i]
                    c = frame[max(0, py-half):py+half, max(0, px-half):px+half]
                    diff = mse(c, self.baseline_crops[i])
                    if diff < 150.0: # รูปปัจจุบันยังเป็นหลังไพ่แสดงว่ายังหลงเหลือ
                        remaining_facedown.append(i)
                        
                if not remaining_facedown:
                    pairs_to_click = []
                    break
                    
                # สร้างคู่จับใหม่ให้เฉพาะไพ่กำพร้า คัดเลือกจากรูปที่ดีที่สุดที่บอทจำไว้
                new_pairs = []
                matched_now = set()
                for i in remaining_facedown:
                    if i in matched_now: continue
                    best_match = -1
                    best_diff = float('inf')
                    for j in remaining_facedown:
                        if i == j or j in matched_now: continue
                        diff = mse(self.best_crops[i], self.best_crops[j])
                        if diff < best_diff:
                            best_diff = diff
                            best_match = j
                            
                    if best_match != -1:
                        new_pairs.append((i, best_match))
                        matched_now.add(i)
                        matched_now.add(best_match)
                        
                pairs_to_click = new_pairs
                retry_count += 1
                
            def end_click():
                if pairs_to_click:
                    self.lbl_status.configure(text="Finished with some unrecoverable errors.", text_color="#FF6B6B")
                else:
                    self.lbl_status.configure(text="Match Complete! Flawless victory.", text_color="#4CAF50")
                self.pairs = []
                self.btn_scan.configure(state="normal")
                self.btn_click.configure(state="disabled")
            
            self.after(0, end_click)
        threading.Thread(target=clicker_thread, daemon=True).start()

if __name__ == "__main__":
    import sys
    def is_admin():
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    if not is_admin():
        # ถ้าเผลอเปิดธรรมดา ให้ระบบเด้งหน้าต่าง UAC ขอสิทธิ์ Admin อัตโนมัติ (บังคับใช้รันโหมด Admin)
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        os._exit(0)

    app = BotApp()
    app.mainloop()
