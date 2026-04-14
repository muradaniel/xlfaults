import tkinter as tk
import time
import math

root = None
running = False

def abrir_loading(texto="Processando"):
    global root, running

    root = tk.Tk()

    largura = 260
    altura = 140

    # Centralizar
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (largura // 2)
    y = (root.winfo_screenheight() // 2) - (altura // 2)
    root.geometry(f"{largura}x{altura}+{x}+{y}")

    root.overrideredirect(True)
    root.configure(bg="#121212")

    canvas = tk.Canvas(root, width=largura, height=altura,
                       bg="#121212", highlightthickness=0)
    canvas.pack()

    canvas.create_text(largura/2, 40,
                       text="Processando",
                       fill="#e0e0e0",
                       font=("Segoe UI Semibold", 13))

    canvas.create_text(largura/2, 65,
                       text=f"{texto}...",
                       fill="#888",
                       font=("Segoe UI", 9))

    dots = []
    for i in range(3):
        dot = canvas.create_oval(100 + i*25, 90, 112 + i*25, 102,
                                 fill="#2a2a2a", outline="")
        dots.append(dot)

    start_time = time.time()
    running = True

    def animar():
        global running
        if not running:
            return

        t = time.time() - start_time

        for i, dot in enumerate(dots):
            intensidade = (math.sin(t * 4 + i) + 1) / 2
            cor = int(50 + intensidade * 150)
            cor_hex = f"#{cor:02x}{cor:02x}{255:02x}"
            canvas.itemconfig(dot, fill=cor_hex)

        root.after(50, animar)

    animar()
    root.update()  # 👈 mostra sem travar

def fechar_loading():
    global root, running
    running = False
    if root:
        root.destroy()
        root = None