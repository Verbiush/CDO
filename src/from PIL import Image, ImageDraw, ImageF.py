import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageDraw, ImageFilter, ImageTk, ImageFont
import pandas as pd
import random, math, os, io

# ---------- CURVAS (Catmull-Rom) ----------
def catmull_rom_spline(P0, P1, P2, P3, n_points=20):
    points = []
    for i in range(n_points):
        t = i / float(n_points)
        t2 = t*t
        t3 = t2*t
        f1 = -0.5*t3 + t2 - 0.5*t
        f2 =  1.5*t3 - 2.5*t2 + 1.0
        f3 = -1.5*t3 + 2.0*t2 + 0.5*t
        f4 =  0.5*t3 - 0.5*t2
        x = P0[0]*f1 + P1[0]*f2 + P2[0]*f3 + P3[0]*f4
        y = P0[1]*f1 + P1[1]*f2 + P2[1]*f3 + P3[1]*f4
        points.append((x,y))
    return points

def smooth_path(raw_pts, samples_per_segment=12):
    if len(raw_pts) < 4:
        return raw_pts
    path = []
    pts = [raw_pts[0]] + raw_pts + [raw_pts[-1]]
    for i in range(len(pts)-3):
        P0, P1, P2, P3 = pts[i], pts[i+1], pts[i+2], pts[i+3]
        segment = catmull_rom_spline(P0,P1,P2,P3, n_points=samples_per_segment)
        path.extend(segment)
    return path

# ---------- EXTRACCIÓN DE LÍNEA CENTRAL A PARTIR DE MÁSCARA DE TEXTO ----------
def text_mask_centerline(text, W=900, H=260, font=None, margin=20):
    """Rasteriza texto en una máscara y devuelve una lista de puntos (x,y) aproximando el centro del trazo."""
    # crear máscara en escala de grises
    mask = Image.new("L", (W, H), 255)  # fondo blanco
    draw = ImageDraw.Draw(mask)
    # si no hay font, usar carga por defecto con tamaño adaptativo
    if font is None:
        # intentar cargar una fuente TrueType del sistema
        try:
            font = ImageFont.truetype("arial.ttf", 140)
        except:
            font = ImageFont.load_default()
    # ajustar tamaño para que el texto quepa
    # empezamos con tamaño grande y reducimos si no cabe
    size = 140
    while True:
        try:
            if isinstance(font, ImageFont.FreeTypeFont):
                font = ImageFont.truetype(font.path if hasattr(font, "path") else "arial.ttf", size)
            else:
                font = ImageFont.load_default()
        except Exception:
            font = ImageFont.load_default()
        tw, th = draw.textsize(text, font=font)
        if tw + 2*margin <= W or size < 10:
            break
        size = int(size * 0.9)
        # recreate font in next loop
        try:
            font = ImageFont.truetype(font.path if hasattr(font, "path") else "arial.ttf", size)
        except Exception:
            font = ImageFont.load_default()

    # centrar el texto horizontalmente y verticalmente
    x = (W - tw) // 2
    y = (H - th) // 2
    draw.text((x, y), text, fill=0, font=font)  # texto negro sobre blanco

    # convertir a pixeles y para cada columna sacar media de y donde pixel < 128
    pix = mask.load()
    center_points = []
    for col in range(W):
        ys = []
        for row in range(H):
            if pix[col, row] < 128:
                ys.append(row)
        if ys:
            mean_y = sum(ys)/len(ys)
            center_points.append((col, mean_y))

    # si no hay puntos (texto no dibujado), devolver vacío
    if not center_points:
        return []

    # reducir puntos: tomar cada N para aligerar
    step = max(1, int(len(center_points) / 200))
    reduced = center_points[::step]

    # normalizar X a margen y ancho (ya está centrado, pero dejamos margen)
    # suavizar
    path = smooth_path(reduced, samples_per_segment=8)
    return path

# Dibuja trazo con grosor variable (simula presión) usando círculos superpuestos
def dibujar_trazo(draw, path, color=(10,10,80), base_width=8, pressure_variation=4):
    if not path:
        return
    L = len(path)
    for i, (x,y) in enumerate(path):
        t = i / max(1, L-1)
        # presión: más fuerte en el centro, menos al inicio/fin
        pressure = base_width + math.sin(t*math.pi) * pressure_variation + random.uniform(-1.0, 1.0)
        r = max(1, abs(pressure))
        bbox = [x - r, y - r, x + r, y + r]
        draw.ellipse(bbox, fill=color)

# Añadir ruido fino (simula papel/tinta)
def add_paper_noise(img, intensity=5):
    W, H = img.size
    noise = Image.effect_noise((W, H), intensity)
    noise = noise.convert("L").point(lambda p: p//6)  # atenuar
    noise_rgb = Image.merge("RGB", (noise, noise, noise))
    return Image.blend(img, Image.composite(img, noise_rgb, noise), 0.05)

# ---------- APLICACIÓN TKINTER ----------
class GeneradorFirmasApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Firmas (Híbrido legible)")
        self.root.geometry("820x720")
        self.root.resizable(False, False)

        self.excel_path = None
        self.df = None
        self.output_dir = "firmas"

        tk.Label(root, text="Generador de Firmas (legible sin TTF / híbrido)", font=("Arial", 16, "bold")).pack(pady=10)

        # fila seleccion archivo
        frame_excel = tk.Frame(root)
        frame_excel.pack(pady=6)
        tk.Button(frame_excel, text="Seleccionar Excel", command=self.cargar_excel).pack(side="left", padx=6)
        self.label_excel = tk.Label(frame_excel, text="Ningún archivo seleccionado", fg="gray")
        self.label_excel.pack(side="left")

        # hoja
        frame_hoja = tk.Frame(root)
        frame_hoja.pack(pady=6)
        tk.Label(frame_hoja, text="Hoja:").pack(side="left", padx=6)
        self.combo_hojas = ttk.Combobox(frame_hoja, state="readonly", width=50)
        self.combo_hojas.pack(side="left")
        self.combo_hojas.bind("<<ComboboxSelected>>", self.mostrar_columnas)

        # columnas
        frame_cols = tk.Frame(root)
        frame_cols.pack(pady=6)
        tk.Label(frame_cols, text="Columna Nombre:").grid(row=0, column=0, padx=6, pady=6, sticky="e")
        tk.Label(frame_cols, text="Columna Apellido:").grid(row=1, column=0, padx=6, pady=6, sticky="e")
        self.combo_nombre = ttk.Combobox(frame_cols, state="readonly", width=40)
        self.combo_apellido = ttk.Combobox(frame_cols, state="readonly", width=40)
        self.combo_nombre.grid(row=0, column=1)
        self.combo_apellido.grid(row=1, column=1)

        # opciones
        frame_opts = tk.Frame(root)
        frame_opts.pack(pady=6)
        self.hybrid_var = tk.BooleanVar(value=True)
        tk.Checkbutton(frame_opts, text="Usar método híbrido (legible sin TTF)", variable=self.hybrid_var).pack()
        tk.Label(frame_opts, text="Si está desactivado dibuja trazos puramente procedurales (menor legibilidad).").pack()

        # TTF opcional
        frame_font = tk.Frame(root)
        frame_font.pack(pady=6)
        tk.Label(frame_font, text="Fuente TTF (opcional):").pack(side="left", padx=6)
        self.entry_font = tk.Entry(frame_font, width=48)
        self.entry_font.pack(side="left")
        tk.Button(frame_font, text="Seleccionar .ttf", command=self.seleccionar_ttf).pack(side="left", padx=6)

        # salida y generar
        frame_out = tk.Frame(root)
        frame_out.pack(pady=8)
        tk.Button(frame_out, text="Seleccionar carpeta de salida", command=self.seleccionar_carpeta).pack(side="left", padx=6)
        self.label_out = tk.Label(frame_out, text=f"Salida: ./{self.output_dir}", fg="gray")
        self.label_out.pack(side="left")

        tk.Button(root, text="Generar Firmas", bg="#1976D2", fg="white", font=("Arial", 12, "bold"), command=self.generar_firmas).pack(pady=10)

        # preview area
        tk.Label(root, text="Vista previa de la firma (fila seleccionada):", font=("Arial", 11)).pack(pady=(8,0))
        self.canvas_preview = tk.Label(root, bd=2, relief="sunken")
        self.canvas_preview.pack(pady=6)

        # controles preview
        frame_preview = tk.Frame(root)
        frame_preview.pack(pady=4)
        tk.Button(frame_preview, text="Cargar fila de ejemplo (primera)", command=self.preview_first).pack(side="left", padx=6)
        tk.Button(frame_preview, text="Generar solo vista previa", command=self.generate_preview_only).pack(side="left", padx=6)

        self.status = tk.Label(root, text="", fg="green")
        self.status.pack(pady=6)

    def cargar_excel(self):
        ruta = filedialog.askopenfilename(title="Selecciona el archivo Excel", filetypes=[("Excel files","*.xlsx *.xls")])
        if not ruta:
            return
        self.excel_path = ruta
        self.label_excel.config(text=os.path.basename(ruta))
        try:
            xls = pd.ExcelFile(ruta)
            self.combo_hojas["values"] = xls.sheet_names
            self.combo_hojas.set("")
            self.combo_nombre.set("")
            self.combo_apellido.set("")
            self.combo_nombre["values"] = []
            self.combo_apellido["values"] = []
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo Excel:\n{e}")

    def mostrar_columnas(self, event=None):
        try:
            hoja = self.combo_hojas.get()
            self.df = pd.read_excel(self.excel_path, sheet_name=hoja)
            columnas = list(self.df.columns)
            self.combo_nombre["values"] = columnas
            self.combo_apellido["values"] = columnas
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron leer las columnas:\n{e}")

    def seleccionar_ttf(self):
        ruta = filedialog.askopenfilename(title="Selecciona fuente TTF", filetypes=[("TTF files","*.ttf")])
        if ruta:
            self.entry_font.delete(0, tk.END)
            self.entry_font.insert(0, ruta)

    def seleccionar_carpeta(self):
        ruta = filedialog.askdirectory(title="Selecciona carpeta de salida")
        if ruta:
            self.output_dir = ruta
            self.label_out.config(text=f"Salida: {ruta}")

    # vista previa para la primera fila
    def preview_first(self):
        if self.df is None:
            messagebox.showwarning("Atención", "Primero selecciona Excel y hoja.")
            return
        row = self.df.iloc[0]
        self._create_and_show_preview(row)

    def generate_preview_only(self):
        sel = self.combo_nombre.get(), self.combo_apellido.get()
        if self.df is None or not sel[0] or not sel[1]:
            messagebox.showwarning("Atención", "Selecciona Excel, hoja y columnas.")
            return
        row = self.df.iloc[0]
        self._create_and_show_preview(row)

    def _create_and_show_preview(self, row):
        col_nombre = self.combo_nombre.get()
        col_apellido = self.combo_apellido.get()
        if not col_nombre or not col_apellido:
            messagebox.showwarning("Atención", "Selecciona las columnas de nombre y apellido.")
            return
        nombre = str(row[col_nombre]).strip()
        apellido = str(row[col_apellido]).strip()
        def norm(s):
            parts = [p for p in s.split() if p]
            parts = [p[0].upper() + p[1:].lower() if len(p)>1 else p.upper() for p in parts]
            return " ".join(parts)
        texto = f"{norm(nombre)} {norm(apellido)}"
        img = self._generate_single(texto, preview=True)
        # mostrar en label
        img_tk = ImageTk.PhotoImage(img.resize((760,140)))
        self.canvas_preview.imgtk = img_tk
        self.canvas_preview.config(image=img_tk)

    def generar_firmas(self):
        if self.df is None:
            messagebox.showwarning("Atención", "Selecciona primero un Excel y una hoja.")
            return
        col_nombre = self.combo_nombre.get()
        col_apellido = self.combo_apellido.get()
        if not col_nombre or not col_apellido:
            messagebox.showwarning("Atención", "Selecciona las columnas de nombre y apellido.")
            return

        os.makedirs(self.output_dir, exist_ok=True)
        hybrid = self.hybrid_var.get()
        font_path = self.entry_font.get().strip()
        total = len(self.df)
        self.status.config(text="Generando...")
        try:
            for idx, row in self.df.iterrows():
                nombre = str(row[col_nombre]).strip()
                apellido = str(row[col_apellido]).strip()
                if not nombre:
                    continue
                def norm(s):
                    parts = [p for p in s.split() if p]
                    parts = [p[0].upper() + p[1:].lower() if len(p)>1 else p.upper() for p in parts]
                    return " ".join(parts)
                texto = f"{norm(nombre)} {norm(apellido)}"
                img = self._generate_single(texto, hybrid=hybrid, font_path=font_path)
                safe_name = "".join(c for c in (nombre + "_" + apellido) if c.isalnum() or c in " _-").strip()
                out_path = os.path.join(self.output_dir, f"firma_{safe_name}_{idx}.jpg")
                img.save(out_path, "JPEG", quality=92)
            self.status.config(text=f"Firmas generadas en '{self.output_dir}'")
            messagebox.showinfo("Listo", f"Firmas generadas ({total} filas). Carpeta: {self.output_dir}")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error:\n{e}")
            self.status.config(text="Error durante la generación.")

    def _generate_single(self, texto, hybrid=True, font_path=None, preview=False):
        W, H = 900, 260
        img = Image.new("RGB", (W, H), (255,255,255))
        draw = ImageDraw.Draw(img)
        # elegir color tinta variantes azul/negro
        color = (random.randint(5,30), random.randint(5,30), random.randint(70,120))

        # intentar cargar font si el usuario dio path
        font_obj = None
        if font_path:
            try:
                font_obj = ImageFont.truetype(font_path, 140)
            except Exception:
                font_obj = None

        if hybrid:
            # generar centerline desde máscara de texto
            # preferir la font si la hay para mejor encaje; si no, usar arial/defecto
            try:
                if font_obj is not None:
                    path = text_mask_centerline(texto, W=W, H=H, font=font_obj)
                else:
                    # intentar cargar arial por defecto; si falla, usar ImageFont.load_default()
                    try:
                        test_font = ImageFont.truetype("arial.ttf", 140)
                    except:
                        test_font = ImageFont.load_default()
                    path = text_mask_centerline(texto, W=W, H=H, font=test_font)
            except Exception:
                path = []

            if not path:
                # fallback a trazo procedural si algo falla
                path = self._procedural_path_for_text(texto, W, H)

            # dibujar trazo variable
            dibujar_trazo(draw, path, color=color, base_width=random.uniform(6,10), pressure_variation=4.0)
            # suavizar y añadir ruido leve
            img = img.filter(ImageFilter.SMOOTH_MORE)
            img = add_paper_noise(img, intensity=6)
        else:
            # puro procedural (menos legible)
            path = self._procedural_path_for_text(texto, W, H)
            dibujar_trazo(draw, path, color=color, base_width=random.uniform(6,10), pressure_variation=4.0)
            img = img.filter(ImageFilter.SMOOTH_MORE)
            img = add_paper_noise(img, intensity=6)

        # si preview, recortar y devolver
        return img

    def _procedural_path_for_text(self, texto, width=900, height=260, margin=40):
        # versión mejorada del procedural simple: intenta respetar posición de letras mediante bloques
        palabras = [p for p in texto.split() if p]
        baseline = height // 2 + random.randint(-8, 8)
        total_chars = sum(len(p) for p in palabras) + max(0, len(palabras)-1)*1
        est_char_w = (width - 2*margin) / max(1, total_chars)
        x = margin
        raw_pts = []
        for ch in texto:
            adv = est_char_w * (0.9 if ch!=" " else 0.6)
            n_sub = max(3, int(adv // 4))
            seed = (ord(ch) + len(texto)) % 97
            amp = 6 + (seed % 8)
            freq = 0.7 + (seed % 6) * 0.15
            for s in range(n_sub):
                t = s / max(1, n_sub-1)
                xpos = x + adv * t
                y_noise = math.sin(t * math.pi * freq * 2 + random.uniform(-0.5,0.5)) * amp
                y = baseline + y_noise + math.sin(xpos * 0.02 + seed) * ( (seed%5) * 0.4 )
                raw_pts.append((xpos, y))
            x += adv
        # jitter
        raw_pts = [(p[0], p[1] + random.uniform(-2.2, 2.2)) for p in raw_pts]
        return smooth_path(raw_pts, samples_per_segment=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = GeneradorFirmasApp(root)
    root.mainloop()
