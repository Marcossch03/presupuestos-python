import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document
from datetime import datetime
import sqlite3

# ==================== 1️⃣ INICIALIZAR BASE DE DATOS ====================
def init_db():
    conn = sqlite3.connect("clientes.db")
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS clientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company TEXT NOT NULL UNIQUE,
            domicilio TEXT,
            localidad TEXT
        )
    ''')

    conn.commit()
    conn.close()

init_db()

# ==================== 2️⃣ FUNCIÓN PARA CONVERTIR NÚMEROS A TEXTO (MEJORADA) ====================
def numero_a_texto(numero):
    """Convierte un número con formato (1.500.000,00) a texto en español correctamente."""
    if isinstance(numero, (int, float)):
        numero = str(numero)

    try:
        numero = float(numero.replace(".", "").replace(",", "."))
    except ValueError:
        return "Número inválido"

    entero = int(numero)
    centavos = int(round((numero - entero) * 100))
    partes = []

    escalas = [
        (1_000_000_000_000, "billón", "billones"),
        (1_000_000_000, "mil millones", "mil millones"),
        (1_000_000, "millón", "millones"),
        (1_000, "mil", "mil")
    ]

    for valor, singular, plural in escalas:
        cantidad = entero // valor
        if cantidad:
            texto = convertir_menores_mil(cantidad)
            if valor == 1_000:
                partes.append("mil" if cantidad == 1 else f"{texto} mil")
            elif cantidad == 1:
                partes.append(f"un {singular}")
            else:
                partes.append(f"{texto} {plural}")
            entero %= valor

    if entero > 0 or not partes:
        partes.append(convertir_menores_mil(entero))

    texto_final = " ".join(partes).strip()

    if centavos > 0:
        texto_final += f" con {centavos:02d}/100"
    else:
        texto_final += " con 00/100"

    return "Son pesos " + texto_final.capitalize()

# ==================== 2️⃣ FUNCIÓN AUXILIAR PARA NÚMEROS < 1000 (MEJORADA) ====================
def convertir_menores_mil(numero):
    """Convierte números menores a 1000 a texto correctamente."""
    unidades = ["", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve"]
    especiales = ["diez", "once", "doce", "trece", "catorce", "quince",
                  "dieciséis", "diecisiete", "dieciocho", "diecinueve"]
    decenas = ["", "diez", "veinte", "treinta", "cuarenta", "cincuenta",
               "sesenta", "setenta", "ochenta", "noventa"]
    centenas = ["", "ciento", "doscientos", "trescientos", "cuatrocientos",
                "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"]

    texto = ""

    if numero == 100:
        return "cien"

    if numero >= 100:
        texto += centenas[numero // 100] + " "
        numero %= 100

    if 10 <= numero < 20:
        texto += especiales[numero - 10]
    elif 21 <= numero <= 29:
        texto += "veinti" + unidades[numero % 10]
    else:
        if numero >= 20:
            texto += decenas[numero // 10]
            if numero % 10 != 0:
                texto += " y "
        if numero % 10 > 0:
            texto += unidades[numero % 10]

    return texto.strip()


# ==================== 3️⃣ ACTUALIZAR PRECIO EN LETRAS AUTOMÁTICAMENTE ====================
def actualizar_precio_texto(event):
    """Convierte el número en el campo de precio a texto y lo muestra en el campo de precio en letras"""
    numero = price_entry.get()
    if numero:
        texto_convertido = numero_a_texto(numero)
        price_text_entry.delete(0, tk.END)
        price_text_entry.insert(0, texto_convertido)

# ==================== 4️⃣ GENERACIÓN DEL DOCUMENTO WORD ====================
def save_to_word():
    """Genera un documento Word basado en la plantilla seleccionada"""
    date = datetime.now().strftime("%d de %B de %Y")

    company1 = company1_entry.get().strip()
    company2 = company2_entry.get().strip()
    company3 = company3_entry.get().strip()
    company4 = company4_entry.get().strip()
    company5 = company5_entry.get().strip()
    reference = reference_entry.get().strip()
    price = price_entry.get().strip()
    price_text = numero_a_texto(price)

    if not company1 or not reference or not price:
        messagebox.showerror("Error", "Debe ingresar al menos la compañía, referencia y precio.")
        return

    template_path = filedialog.askopenfilename(title="Selecciona la plantilla", filetypes=[("Word Documents", "*.docx")])
    if not template_path:
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
    if not file_path:
        return

    doc = Document(template_path)

    for para in doc.paragraphs:
        para.text = para.text.replace("{DateModify}", date)
        para.text = para.text.replace("{Company}", company1)
        para.text = para.text.replace("{domicilio}", company2)
        para.text = para.text.replace("{localidad}", company3)
        para.text = para.text.replace("{tipo_servicio}", company4)
        para.text = para.text.replace("{Mensual o Meses}", company5)
        para.text = para.text.replace("{INSERTAR PRECIO Y CONDICION DE IVA INC o +IVA)}", price)
        para.text = para.text.replace("{(VALOR PRESUPUESTO EN LETRAS y CONDICION IVA)}", price_text)

    doc.save(file_path)
    messagebox.showinfo("Éxito", "Documento generado correctamente.")

# ==================== 5️⃣ INTERFAZ GRÁFICA ====================
def create_gui():
    global company1_entry, company2_entry, company3_entry, company4_entry, company5_entry
    global reference_entry, price_entry, price_text_entry

    root = tk.Tk()
    root.title("Gestión de Documentos")
    root.geometry("600x400")
    root.configure(bg="#f0f0f0")

    frame = ttk.Frame(root, padding=20)
    frame.pack(expand=True)

    campos = [
        ("Ingrese Compañía:", "company1_entry"),
        ("Ingrese Domicilio:", "company2_entry"),
        ("Ingrese Localidad:", "company3_entry"),
        ("Ingrese Tipo de Servicio:", "company4_entry"),
        ("Ingrese Cantidad de meses o Mensual:", "company5_entry"),
        ("Ingrese Referencia:", "reference_entry"),
        ("Ingrese Precio:", "price_entry"),
        ("Precio en Letras:", "price_text_entry")
    ]

    global_vars = {}

    for i, (label, var_name) in enumerate(campos, start=0):
        ttk.Label(frame, text=label).grid(row=i, column=0, sticky="w", pady=5)
        global_vars[var_name] = ttk.Entry(frame, width=50)
        global_vars[var_name].grid(row=i, column=1, pady=5)

    globals().update(global_vars)

    price_entry.bind("<KeyRelease>", actualizar_precio_texto)

    ttk.Button(frame, text="Guardar Cliente").grid(row=len(campos), column=0, pady=10)
    ttk.Button(frame, text="Guardar en Word", command=save_to_word).grid(row=len(campos), column=1, pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
