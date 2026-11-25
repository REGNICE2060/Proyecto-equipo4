import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date
import mysql.connector # type: ignore
import os
import fitz
#Uriel hizo este comentario desde la cuenta de Regina 
#URIEL VAZQUEZ 
from PIL import Image, ImageTk

# ===== CONEXION XAMPP =====
con = mysql.connector.connect(
    host="localhost",
    user="root",
    password="",
    database="proyecto1",
    port=3306
)
cur = con.cursor()

# Crear tablas si no existen
cur.execute("CREATE TABLE IF NOT EXISTS docentes (id INT AUTO_INCREMENT PRIMARY KEY, nombre VARCHAR(100) UNIQUE, grado VARCHAR(50))")
cur.execute("CREATE TABLE IF NOT EXISTS directores (id INT AUTO_INCREMENT PRIMARY KEY, nombre VARCHAR(100) UNIQUE)")
cur.execute("CREATE TABLE IF NOT EXISTS alumnos (id INT AUTO_INCREMENT PRIMARY KEY, nombre VARCHAR(100), grado VARCHAR(50), estado VARCHAR(50))")
cur.execute("CREATE TABLE IF NOT EXISTS asistencia (alumno VARCHAR(100), grado VARCHAR(50), retardos INT DEFAULT 0, faltas INT DEFAULT 0, PRIMARY KEY(alumno, grado))")
cur.execute("CREATE TABLE IF NOT EXISTS incidencias (id INT AUTO_INCREMENT PRIMARY KEY, alumno VARCHAR(100), grado VARCHAR(50), motivo VARCHAR(255), fecha VARCHAR(20), docente VARCHAR(100))")
cur.execute("CREATE TABLE IF NOT EXISTS pendientes (id INT AUTO_INCREMENT PRIMARY KEY, texto VARCHAR(255))")
con.commit()

#VALENTINA REYES GUTIERREZ 
import openpyxl

def carga_masiva(nombre_excel):
    wb = openpyxl.load_workbook(nombre_excel)
    ws = wb.active

    encabezados = [celda.value for celda in ws[1]]

    for fila in ws.iter_rows(min_row=2, values_only=True):
        registro = dict(zip(encabezados, fila))
        print(registro)

carga_masiva("datos.xlsx")

# Datos iniciales
def inicializar():
    cur.execute("SELECT COUNT(*) FROM docentes")
    if cur.fetchone()[0]==0:
        cur.executemany("INSERT INTO docentes (nombre, grado) VALUES (%s,%s)", [
            ("Prof. Maria Lopez","Primero"),("Prof. Juan Perez","Segundo"),("Prof. Ana Torres","Tercero"),
            ("Prof. Carlos Ruiz","Cuarto"),("Prof. Fernanda Gomez","Quinto"),("Prof. Luis Mendoza","Sexto")])
    cur.execute("SELECT COUNT(*) FROM directores")
    if cur.fetchone()[0]==0:
        cur.execute("INSERT INTO directores (nombre) VALUES (%s)", ("Director General",))
    cur.execute("SELECT COUNT(*) FROM alumnos")
    if cur.fetchone()[0]==0:
        nombres=["Juan Perez","Ana Lopez","Carlos Garcia","Maria Hernandez","Luis Torres","Sofia Ramirez","Jorge Cruz","Fernanda Diaz","Pedro Morales","Valeria Ortiz"]
        for g in ["Primero","Segundo","Tercero","Cuarto","Quinto","Sexto"]:
            for n in nombres:
                cur.execute("INSERT INTO alumnos (nombre, grado, estado) VALUES (%s,%s,%s)", (n,g,"✅ Asistencia"))
                cur.execute("INSERT IGNORE INTO asistencia (alumno, grado) VALUES (%s,%s)", (n,g))
    cur.execute("SELECT COUNT(*) FROM pendientes")
    if cur.fetchone()[0]==0:
        for p in ["📝 Revisar 3 reportes","📋 Aprobar asistencia","📚 Firmar boletas","👨‍👩‍👧 Junta padres 10am","📢 Honores lunes - Sexto"]:
            cur.execute("INSERT INTO pendientes (texto) VALUES (%s)", (p,))
    con.commit()
inicializar()

#MIGUEL SALVADOR 
from PIL import Image, ImageTk

def get_docentes():
    cur.execute("SELECT nombre, grado FROM docentes"); return dict(cur.fetchall())
def get_directores():
    cur.execute("SELECT nombre FROM directores"); return [r[0] for r in cur.fetchall()]
def get_alumnos(grado):
    cur.execute("SELECT nombre, grado, estado FROM alumnos WHERE grado=%s", (grado,)); return [[a,b,c] for a,b,c in cur.fetchall()]
def get_contador(grado, alumno):
    cur.execute("SELECT retardos, faltas FROM asistencia WHERE alumno=%s AND grado=%s", (alumno, grado))
    r=cur.fetchone(); return {"retardos":r[0],"faltas":r[1]} if r else {"retardos":0,"faltas":0}
def set_estado(alumno, grado, estado):
    cur.execute("UPDATE alumnos SET estado=%s WHERE nombre=%s AND grado=%s", (estado, alumno, grado)); con.commit()
def set_contador(alumno, grado, ret, fal):
    cur.execute("UPDATE asistencia SET retardos=%s, faltas=%s WHERE alumno=%s AND grado=%s", (ret, fal, alumno, grado)); con.commit()
def get_pendientes():
    cur.execute("SELECT id, texto FROM pendientes"); return cur.fetchall()
def add_pendiente(t):
    cur.execute("INSERT INTO pendientes (texto) VALUES (%s)", (t,)); con.commit()
def del_pendiente(pid):
    cur.execute("DELETE FROM pendientes WHERE id=%s", (pid,)); con.commit()
def get_incidencias():
    cur.execute("SELECT alumno, grado, motivo, fecha, docente FROM incidencias"); return cur.fetchall()
def add_incidencia(a,g,m,f,d):
    cur.execute("INSERT INTO incidencias (alumno, grado, motivo, fecha, docente) VALUES (%s,%s,%s,%s,%s)", (a,g,m,f,d)); con.commit()

# ===== INTERFAZ =====
ROSA="#8B1B3D"; AZUL="#1E293B"; FONDO="#B1B4A3"; BLANCO="white"
VERDE="#10B981"; ROJO="#EF4444"; AMARILLO="#F59E0B"; BORDE_ROSA="#F9A8D4"; ROSA_CLARO="#FFF1F2"; GUINDA_BOTON="#7A1A35"
COLORES_GRADOS=["#E85D5D","#2EB5A6","#1E9BC1","#E8A87C","#7EC8BB","#F2C94C"]
FUENTE_TITULO=("Arial",18,"bold"); FUENTE_SUB=("Arial",14,"bold"); FUENTE_NORMAL=("Arial",13,"bold")
FUENTE_LABEL=("Arial",13,"bold"); FUENTE_BOTON=("Arial",13,"bold"); FUENTE_CHICA=("Arial",12,"bold")

RUTA_GLOSARIO = os.path.join(os.path.dirname(__file__), "glosario.pdf")

root=tk.Tk(); root.title("ESCUELA PRIMARIA NEZAHUALCOYOTL"); root.configure(bg=FONDO)
try: root.state('zoomed')
except: root.attributes('-zoomed', True)
root.geometry("1200x700")
main=tk.Frame(root, bg=FONDO); main.pack(fill="both", expand=True)
GRADOS=["Primero","Segundo","Tercero","Cuarto","Quinto","Sexto"]
docente_grado_map=get_docentes(); lista_docentes_global=list(docente_grado_map.keys()); lista_directores=get_directores()

login_frame=tk.Frame(main, bg=FONDO); director_frame=tk.Frame(main, bg=BLANCO); docente_frame=tk.Frame(main, bg=BLANCO)
login_frame.pack(fill="both", expand=True)
header_login=tk.Frame(login_frame, bg=ROSA, height=50); header_login.pack(fill="x")
tk.Label(header_login, text="🏫 ESCUELA PRIMARIA - SISTEMA 2026 🎒 ", bg=ROSA, fg=BLANCO, font=("Arial",14,"bold")).pack(side="left", padx=20, pady=12)
caja=tk.Frame(login_frame, bg=BLANCO, padx=32, pady=28, highlightbackground=BORDE_ROSA, highlightthickness=2); caja.place(relx=0.5, rely=0.5, anchor="center", width=500)
usuario_var=tk.StringVar(value="👇 Selecciona usuario"); usuario_real=tk.StringVar(value=""); nombre_seleccionado=tk.StringVar(value="")
tk.Label(caja, text="👤 Usuario:", bg=BLANCO, anchor="w", font=FUENTE_LABEL, fg=AZUL).pack(fill="x")
combo_container=tk.Frame(caja, bg=BLANCO, highlightbackground=BORDE_ROSA, highlightthickness=2); combo_container.pack(fill="x", pady=8, ipady=9)
lbl_selected=tk.Label(combo_container, textvariable=usuario_var, bg=BLANCO, anchor="w", font=("Arial",12,"bold"), fg=AZUL, padx=12); lbl_selected.pack(side="left", fill="both", expand=True)
canvas_flecha=tk.Canvas(combo_container, width=34, height=34, bg=BLANCO, highlightthickness=0); canvas_flecha.pack(side="right", padx=4)
canvas_flecha.create_polygon(17,5, 29,17, 17,29, 5,17, fill=ROSA, outline=ROSA); canvas_flecha.create_text(17,17, text="♦", fill=BLANCO, font=("Arial",13,"bold"))
dropdown=None
def toggle_dropdown(event=None):
    global dropdown
    if dropdown and dropdown.winfo_exists(): dropdown.destroy(); dropdown=None; return
    dropdown=tk.Toplevel(root); dropdown.overrideredirect(True); dropdown.configure(bg=BLANCO, highlightbackground=BORDE_ROSA, highlightthickness=2)
    x=combo_container.winfo_rootx(); y=combo_container.winfo_rooty()+combo_container.winfo_height()+5; w=combo_container.winfo_width()
    h=45*(len(lista_docentes_global)+len(lista_directores))+10; dropdown.geometry(f"{w}x{h}+{x}+{y}")
    def seleccionar(rol, nombre, texto): usuario_var.set(texto); usuario_real.set(rol); nombre_seleccionado.set(nombre); dropdown.destroy()
    for d in lista_directores:
        op=tk.Frame(dropdown, bg=BLANCO, padx=12, pady=11); op.pack(fill="x")
        tk.Label(op, text=f"👨‍💼 {d} - 6 Grados", bg=BLANCO, font=("Arial",11,"bold"), fg=AZUL).pack(side="left", fill="x", expand=True)
        op.bind("<Button-1>", lambda e, dd=d: seleccionar("director", dd, f"👨‍💼 {dd}"))
    for doc in lista_docentes_global:
        grado=docente_grado_map.get(doc, "Sin grado")
        op=tk.Frame(dropdown, bg=BLANCO, padx=12, pady=11); op.pack(fill="x")
        tk.Label(op, text=f"👩‍🏫 {doc} - {grado}", bg=BLANCO, font=("Arial",11,"bold"), fg=AZUL).pack(side="left", fill="x", expand=True)
        op.bind("<Button-1>", lambda e, dd=doc: seleccionar("docente", dd, f"👩‍🏫 {dd}"))
        tk.Frame(dropdown, bg="#FCE4EC", height=1).pack(fill="x", padx=10)
combo_container.bind("<Button-1>", toggle_dropdown); lbl_selected.bind("<Button-1>", toggle_dropdown); canvas_flecha.bind("<Button-1>", toggle_dropdown)
pass_frame=tk.Frame(caja, bg=BLANCO, highlightbackground=BORDE_ROSA, highlightthickness=1); pass_frame.pack(fill="x", pady=6, ipady=7)
entry_pass=tk.Entry(pass_frame, bd=0, font=("Arial",13,"bold"), show="•"); entry_pass.insert(0, "123"); entry_pass.pack(fill="x", padx=12)

def crear_sidebar(principal, botones, comando_salir):
    sidebar=tk.Frame(principal, bg=AZUL, width=260); sidebar.pack(side="left", fill="y"); sidebar.pack_propagate(False)
    tk.Label(sidebar, text="📚 ESCUELA PRIMARIA", bg=AZUL, fg="#94A3B8", font=("Arial",13,"bold")).pack(anchor="w", padx=20, pady=(20,12))
    for txt, cmd in botones:
        b=tk.Button(sidebar, text=f"  {txt}", bg=AZUL, fg=BLANCO, bd=0, anchor="w", padx=20, pady=13, font=("Arial",13,"bold"), activebackground="#334155", command=cmd); b.pack(fill="x", pady=1)
    tk.Button(sidebar, text="🚪 Cerrar Sesión", bg="#BE123C", fg=BLANCO, bd=0, font=FUENTE_BOTON, pady=11, command=comando_salir).pack(side="bottom", fill="x", padx=15, pady=15)
    return sidebar

def dialogo_acoplado(titulo, mensaje):
    win=tk.Toplevel(root); win.geometry("360x200"); win.configure(bg=BLANCO); win.transient(root); win.grab_set()
    tk.Label(win, text=titulo, bg=BLANCO, font=FUENTE_SUB, fg=ROSA).pack(pady=15)
    tk.Label(win, text=mensaje, bg=BLANCO, font=("Arial",13,"bold"), wraplength=300, justify="center").pack(pady=10)
    tk.Button(win, text="Aceptar", bg=ROSA, fg=BLANCO, bd=0, font=FUENTE_BOTON, pady=8, padx=22, command=win.destroy).pack(pady=15)

# ===== NUEVO: GLOSARIO PDF =====
def crear_vista_glosario(parent):
    for w in parent.winfo_children(): w.destroy()
    parent.configure(bg=FONDO)
    tk.Label(parent, text="📖 Glosario Inglés - Español", bg=FONDO, font=FUENTE_TITULO, fg=AZUL).pack(anchor="w", padx=25, pady=(15,2))
    tk.Label(parent, text="📄 Glosario oficial | Usa ◀ ▶ para pasar de página", bg=FONDO, font=FUENTE_CHICA, fg="#3A3A3A").pack(anchor="w", padx=25, pady=(0,8))
    if not os.path.exists(RUTA_GLOSARIO):
        tk.Label(parent, text="⚠️ No se encontró glosario.pdf\nPon el archivo junto a tu programa y renómbralo a glosario.pdf", bg=BLANCO, font=FUENTE_NORMAL, fg=ROJO, justify="center").pack(fill="both", expand=True, padx=25, pady=10)
        return
    doc = fitz.open(RUTA_GLOSARIO)
    estado = {"pagina": 0, "img": None}
    controles = tk.Frame(parent, bg=FONDO); controles.pack(fill="x", padx=25, pady=5)
    lbl_pag = tk.Label(controles, text="", bg=FONDO, font=FUENTE_BOTON, fg=AZUL); lbl_pag.pack(side="left", expand=True)
    marco_img = tk.Frame(parent, bg=BLANCO, highlightbackground="#E2E8F0", highlightthickness=2); marco_img.pack(fill="both", expand=True, padx=25, pady=5)
    lbl_img = tk.Label(marco_img, bg=BLANCO); lbl_img.pack(fill="both", expand=True)
    def mostrar_pagina():
        pagina = doc[estado["pagina"]]
        pix = pagina.get_pixmap(dpi=150)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail((900, 620))
        estado["img"] = ImageTk.PhotoImage(img)
        lbl_img.config(image=estado["img"])
        lbl_pag.config(text=f"Página {estado['pagina']+1} de {len(doc)}")
    def anterior():
        if estado["pagina"] > 0:
            estado["pagina"] -= 1; mostrar_pagina()
    def siguiente():
        if estado["pagina"] < len(doc)-1:
            estado["pagina"] += 1; mostrar_pagina()
    tk.Button(controles, text="◀ Anterior", bg=AZUL, fg=BLANCO, bd=0, font=FUENTE_BOTON, padx=14, pady=8, command=anterior).pack(side="left", padx=4)
    tk.Button(controles, text="Siguiente ▶", bg=ROSA, fg=BLANCO, bd=0, font=FUENTE_BOTON, padx=14, pady=8, command=siguiente).pack(side="left", padx=4)
    tk.Button(controles, text="📂 Abrir PDF", bg=VERDE, fg=BLANCO, bd=0, font=FUENTE_BOTON, padx=12, pady=8, command=lambda: os.startfile(RUTA_GLOSARIO) if os.name=="nt" else None).pack(side="right", padx=4)
    mostrar_pagina()

def crear_vista_pendientes(parent):
    for w in parent.winfo_children(): w.destroy()
    parent.configure(bg=FONDO)
    tk.Label(parent, text="📌 Pendientes - Dirección Primaria", bg=FONDO, font=FUENTE_TITULO, fg=AZUL).pack(anchor="w", padx=25, pady=(15,8))
    banner=tk.Frame(parent, bg=ROSA_CLARO, highlightbackground=BORDE_ROSA, highlightthickness=1, padx=12, pady=12); banner.pack(fill="x", padx=25, pady=8)
    c=tk.Canvas(banner, width=50, height=50, bg=ROSA_CLARO, highlightthickness=0); c.pack(side="left")
    c.create_polygon(25,3, 47,25, 25,47, 3,25, fill=ROSA, outline=ROSA); c.create_text(25,25, text="♦", fill=BLANCO, font=("Arial",14,"bold"))
    tk.Label(banner, text=f"👋 Bienvenido Director {nombre_seleccionado.get()}\n📅 {date.today()} - {len(get_pendientes())} pendientes en MySQL", bg=ROSA_CLARO, font=FUENTE_NORMAL, fg=AZUL, justify="left").pack(side="left", padx=15)
    contenedor_lista=tk.Frame(parent, bg=FONDO); contenedor_lista.pack(fill="both", expand=True, padx=25, pady=10)
    def refrescar_pendientes():
        for w in contenedor_lista.winfo_children(): w.destroy()
        for pid, p in get_pendientes():
            row=tk.Frame(contenedor_lista, bg=BLANCO, padx=12, pady=12, highlightbackground="#E2E8F0", highlightthickness=1); row.pack(fill="x", pady=4)
            tk.Label(row, text="♦", bg=BLANCO, fg=ROSA, font=("Arial",14,"bold")).pack(side="left")
            tk.Label(row, text=p, bg=BLANCO, font=FUENTE_NORMAL, fg=AZUL, anchor="w").pack(side="left", padx=10, fill="x", expand=True)
            tk.Button(row, text="✅ Hecho", bg=VERDE, fg=BLANCO, bd=0, font=("Arial",10,"bold"), command=lambda i=pid: [del_pendiente(i), refrescar_pendientes()]).pack(side="right", padx=3)
            tk.Button(row, text="🗑️", bg=ROJO, fg=BLANCO, bd=0, font=("Arial",11,"bold"), command=lambda i=pid: [del_pendiente(i), refrescar_pendientes()]).pack(side="right", padx=3)
    form=tk.Frame(parent, bg=BLANCO, padx=12, pady=10, highlightbackground=BORDE_ROSA, highlightthickness=1); form.pack(fill="x", padx=25, pady=10)
    tk.Label(form, text="➕ Nuevo:", bg=BLANCO, font=FUENTE_LABEL, fg=AZUL).pack(side="left")
    e_nuevo=tk.Entry(form, font=FUENTE_NORMAL, bd=1, relief="solid", width=30); e_nuevo.pack(side="left", padx=10, ipady=4)
    def agregar():
        if not e_nuevo.get().strip(): return
        add_pendiente(f"📌 {e_nuevo.get().strip()}"); e_nuevo.delete(0, tk.END); refrescar_pendientes()
    tk.Button(form, text="Agregar", bg=ROSA, fg=BLANCO, bd=0, font=FUENTE_BOTON, padx=12, pady=6, command=agregar).pack(side="left")
    refrescar_pendientes()

def crear_tabla_asistencia_con_colores(parent, grado_actual="Tercero", es_docente=False):
    for w in parent.winfo_children(): w.destroy()
    parent.configure(bg=FONDO)
    tk.Label(parent, text=f"📋 Asistencia Primaria - {grado_actual}° (XAMPP)", bg=FONDO, font=FUENTE_TITULO, fg=AZUL).pack(anchor="w", padx=25, pady=(15,2))
    tk.Label(parent, text=f"🎓 {grado_actual}° Primaria | 📅 {date.today()} | 👩‍🏫 {nombre_seleccionado.get()}", bg=FONDO, font=FUENTE_CHICA, fg="#3A3A3A").pack(anchor="w", padx=25, pady=(0,5))
    resumen=tk.Frame(parent, bg=FONDO); resumen.pack(fill="x", padx=25, pady=5)
    reglas=tk.Frame(parent, bg="#FEF9C3", highlightbackground=AMARILLO, highlightthickness=2, padx=12, pady=6); reglas.pack(fill="x", padx=25, pady=4)
    tk.Label(reglas, text="⚠️ REGLAS: 3 ⏰ Retardos = 1 ❌ Falta | +4 ❌ Faltas = 🚨 Sanción | Doble click para cambiar - SE GUARDA EN XAMPP", bg="#FEF9C3", font=("Arial",11,"bold"), fg="#92400E").pack()
    btns_arriba=tk.Frame(parent, bg=FONDO); btns_arriba.pack(fill="x", padx=25, pady=6)
    tabla_frame=tk.Frame(parent, bg=BLANCO, highlightbackground="#E2E8F0", highlightthickness=2); tabla_frame.pack(fill="both", expand=True, padx=25, pady=5)
    style=ttk.Style(); style.configure("Treeview", font=("Arial",13,"bold"), rowheight=38); style.configure("Treeview.Heading", font=("Arial",12,"bold"))
    cols=("No","Alumno","Grado","Estado","Alerta"); tabla=ttk.Treeview(tabla_frame, columns=cols, show="headings", height=10)
    for col, txt in zip(cols, ["No","Alumno","Grado","Estado","⚠️ Advertencia"]): tabla.heading(col, text=txt)
    tabla.column("No", width=40, anchor="center"); tabla.column("Alumno", width=180); tabla.column("Grado", width=90, anchor="center"); tabla.column("Estado", width=140, anchor="center"); tabla.column("Alerta", width=260, anchor="center")
    tabla.pack(fill="both", expand=True, padx=8, pady=8)
    tabla.tag_configure("ASISTENCIA", background="#DCFCE7", foreground="#065F46"); tabla.tag_configure("FALTA", background="#FEE2E2", foreground="#991B1B"); tabla.tag_configure("RETARDO", background="#FEF3C7", foreground="#92400E")
    def refrescar():
        datos=get_alumnos(grado_actual)
        for i in tabla.get_children(): tabla.delete(i)
        for idx, (nom, grupo, est) in enumerate(datos, 1):
            cont=get_contador(grado_actual, nom)
            alerta=""
            if cont["faltas"]>=5: alerta=f"🚨 SANCIÓN - {cont['faltas']} faltas"
            elif cont["faltas"]==4: alerta="⚠️ CITATORIO - 4 faltas"
            elif cont["retardos"]>0: alerta=f"⏰ {cont['retardos']}/3 retardos"
            tag="ASISTENCIA" if "Asistencia" in est else "FALTA" if "Falta" in est else "RETARDO"
            tabla.insert("", "end", values=(idx, nom, grupo, est, alerta), tags=(tag,))
        a=sum(1 for _,_,e in datos if "Asistencia" in e); f=sum(1 for _,_,e in datos if "Falta" in e); r=sum(1 for _,_,e in datos if "Retardo" in e)
        for w in resumen.winfo_children(): w.destroy()
        for color, txt, cant in [(VERDE,"✅ Asistencia",a),(ROJO,"❌ Falta",f),(AMARILLO,"⏰ Retardo",r)]:
            card=tk.Frame(resumen, bg=BLANCO, padx=12, pady=6, highlightbackground=color, highlightthickness=3); card.pack(side="left", padx=(0,8))
            tk.Label(card, text=f"{txt} {cant}", bg=BLANCO, font=("Arial",11,"bold"), fg=color).pack()
    refrescar()
    def cambiar(event=None):
        sel=tabla.selection()
        if not sel: messagebox.showwarning("Aviso","👆 Selecciona un alumno"); return
        vals=tabla.item(sel[0])["values"]; idx=int(vals[0])-1
        datos=get_alumnos(grado_actual); nom_alumno=datos[idx][0]; estado_actual=datos[idx][2]
        cont=get_contador(grado_actual, nom_alumno)
        if "Asistencia" in estado_actual:
            nuevo="❌ Falta"; cont["faltas"]+=1
            if cont["faltas"]==4: dialogo_acoplado("⚠️ ADVERTENCIA", f"{nom_alumno}\n4 faltas")
            elif cont["faltas"]>=5: dialogo_acoplado("🚨 SANCIÓN", f"{nom_alumno}\n{cont['faltas']} faltas\n¡CITATORIO PADRES!")
        elif "Falta" in estado_actual:
            nuevo="⏰ Retardo"
            if cont["faltas"]>0: cont["faltas"]-=1
            cont["retardos"]+=1
            if cont["retardos"]>=3:
                cont["retardos"]=0; cont["faltas"]+=1
                dialogo_acoplado("⏰ → ❌ AUTOMÁTICO", f"{nom_alumno}\n3 retardos = 1 falta"); nuevo="❌ Falta"
        else:
            nuevo="✅ Asistencia"
            if cont["retardos"]>0: cont["retardos"]-=1
        set_estado(nom_alumno, grado_actual, nuevo); set_contador(grado_actual, nom_alumno, cont["retardos"], cont["faltas"]); refrescar()
    tk.Button(btns_arriba, text="🚦 Cambiar: ✅ → ❌ → ⏰", bg="#1E293B", fg=BLANCO, bd=0, font=("Arial",13,"bold"), pady=10, padx=16, command=cambiar).pack(side="left", padx=3)
    tk.Button(btns_arriba, text="💾 Guardar ", bg=VERDE, fg=BLANCO, bd=0, font=FUENTE_BOTON, pady=10, padx=14, command=lambda: dialogo_acoplado("✅ Guardado en XAMPP", f"Lista {grado_actual}° guardada")).pack(side="left", padx=4)
    tk.Button(btns_arriba, text="↩️ Volver", bg=ROSA, fg=BLANCO, bd=0, font=FUENTE_BOTON, pady=10, padx=14, command=lambda: parent.cambiar_a_grados()).pack(side="left", padx=4)
    tabla.bind("<Double-1>", cambiar)

def crear_vista_conducta_docente(parent, nombre_docente_actual):
    for w in parent.winfo_children(): w.destroy()
    grado_doc=docente_grado_map.get(nombre_docente_actual, "General")
    tk.Label(parent, text=f"⚠️ Reporte Conducta - {grado_doc} Primaria", bg=FONDO, font=FUENTE_TITULO, fg=AZUL).pack(anchor="w", padx=25, pady=(15,5))
    form=tk.Frame(parent, bg=BLANCO, padx=15, pady=15, highlightbackground=BORDE_ROSA, highlightthickness=1); form.pack(fill="x", padx=25, pady=10)
    tk.Label(form, text="👦 Alumno:", bg=BLANCO, font=FUENTE_LABEL).grid(row=0, column=0, sticky="w")
    e_alumno=tk.Entry(form, font=FUENTE_NORMAL, bd=1, relief="solid", width=25); e_alumno.grid(row=0, column=1, padx=10, pady=5, ipady=5)
    tk.Label(form, text="⚠️ Motivo:", bg=BLANCO, font=FUENTE_LABEL).grid(row=1, column=0, sticky="w")
    e_incid=tk.Entry(form, font=FUENTE_NORMAL, bd=1, relief="solid", width=25); e_incid.grid(row=1, column=1, padx=10, pady=5, ipady=5)
    def registrar():
        if not e_alumno.get() or not e_incid.get(): messagebox.showwarning("Aviso","Llena todo"); return
        add_incidencia(e_alumno.get(), grado_doc, e_incid.get(), str(date.today()), nombre_docente_actual)
        dialogo_acoplado("✅ Registrado ", f"Reporte {grado_doc} guardado"); e_alumno.delete(0, tk.END); e_incid.delete(0, tk.END)
    tk.Button(form, text=f"✅ Registrar en {grado_doc}", bg=ROSA, fg=BLANCO, bd=0, font=FUENTE_BOTON, pady=10, padx=15, command=registrar).grid(row=2, column=0, columnspan=2, pady=15, sticky="ew")

def crear_vista_mis_reportes_docente(parent, nombre_docente_actual, es_director=False):
    for w in parent.winfo_children(): w.destroy()		
    titulo="📂 Reportes - Toda la Primaria" if es_director else f"📂 Mis Reportes - {docente_grado_map.get(nombre_docente_actual)} (XAMPP)"
    tk.Label(parent, text=titulo, bg=FONDO, font=FUENTE_TITULO, fg=AZUL).pack(anchor="w", padx=25, pady=(15,5))
    tabla_frame=tk.Frame(parent, bg=BLANCO, highlightbackground="#E2E8F0", highlightthickness=1); tabla_frame.pack(fill="both", expand=True, padx=25, pady=10)
    cols=("Alumno","Grado","Motivo","Fecha","Docente"); tabla=ttk.Treeview(tabla_frame, columns=cols, show="headings", height=12)
    for c in cols: tabla.heading(c, text=c); tabla.pack(fill="both", expand=True, padx=8, pady=8)
    def refrescar_tabla():
        for i in tabla.get_children(): tabla.delete(i)
        for alum, grupo, inc, fecha, docente in get_incidencias():
            if es_director or docente==nombre_docente_actual: tabla.insert("", "end", values=(alum, grupo, inc, fecha, docente))
    refrescar_tabla()
    btns=tk.Frame(parent, bg=FONDO, height=60); btns.pack(side="bottom", fill="x", padx=25, pady=10)
    tk.Button(btns, text="🔄 Actualizar", bg=AZUL, fg=BLANCO, bd=0, font=FUENTE_BOTON, pady=9, padx=12, command=refrescar_tabla).pack(side="left", padx=4)

def crear_vista_buscar_alumnos(parent, grado_filtro=None):
    for w in parent.winfo_children(): w.destroy()
    texto_grado=f" - {grado_filtro}° Primaria" if grado_filtro else " - Toda la Primaria"
    tk.Label(parent, text=f"🔍 Buscar Alumnos{texto_grado} ", bg=FONDO, font=FUENTE_TITULO, fg=AZUL).pack(anchor="w", padx=25, pady=(15,5))
    buscador_frame=tk.Frame(parent, bg=BLANCO, highlightbackground=BORDE_ROSA, highlightthickness=2); buscador_frame.pack(fill="x", padx=25, pady=8, ipady=6)
    tk.Label(buscador_frame, text="🔎", bg=BLANCO, font=("Arial",16)).pack(side="left", padx=10)
    entry_buscar=tk.Entry(buscador_frame, bd=0, font=FUENTE_NORMAL, fg=AZUL); entry_buscar.pack(side="left", fill="x", expand=True, padx=5)
    tabla_frame=tk.Frame(parent, bg=BLANCO, highlightbackground="#E2E8F0", highlightthickness=1); tabla_frame.pack(fill="both", expand=True, padx=25, pady=10)
    cols=("Alumno","Grado","Asistencia"); tabla=ttk.Treeview(tabla_frame, columns=cols, show="headings", height=15)
    for c in cols: tabla.heading(c, text=c); tabla.pack(fill="both", expand=True, padx=8, pady=8)
    todos=[]
    for grado in ([grado_filtro] if grado_filtro else GRADOS):
        for nom, grupo, est in get_alumnos(grado): todos.append((nom, grupo, est))
    def filtrar(*args):
        texto=entry_buscar.get().lower()
        for i in tabla.get_children(): tabla.delete(i)
        for nom, grupo, est in todos:
            if texto in nom.lower() or texto=="": tabla.insert("", "end", values=(nom, grupo, est))
    entry_buscar.bind("<KeyRelease>", filtrar); filtrar()

def crear_vista_gestion_docentes(parent):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent, text="👩‍🏫 Plantilla Docente - Primaria", bg=FONDO, font=FUENTE_TITULO, fg=AZUL).pack(anchor="w", padx=25, pady=(15,5))
    form=tk.Frame(parent, bg=BLANCO, padx=15, pady=12, highlightbackground=BORDE_ROSA, highlightthickness=1); form.pack(fill="x", padx=25, pady=8)
    tk.Label(form, text="👤 Nombre:", bg=BLANCO, font=FUENTE_LABEL, fg=AZUL).grid(row=0, column=0)
    entry_doc=tk.Entry(form, font=FUENTE_NORMAL, bd=1, relief="solid", width=22); entry_doc.grid(row=0, column=1, padx=10, ipady=5)
    tk.Label(form, text="🎓 Grado:", bg=BLANCO, font=FUENTE_LABEL, fg=AZUL).grid(row=0, column=2)
    combo_grado=ttk.Combobox(form, values=GRADOS, width=12, font=FUENTE_NORMAL); combo_grado.grid(row=0, column=3, padx=10); combo_grado.current(0)
    tabla_frame=tk.Frame(parent, bg=BLANCO, highlightbackground="#E2E8F0", highlightthickness=1); tabla_frame.pack(fill="both", expand=True, padx=25, pady=10)
    tabla=ttk.Treeview(tabla_frame, columns=("No","Docente","Grado"), show="headings", height=12); tabla.heading("No", text="No"); tabla.heading("Docente", text="Docente"); tabla.heading("Grado", text="Grado"); tabla.column("No", width=50, anchor="center"); tabla.pack(fill="both", expand=True, padx=8, pady=8)
    def refrescar():
        for i in tabla.get_children(): tabla.delete(i)
        for idx, (doc, grd) in enumerate(get_docentes().items(), 1): tabla.insert("", "end", values=(idx, doc, grd))
    def agregar():
        nombre=entry_doc.get().strip(); grado=combo_grado.get()
        if not nombre: return
        cur.execute("INSERT INTO docentes (nombre, grado) VALUES (%s,%s) ON DUPLICATE KEY UPDATE grado=VALUES(grado)", (nombre, grado)); con.commit()
        docente_grado_map[nombre]=grado
        if nombre not in lista_docentes_global: lista_docentes_global.append(nombre)
        entry_doc.delete(0, tk.END); refrescar(); dialogo_acoplado("✅ Agregado", f"{nombre} a {grado}°")
    def quitar():
        sel=tabla.selection()
        if not sel: return
        vals=tabla.item(sel[0])["values"]; doc=vals[1]
        cur.execute("DELETE FROM docentes WHERE nombre=%s", (doc,)); con.commit()
        if doc in lista_docentes_global: lista_docentes_global.remove(doc)
        docente_grado_map.pop(doc, None); refrescar(); dialogo_acoplado("🗑️ Eliminado", f"{doc} eliminado")
    btns=tk.Frame(parent, bg=FONDO, height=60); btns.pack(side="bottom", fill="x", padx=25, pady=10)
    tk.Button(btns, text="➕ Agregar", bg=VERDE, fg=BLANCO, bd=0, font=FUENTE_BOTON, pady=10, padx=14, command=agregar).pack(side="left", padx=4)
    tk.Button(btns, text="🗑️ Quitar", bg=ROJO, fg=BLANCO, bd=0, font=FUENTE_BOTON, pady=10, padx=14, command=quitar).pack(side="left", padx=4)
    refrescar()

def crear_tarjeta_con_diseno(parent_grid, grado, indice, comando):
    color=COLORES_GRADOS[indice]; contenedor=tk.Frame(parent_grid, bg=FONDO)
    card=tk.Frame(contenedor, bg=BLANCO, bd=1, relief="solid", width=175, height=115); card.pack(padx=6, pady=6); card.pack_propagate(False)
    header=tk.Frame(card, bg=color, height=36); header.pack(fill="x"); tk.Label(header, text=f"{grado}° Primaria", bg=color, fg=BLANCO, font=("Arial",11,"bold")).pack(expand=True, pady=6)
    body=tk.Frame(card, bg=BLANCO); body.pack(fill="both", expand=True)
    tk.Label(body, text=f"👦 {len(get_alumnos(grado))} alumnos", bg=BLANCO, fg=AZUL, font=("Arial",10,"bold")).pack(pady=4)
    tk.Button(card, text="📋 Ver Lista", bg=GUINDA_BOTON, fg=BLANCO, bd=0, font=("Arial",10,"bold"), pady=5, command=comando).pack(fill="x", side="bottom")
    return contenedor

def mostrar_director():
    login_frame.pack_forget(); docente_frame.pack_forget(); director_frame.pack(fill="both", expand=True)
    for w in director_frame.winfo_children(): w.destroy()
    header=tk.Frame(director_frame, bg=ROSA, height=50); header.pack(fill="x")
    tk.Label(header, text=f"🏫 DIRECTOR PRIMARIA: {nombre_seleccionado.get()}", bg=ROSA, fg=BLANCO, font=FUENTE_SUB).pack(side="left", padx=20, pady=12)
    principal=tk.Frame(director_frame, bg=BLANCO); principal.pack(fill="both", expand=True)
    area=tk.Frame(principal, bg=FONDO); area.pack(side="left", fill="both", expand=True)
    vistas={}
    f_dash=tk.Frame(area, bg=FONDO); f_grados=tk.Frame(area, bg=FONDO); tk.Label(f_grados, text="🎓 6 Grados de Primaria", bg=FONDO, font=FUENTE_TITULO, fg=AZUL).pack(anchor="w", padx=25, pady=(15,8))
    grid_frame=tk.Frame(f_grados, bg=FONDO); grid_frame.pack(fill="both", expand=True, padx=15, pady=8)
    def ver_grado_dir(grado_nombre):
        if grado_nombre in vistas and grado_nombre not in ["CP1 Pendientes", "Grados"]: cambiar(grado_nombre); return
        f_nuevo=tk.Frame(area, bg=FONDO); f_nuevo.cambiar_a_grados=lambda: cambiar("Grados")
        crear_tabla_asistencia_con_colores(f_nuevo, grado_actual=grado_nombre, es_docente=False); vistas[grado_nombre]=f_nuevo; cambiar(grado_nombre)
    for i, g in enumerate(GRADOS): crear_tarjeta_con_diseno(grid_frame, g, i, lambda gg=g: ver_grado_dir(gg)).grid(row=i//3, column=i%3, padx=6, pady=6)
    f_conducta=tk.Frame(area, bg=FONDO); f_mis=tk.Frame(area, bg=FONDO); f_buscar=tk.Frame(area, bg=FONDO); f_docentes=tk.Frame(area, bg=FONDO); f_glosario=tk.Frame(area, bg=FONDO)
    vistas["CP1 Pendientes"]=f_dash; vistas["Grados"]=f_grados; vistas["Conducta"]=f_conducta; vistas["Mis Reportes"]=f_mis; vistas["Buscar Alumnos"]=f_buscar; vistas["Docentes"]=f_docentes; vistas["Glosario"]=f_glosario
    def cambiar(n):
        for v in vistas.values(): v.pack_forget()
        if n=="CP1 Pendientes": crear_vista_pendientes(f_dash)
        elif n=="Conducta": crear_vista_conducta_docente(f_conducta, nombre_seleccionado.get())
        elif n=="Mis Reportes": crear_vista_mis_reportes_docente(f_mis, nombre_seleccionado.get(), True)
        elif n=="Buscar Alumnos": crear_vista_buscar_alumnos(f_buscar, None)
        elif n=="Docentes": crear_vista_gestion_docentes(f_docentes)
        elif n=="Glosario": crear_vista_glosario(f_glosario)
        vistas[n].pack(fill="both", expand=True)
    def salir(): director_frame.pack_forget(); login_frame.pack(fill="both", expand=True)
    crear_sidebar(principal, [("📌 Pendientes", lambda: cambiar("CP1 Pendientes")), ("🎓 6 Grados", lambda: cambiar("Grados")), ("⚠️ Conducta", lambda: cambiar("Conducta")), ("📂 Reportes", lambda: cambiar("Mis Reportes")), ("🔍 Buscar", lambda: cambiar("Buscar Alumnos")), ("👩‍🏫 Docentes", lambda: cambiar("Docentes")), ("📖 Glosario", lambda: cambiar("Glosario"))], salir)
    cambiar("CP1 Pendientes")

def mostrar_docente():
    login_frame.pack_forget(); director_frame.pack_forget(); docente_frame.pack(fill="both", expand=True)
    for w in docente_frame.winfo_children(): w.destroy()
    grado_asignado=docente_grado_map.get(nombre_seleccionado.get(), "Sin grado")
    header=tk.Frame(docente_frame, bg=ROSA, height=50); header.pack(fill="x")
    tk.Label(header, text=f"🏫 DOCENTE {grado_asignado}° PRIMARIA: {nombre_seleccionado.get()}", bg=ROSA, fg=BLANCO, font=FUENTE_SUB).pack(side="left", padx=20, pady=12)
    principal=tk.Frame(docente_frame, bg=BLANCO); principal.pack(fill="both", expand=True)
    area=tk.Frame(principal, bg=FONDO); area.pack(side="left", fill="both", expand=True)
    f_cond=tk.Frame(area, bg=FONDO); f_mis=tk.Frame(area, bg=FONDO)
    f_grados_doc=tk.Frame(area, bg=FONDO); tk.Label(f_grados_doc, text=f"🎓 Mi Grupo: {grado_asignado}° Primaria", bg=FONDO, font=FUENTE_TITULO, fg=AZUL).pack(anchor="w", padx=25, pady=(15,8))
    grid_frame_doc=tk.Frame(f_grados_doc, bg=FONDO); grid_frame_doc.pack(fill="both", expand=True, padx=15, pady=8)
    vistas_doc={"Conducta": f_cond, "Mis Reportes": f_mis, "Grados": f_grados_doc}
    f_buscar_doc=tk.Frame(area, bg=FONDO); vistas_doc["Buscar Alumnos"]=f_buscar_doc
    f_glosario_doc=tk.Frame(area, bg=FONDO); vistas_doc["Glosario"]=f_glosario_doc
    f_mi_grado_detalle=tk.Frame(area, bg=FONDO); f_mi_grado_detalle.cambiar_a_grados=lambda: cambiar_doc("Grados")
    crear_tabla_asistencia_con_colores(f_mi_grado_detalle, grado_actual=grado_asignado, es_docente=True)
    vistas_doc[grado_asignado]=f_mi_grado_detalle
    crear_tarjeta_con_diseno(grid_frame_doc, grado_asignado, GRADOS.index(grado_asignado) if grado_asignado in GRADOS else 0, lambda: cambiar_doc(grado_asignado)).grid(row=0, column=0, padx=6, pady=6)
    def cambiar_doc(n):
        for v in vistas_doc.values(): v.pack_forget()
        vistas_doc[n].pack(fill="both", expand=True)
        if n=="Conducta": crear_vista_conducta_docente(f_cond, nombre_seleccionado.get())
        if n=="Mis Reportes": crear_vista_mis_reportes_docente(f_mis, nombre_seleccionado.get(), False)
        if n=="Buscar Alumnos": crear_vista_buscar_alumnos(f_buscar_doc, grado_filtro=grado_asignado)
        if n=="Glosario": crear_vista_glosario(f_glosario_doc)
    def salir_doc(): docente_frame.pack_forget(); login_frame.pack(fill="both", expand=True)
    crear_sidebar(principal, [(f"🎓 Mi Grupo: {grado_asignado}°", lambda: cambiar_doc("Grados")), ("⚠️ Conducta", lambda: cambiar_doc("Conducta")), ("📂 Mis Reportes", lambda: cambiar_doc("Mis Reportes")), (f"🔍 Alumnos {grado_asignado}°", lambda: cambiar_doc("Buscar Alumnos")), ("📖 Glosario", lambda: cambiar_doc("Glosario"))], salir_doc)
    cambiar_doc("Grados")

def login():
    if usuario_real.get() not in ["director", "docente"]: messagebox.showwarning("Aviso", "Selecciona usuario ♦"); return
    if usuario_real.get()=="director": mostrar_director()
    else: mostrar_docente()

tk.Button(caja, text="🚀 Ingresar Primaria", bg=ROSA, fg=BLANCO, bd=0, font=("Arial",14,"bold"), pady=12, command=login).pack(fill="x", pady=(12,5))
root.mainloop()