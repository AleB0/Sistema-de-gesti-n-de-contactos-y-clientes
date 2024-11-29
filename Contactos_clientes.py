import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
import ctypes
import pyodbc
#-------------------------------------------------------------------------------------------------
OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / "img"

def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)
#-------------------------------------------------------------------------------------------------

# Crear ventana principal
root = tk.Tk()
root.title("Gestión de Contactos y Clientes")
root.geometry("800x600")

root.iconbitmap(os.path.join(os.path.dirname(__file__), 'img/libreta.ico'))                             
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("MiAplicacionID")

# Establecer la conexión con la base de datos de Access
conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=Contactos_clientes.accdb')
cursor = conn.cursor()

# Función para cargar los contactos desde la base de datos al iniciar
def load_contacts_from_db():
    cursor.execute("SELECT * FROM Contactos_clientes")
    rows = cursor.fetchall()
    for row in rows:
        tree.insert("", "end", values=row)

# Frame para la barra lateral
frame_sidebar = tk.Frame(root, bg="#2E3B4E", width=200)
frame_sidebar.pack(fill="y", side="left")

# Títulos de la barra lateral
label_sidebar_title = tk.Label(frame_sidebar, text="Contactos y clientes", font=("Arial", 14, "bold"), fg="white", bg="#2E3B4E")
label_sidebar_title.pack(pady=10)

# Crear funciones para cada botón
def show_all():
    tree.delete(*tree.get_children())
    load_contacts_from_db()
    messagebox.showinfo("Inicio", "Mostrando todos los contactos.")

def filter_clients():
    tree.delete(*tree.get_children())
    cursor.execute("SELECT * FROM Contactos_clientes WHERE Categoria = 'Cliente'")
    rows = cursor.fetchall()
    for row in rows:
        tree.insert("", "end", values=row)
    messagebox.showinfo("Clientes", "Mostrando solo clientes.")

def filter_contacts():
    tree.delete(*tree.get_children())
    cursor.execute("SELECT * FROM Contactos_clientes WHERE Categoria != 'Cliente'")
    rows = cursor.fetchall()
    for row in rows:
        tree.insert("", "end", values=row)
    messagebox.showinfo("Contactos", "Mostrando solo contactos que no son clientes.")

def add_new_contact():
    add_window = tk.Toplevel(root)
    add_window.title("Nuevo Contacto")

    # Campos de entrada para nuevo contacto
    tk.Label(add_window, text="ID").grid(row=0, column=0)
    tk.Label(add_window, text="Nombre").grid(row=1, column=0)
    tk.Label(add_window, text="Teléfono").grid(row=2, column=0)
    tk.Label(add_window, text="Email").grid(row=3, column=0)
    tk.Label(add_window, text="Categoría").grid(row=4, column=0)

    id_entry = tk.Entry(add_window)
    id_entry.grid(row=0, column=1)
    name_entry = tk.Entry(add_window)
    name_entry.grid(row=1, column=1)
    phone_entry = tk.Entry(add_window)
    phone_entry.grid(row=2, column=1)
    email_entry = tk.Entry(add_window)
    email_entry.grid(row=3, column=1)
    category_entry = tk.Entry(add_window)
    category_entry.grid(row=4, column=1)

    # Función para guardar el nuevo contacto en la base de datos
    def save_new_contact():
        new_contact = (
            int(id_entry.get()),
            name_entry.get(),
            phone_entry.get(),
            email_entry.get(),
            category_entry.get()
        )
        
        # Insertar en la base de datos
        cursor.execute("""
            INSERT INTO Contactos_clientes (ID, Nombre, Telefono, Email, Categoria)
            VALUES (?, ?, ?, ?, ?)
        """, new_contact)
        conn.commit()

        # Insertar en el Treeview 
        tree.insert("", "end", values=new_contact)
        add_window.destroy()
    
    tk.Button(add_window, text="Guardar", command=save_new_contact).grid(row=5, column=1)

# Botones de la barra lateral
btn_inicio = tk.Button(frame_sidebar, text="Inicio", bg="#2E3B4E", fg="white", font=("Arial", 12), bd=0, anchor="w", command=show_all)
btn_inicio.pack(fill="x", padx=10, pady=5)

btn_clientes = tk.Button(frame_sidebar, text="Clientes", bg="#2E3B4E", fg="white", font=("Arial", 12), bd=0, anchor="w", command=filter_clients)
btn_clientes.pack(fill="x", padx=10, pady=5)

btn_contactos = tk.Button(frame_sidebar, text="Contactos", bg="#2E3B4E", fg="white", font=("Arial", 12), bd=0, anchor="w", command=filter_contacts)
btn_contactos.pack(fill="x", padx=10, pady=5)

# Frame para la barra superior
frame_top = tk.Frame(root, bg="#4C9BE8", height=50)
frame_top.pack(fill="x", side="top")

# Botones en la barra superior
btn_nuevo = tk.Button(frame_top, text="Nuevo", bg="#4C9BE8", fg="white", font=("Arial", 12), command=add_new_contact)
btn_nuevo.pack(side="left", padx=10)

# Frame para la tabla
frame_table = tk.Frame(root)
frame_table.pack(fill="both", expand=True, padx=10, pady=10)

# Crear tabla con Treeview
columns = ("ID", "Nombre", "Teléfono", "Email", "Categoría")
tree = ttk.Treeview(frame_table, columns=columns, show="headings")
tree.pack(fill="both", expand=True)

# Configurar encabezados
for col in columns:
    tree.heading(col, text=col)

# Cargar los contactos desde la base de datos
load_contacts_from_db()

# Funciones de edición y eliminación para las filas seleccionadas
def edit_action():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Advertencia", "Por favor, selecciona un contacto para editar.")
        return
    
    item = tree.item(selected_item)
    values = item['values']
    
    edit_window = tk.Toplevel(root)
    edit_window.title("Editar Contacto")

    tk.Label(edit_window, text="ID").grid(row=0, column=0)
    tk.Label(edit_window, text="Nombre").grid(row=1, column=0)
    tk.Label(edit_window, text="Teléfono").grid(row=2, column=0)
    tk.Label(edit_window, text="Email").grid(row=3, column=0)
    tk.Label(edit_window, text="Categoría").grid(row=4, column=0)

    id_entry = tk.Entry(edit_window)
    id_entry.grid(row=0, column=1)
    id_entry.insert(0, values[0])
    id_entry.config(state="disabled")

    name_entry = tk.Entry(edit_window)
    name_entry.grid(row=1, column=1)
    name_entry.insert(0, values[1])

    phone_entry = tk.Entry(edit_window)
    phone_entry.grid(row=2, column=1)
    phone_entry.insert(0, values[2])

    email_entry = tk.Entry(edit_window)
    email_entry.grid(row=3, column=1)
    email_entry.insert(0, values[3])

    category_entry = tk.Entry(edit_window)
    category_entry.grid(row=4, column=1)
    category_entry.insert(0, values[4])

    def save_changes():
        updated_values = (
            values[0],  # Mantener el ID
            name_entry.get(),
            phone_entry.get(),
            email_entry.get(),
            category_entry.get()
        )

        # Actualizar en la base de datos
        cursor.execute("""
            UPDATE Contactos_clientes SET Nombre = ?, Telefono = ?, Email = ?, Categoria = ?
            WHERE ID = ?
        """, (updated_values[1], updated_values[2], updated_values[3], updated_values[4], updated_values[0]))
        conn.commit()

        # Actualizar en el Treeview
        tree.item(selected_item, values=updated_values)
        edit_window.destroy()

    tk.Button(edit_window, text="Guardar Cambios", command=save_changes).grid(row=5, column=1)

def delete_action():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Advertencia", "Por favor, selecciona un contacto para eliminar.")
        return
    
    item = tree.item(selected_item)
    contact_id = item['values'][0]

    response = messagebox.askyesno("Confirmar eliminación", f"¿Estás seguro de que quieres eliminar el contacto con ID {contact_id}?")
    if response:
        cursor.execute("DELETE FROM Contactos_clientes WHERE ID = ?", contact_id)
        conn.commit()
        tree.delete(selected_item)

# Iniciar la interfaz gráfica
root.mainloop()
