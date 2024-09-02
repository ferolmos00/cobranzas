# cobranzas
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pywhatkit as kit
import time
import logging
import os

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def enviar_correo(smtp_server, smtp_port, smtp_user, smtp_password, email, mensaje):
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = email
    msg['Subject'] = 'Recordatorio de Pago'
    msg.attach(MIMEText(mensaje, 'plain'))

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.sendmail(smtp_user, email, msg.as_string())
        logging.info(f"Correo enviado a {email}")
        messagebox.showinfo("Éxito", f"Correo enviado a {email}")
        return True
    except smtplib.SMTPAuthenticationError:
        logging.error("Error de Autenticación: Las credenciales de usuario son incorrectas.")
        messagebox.showerror("Error de Autenticación", "Las credenciales de usuario son incorrectas. Por favor, verifica tu correo y contraseña.")
    except smtplib.SMTPException as e:
        logging.error(f"Error de SMTP: No se pudo enviar el correo a {email}: {str(e)}")
        messagebox.showerror("Error de SMTP", f"No se pudo enviar el correo a {email}: {str(e)}")
    except Exception as e:
        logging.error(f"Error inesperado: {str(e)}")
        messagebox.showerror("Error", f"Ha ocurrido un error inesperado: {str(e)}")
    return False

def enviar_correos_y_whatsapp():
    try:
        file_path = entry_excel.get()
        df = pd.read_excel(file_path)

        df['Cliente'] = df['Cliente'].astype(str).str.strip()
        df['Saldo'] = df['Saldo'].astype(str).str.strip()
        df['Email'] = df['Email'].astype(str).str.strip()
        df['WhatsApp'] = "+" + df['WhatsApp'].astype(str).str.strip()

        if df.isnull().sum().any():
            messagebox.showerror("Error", "Hay datos faltantes en el archivo Excel.")
            return

        smtp_server = entry_smtp_server.get()
        smtp_port = int(entry_smtp_port.get())
        smtp_user = entry_smtp_user.get()
        smtp_password = entry_smtp_password.get()

        if not smtp_server or not smtp_port or not smtp_user or not smtp_password:
            messagebox.showerror("Error", "Por favor, completa todos los campos de configuración del correo.")
            return

        progress_bar['maximum'] = len(df)
        progress_bar['value'] = 0

        for index, row in df.iterrows():
            cliente = row['Cliente']
            saldo = row['Saldo']
            email = row['Email']
            whatsapp = row['WhatsApp']

            mensaje = f'Hola {cliente}, tu saldo adeudado en Electricidad Luján es de ${saldo}. Por favor, realiza el pago a la brevedad. Gracias.'

            if email and "@" in email:
                if enviar_correo(smtp_server, smtp_port, smtp_user, smtp_password, email, mensaje):
                    tree.insert("", "end", values=(cliente, saldo, email, "Correo Enviado"))
                else:
                    tree.insert("", "end", values=(cliente, saldo, email, "Error al enviar"))
            else:
                tree.insert("", "end", values=(cliente, saldo, "SIN EMAIL", "No enviado"))

            if whatsapp and whatsapp.startswith("+"):
                try:
                    # Enviar mensaje de WhatsApp usando pywhatkit
                    hora = time.localtime().tm_hour
                    minuto = time.localtime().tm_min + 2  # Enviar el mensaje 2 minutos después de la hora actual
                    if minuto >= 60:
                        hora += 1
                        minuto -= 60
                    kit.sendwhatmsg(whatsapp, mensaje, hora, minuto, 10, True, 2)
                    tree.insert("", "end", values=(cliente, saldo, whatsapp, "WhatsApp Enviado"))
                except Exception as e:
                    logging.error(f"Error al enviar WhatsApp a {whatsapp}: {str(e)}")
                    tree.insert("", "end", values=(cliente, saldo, whatsapp, f"Error WhatsApp: {str(e)}"))
            else:
                tree.insert("", "end", values=(cliente, saldo, "SIN WHATSAPP", "No enviado"))

            progress_bar['value'] += 1
            root.update_idletasks()

        messagebox.showinfo("Éxito", "Todos los mensajes han sido enviados.")
    except Exception as e:
        logging.error(f"Error inesperado: {str(e)}")
        messagebox.showerror("Error", str(e))

def seleccionar_archivo():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_excel.delete(0, tk.END)
        entry_excel.insert(0, file_path)
        mostrar_contenido_excel(file_path)

def mostrar_contenido_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        for item in tree.get_children():
            tree.delete(item)
        for index, row in df.iterrows():
            tree.insert("", "end", values=(row['Cliente'], row['Saldo'], row['Email'], row['WhatsApp']))
    except Exception as e:
        logging.error(f"No se pudo leer el archivo Excel: {str(e)}")
        messagebox.showerror("Error", f"No se pudo leer el archivo Excel: {str(e)}")

# Configuración de la interfaz gráfica
root = tk.Tk()
root.title("Envío de Correos y WhatsApp")
root.geometry("1200x1000")

style = ttk.Style()
style.theme_use("clam")

frame_config = ttk.LabelFrame(root, text="Configuración del Correo")
frame_config.pack(pady=10, padx=10, fill="x")

ttk.Label(frame_config, text="Servidor SMTP:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
entry_smtp_server = ttk.Entry(frame_config)
entry_smtp_server.grid(row=0, column=1, padx=5, pady=5, sticky="w")
entry_smtp_server.insert(0, 'smtp-mail.outlook.com')

ttk.Label(frame_config, text="Puerto SMTP:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
entry_smtp_port = ttk.Entry(frame_config)
entry_smtp_port.grid(row=1, column=1, padx=5, pady=5, sticky="w")
entry_smtp_port.insert(0, '587')

ttk.Label(frame_config, text="Usuario SMTP:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
entry_smtp_user = ttk.Entry(frame_config)
entry_smtp_user.grid(row=2, column=1, padx=5, pady=5, sticky="w")
entry_smtp_user.insert(0, 'matidole@hotmail.com')

ttk.Label(frame_config, text="Contraseña SMTP:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
entry_smtp_password = ttk.Entry(frame_config, show="*")
entry_smtp_password.grid(row=3, column=1, padx=5, pady=5, sticky="w")
entry_smtp_password.insert(0, 'Piliema00')

frame_excel = ttk.LabelFrame(root, text="Archivo Excel")
frame_excel.pack(pady=10, padx=10, fill="x")

ttk.Label(frame_excel, text="Ruta del archivo Excel:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
entry_excel = ttk.Entry(frame_excel, width=40)
entry_excel.grid(row=0, column=1, padx=5, pady=5, sticky="w")
btn_seleccionar = ttk.Button(frame_excel, text="Seleccionar", command=seleccionar_archivo)
btn_seleccionar.grid(row=0, column=2, padx=5, pady=5)

btn_enviar = ttk.Button(root, text="Enviar Correos y WhatsApp", command=enviar_correos_y_whatsapp)
btn_enviar.pack(pady=10)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.pack(pady=10)

tree = ttk.Treeview(root, columns=("Cliente", "Saldo", "Contacto", "Estado"), show="headings")
tree.heading("Cliente", text="Cliente")
tree.heading("Saldo", text="Saldo")
tree.heading("Contacto", text="Contacto")
tree.heading("Estado", text="Estado")
tree.pack(pady=10, fill="both", expand=True)

root.mainloop()
