import os
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox

import fitz
from docx import Document
from docx.shared import Inches


def pdf_a_word(pdf_path, output_docx):
    doc = Document()

    with tempfile.TemporaryDirectory() as temp_dir:
        pdf = fitz.open(pdf_path)

        for i, pagina in enumerate(pdf):
            pixmap = pagina.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            nombre_imagen = os.path.join(temp_dir, f"temp_page_{i}.png")
            pixmap.save(nombre_imagen)
            doc.add_picture(nombre_imagen, width=Inches(6.0))
            if i < len(pdf) - 1:
                doc.add_page_break()

    doc.save(output_docx)
    return output_docx


def seleccionar_pdf_y_exportar():
    root = tk.Tk()
    root.withdraw()
    root.update()

    pdf_path = filedialog.askopenfilename(
        title="Selecciona un PDF",
        filetypes=[("Archivos PDF", "*.pdf")],
    )
    if not pdf_path:
        return

    output_docx = filedialog.asksaveasfilename(
        title="Guardar como Word",
        defaultextension=".docx",
        filetypes=[("Documento Word", "*.docx")],
        initialfile=os.path.splitext(os.path.basename(pdf_path))[0] + ".docx",
    )
    if not output_docx:
        return

    try:
        pdf_a_word(pdf_path, output_docx)
        messagebox.showinfo("Listo", f"Documento guardado como:\n{output_docx}")
    except Exception as error:
        messagebox.showerror("Error", f"No se pudo convertir el PDF.\n\n{error}")
    finally:
        root.destroy()


if __name__ == "__main__":
    seleccionar_pdf_y_exportar()
