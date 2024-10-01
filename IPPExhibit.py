import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ctypes
from pypdf import PdfReader, PdfWriter
import pdfplumber
import os
import fitz
import sv_ttk
from tkinterdnd2 import TkinterDnD, DND_FILES

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception as e:
    print(f"Could not set DPI awareness: {e}")

def reset_eof_of_pdf_return_stream(pdf_stream_in):
    actual_line = None
    for i, x in enumerate(pdf_stream_in[::-1]):
        if b'%%EOF' in x:
            actual_line = len(pdf_stream_in) - i
            print(f'EOF found at line position {-i} = actual {actual_line}, with value {x}')
            break
    if actual_line is not None:
        return pdf_stream_in[:actual_line]
    else:
        return pdf_stream_in

def clean_pdf(input_path, output_path):
    try:
        document = fitz.open(input_path)
        document.save(output_path, garbage=4, deflate=True, clean=True)
        document.close()
    except Exception as e:
        print(f"Failed to clean PDF {input_path}: {e}")
        try:
            with open(input_path, 'rb') as p:
                txt = p.readlines()
            txtx = reset_eof_of_pdf_return_stream(txt)
            repaired_path = output_path
            with open(repaired_path, 'wb') as f:
                f.writelines(txtx)
            try:
                fixed_pdf = PdfReader(repaired_path)
                print(f"Repaired PDF {input_path} saved as {repaired_path}")
                return True
            except Exception as e:
                print(f"Failed to read repaired PDF {repaired_path}: {e}")
                return False
        except Exception as repair_error:
            print(f"Failed to repair PDF {input_path}: {repair_error}")
            return False
    return True

def find_exhibit_pages(pdf_path):
    exhibit_pages = {}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    for line in lines:
                        normalized_line = line.strip().upper()
                        if normalized_line.startswith("EXHIBIT "):
                            exhibit_letter = normalized_line.split(" ")[1].strip(":")
                            exhibit_pages[exhibit_letter] = i
                            break
    except Exception as e:
        raise RuntimeError(f"Failed to find exhibit pages in {pdf_path}: {e}")
    return exhibit_pages

def merge_pdfs(main_pdf_path, exhibit_paths, output_path):
    try:
        cleaned_main_pdf_path = "cleaned_main.pdf"
        cleaned_exhibit_paths = {k: [f"cleaned_{os.path.basename(path)}" for path in paths] for k, paths in exhibit_paths.items()}
        
        clean_pdf(main_pdf_path, cleaned_main_pdf_path)
        for exhibit_letter, paths in exhibit_paths.items():
            for path in paths:
                cleaned_path = cleaned_exhibit_paths[exhibit_letter][paths.index(path)]
                clean_pdf(path, cleaned_path)

        exhibit_pages = find_exhibit_pages(cleaned_main_pdf_path)
        main_reader = PdfReader(cleaned_main_pdf_path)
        writer = PdfWriter()

        for i, page in enumerate(main_reader.pages):
            writer.add_page(page)
            for exhibit_letter, page_num in exhibit_pages.items():
                if i == page_num:
                    exhibit_path_list = exhibit_paths.get(exhibit_letter)
                    if exhibit_path_list:
                        for exhibit_path in exhibit_path_list:
                            try:
                                exhibit_reader = PdfReader(exhibit_path)
                                for exhibit_page in exhibit_reader.pages:
                                    writer.add_page(exhibit_page)
                                print(f"Inserted {exhibit_letter} from {exhibit_path} at page {i + 1}")
                            except Exception as e:
                                raise RuntimeError(f"Failed to read exhibit PDF {exhibit_path}: {e}")

        with open(output_path, "wb") as output_file:
            writer.write(output_file)
        messagebox.showinfo("Success", f"Merged PDF saved as {output_path}")
        print(f"Merged PDF saved as {output_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))
        print(f"Error: {e}")
def manage_exhibit(event):
    selection = exhibits_listbox.curselection()
    if selection:
        index = selection[0]
        exhibit_letter = exhibits_listbox.get(index).split(":")[0].strip()
        if exhibit_letter in exhibit_paths:
            files = exhibit_paths[exhibit_letter]
            manage_window = tk.Toplevel(root)
            manage_window.iconbitmap('C:/Users/omars_qgwtbpx/Documents/ilprisonproject_logo.ico')
            manage_window.title(f"Manage Exhibit {exhibit_letter}")
            manage_window.geometry("340x690")
            manage_frame = ttk.Frame(manage_window, padding="10 10 10 10")
            manage_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            manage_listbox = tk.Listbox(manage_frame, width=30, height=20)
            manage_listbox.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky=(tk.W, tk.E))
            for file in files:
                manage_listbox.insert(tk.END, os.path.basename(file))

            def add_files():
                new_files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
                for new_file in new_files:
                    files.append(new_file)
                    manage_listbox.insert(tk.END, os.path.basename(new_file))

            def remove_selected():
                selected_indices = manage_listbox.curselection()
                for index in reversed(selected_indices):
                    manage_listbox.delete(index)
                    del files[index]

            def apply_changes():
                exhibit_paths[exhibit_letter] = files
                exhibits_listbox.delete(index)
                exhibits_listbox.insert(index, f"{exhibit_letter}: {', '.join([os.path.basename(f) for f in files])}")
                manage_window.destroy()

            def on_drag_start(event):
                widget = event.widget
                start_index = widget.nearest(event.y)
                widget.selection_set(start_index)
                widget.drag_data = (start_index, widget.get(start_index))

            def on_drag_motion(event):
                widget = event.widget
                widget.selection_clear(0, tk.END)
                current_index = widget.nearest(event.y)
                widget.selection_set(current_index)

            def on_drag_drop(event):
                widget = event.widget
                start_index, text = widget.drag_data
                end_index = widget.nearest(event.y)
                if start_index != end_index:
                    widget.delete(start_index)
                    widget.insert(end_index, text)
                    widget.selection_set(end_index)
                    files.insert(end_index, files.pop(start_index))

            def drop(event):
                dropped_files = event.data.split()
                for file in dropped_files:
                    if file.lower().endswith('.pdf'):
                        files.append(file)
                        manage_listbox.insert(tk.END, os.path.basename(file))

            def move_item(direction):
                selected_indices = manage_listbox.curselection()
                if not selected_indices:
                    return
                for index in selected_indices:
                    new_index = index + direction
                    if 0 <= new_index < manage_listbox.size():
                        item_text = manage_listbox.get(index)
                        manage_listbox.delete(index)
                        manage_listbox.insert(new_index, item_text)
                        manage_listbox.selection_set(new_index)
                        files.insert(new_index, files.pop(index))

            manage_listbox.drop_target_register(DND_FILES)
            manage_listbox.dnd_bind('<<Drop>>', drop)
            manage_listbox.bind('<Button-1>', on_drag_start)
            manage_listbox.bind('<B1-Motion>', on_drag_motion)
            manage_listbox.bind('<ButtonRelease-1>', on_drag_drop)
            up_button = ttk.Button(manage_frame, text="Up", command=lambda: move_item(-1))
            up_button.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
            down_button = ttk.Button(manage_frame, text="Down", command=lambda: move_item(1))
            down_button.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
            add_button = ttk.Button(manage_frame, text="Add Files", command=add_files)
            add_button.grid(row=1, column=0, padx=8, pady=5, sticky=tk.W)
            remove_button = ttk.Button(manage_frame, text="Remove Selected", command=remove_selected)
            remove_button.grid(row=2, column=0, padx=8, pady=5, sticky=tk.W)
            apply_button = ttk.Button(manage_frame, text="Apply", command=apply_changes)
            apply_button.grid(row=3, column=0, columnspan=3, pady=5)


def add_exhibit():
    exhibit_letter = exhibit_letter_entry.get().strip().upper()
    if not exhibit_letter:
        messagebox.showwarning("Input Error", "Exhibit letter cannot be empty.")
        return
    if exhibit_letter in exhibit_paths:
        messagebox.showwarning("Duplicate Error", f"Exhibit {exhibit_letter} already added.")
        return
    exhibit_files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if exhibit_files:
        exhibit_paths[exhibit_letter] = list(exhibit_files)
        exhibits_listbox.insert(tk.END, f"{exhibit_letter}: {', '.join([os.path.basename(f) for f in exhibit_files])}")

def select_main_pdf():
    main_pdf_path.set(filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")]))

def select_output_path():
    output_path.set(filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")]))

root = TkinterDnD.Tk()
root.title("IPP Exhibits Manager")

style = ttk.Style(root)

root.iconbitmap('C:/Users/omars_qgwtbpx/Documents/ilprisonproject_logo.ico')

frame = ttk.Frame(root, padding="10 10 10 10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

frame.columnconfigure(0, weight=1)
frame.columnconfigure(1, weight=3)
frame.columnconfigure(2, weight=1)

main_pdf_path = tk.StringVar()
output_path = tk.StringVar()
exhibit_paths = {}

ttk.Label(frame, text="Main PDF:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
main_pdf_entry = ttk.Entry(frame, textvariable=main_pdf_path, width=50)
main_pdf_entry.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
ttk.Button(frame, text="Browse", command=select_main_pdf).grid(row=0, column=2, padx=5, pady=5)

ttk.Label(frame, text="Exhibit Letter:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
exhibit_letter_entry = ttk.Entry(frame)
exhibit_letter_entry.grid(row=1, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
ttk.Button(frame, text="Add Exhibit", command=add_exhibit).grid(row=1, column=2, padx=5, pady=5)

exhibits_listbox = tk.Listbox(frame, width=50, height=20)
exhibits_listbox.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky=(tk.W, tk.E))
exhibits_listbox.bind("<Double-1>", manage_exhibit)

created_by_label = ttk.Label(root, text="Created by Omar Sobhy")
created_by_label.grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)


ttk.Label(frame, text="Output Path:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
output_path_entry = ttk.Entry(frame, textvariable=output_path, width=50)
output_path_entry.grid(row=3, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
ttk.Button(frame, text="Browse", command=select_output_path).grid(row=3, column=2, padx=5, pady=5)

merge_button = ttk.Button(frame, text="Merge PDFs", command=lambda: merge_pdfs(main_pdf_path.get(), exhibit_paths, output_path.get()))
merge_button.grid(row=4, column=0, columnspan=3, pady=10)

for child in frame.winfo_children():
    child.grid_configure(padx=5, pady=5)

sv_ttk.set_theme("dark")

root.mainloop()


# pyinstaller --onefile --noconsole --icon "C:/Users/omars_qgwtbpx/Documents/ilprisonproject_logo.ico" IPPExhibit.py