import re
import gi
import subprocess
import os
import shutil

gi.require_version("Gtk", "3.0")
from gi.repository import Gtk

from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P

'''
def select_input_file():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="Select UTF-8 text file",
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
    )
'''
    
def select_input_file():
    dialog = Gtk.FileChooserNative(
        title="Select UTF-8 text file",
        action=Gtk.FileChooserAction.OPEN
    )

    txt_filter = Gtk.FileFilter()
    txt_filter.set_name("Text files")
    txt_filter.add_pattern("*.txt")
    dialog.add_filter(txt_filter)

    all_filter = Gtk.FileFilter()
    all_filter.set_name("All files")
    all_filter.add_pattern("*")
    dialog.add_filter(all_filter)

    response = dialog.run()
    filename = dialog.get_filename() if response == Gtk.ResponseType.ACCEPT else None
    dialog.destroy()
    return filename
    
'''
def select_output_file():
    root = tk.Tk()
    root.withdraw()
    return filedialog.asksaveasfilename(
        title="Save LibreOffice Calc file",
        defaultextension=".ods",
        filetypes=[("LibreOffice Calc", "*.ods")]
    )
'''

def select_output_file(default_dir, input_file_path):
    # Extract the base name of the input file without extension
    base_name = os.path.splitext(os.path.basename(input_file_path))[0]
    default_save_name = f"{base_name}.ods"

    dialog = Gtk.FileChooserNative(
        title="Save LibreOffice Calc file",
        action=Gtk.FileChooserAction.SAVE
    )
    
    # Suggest default name
    dialog.set_do_overwrite_confirmation(True)
    dialog.set_current_name(default_save_name)#("keff_results.ods")

    # Set default folder
    if default_dir and os.path.isdir(default_dir):
        dialog.set_current_folder(default_dir)

    # Add ODS filter
    ods_filter = Gtk.FileFilter()
    ods_filter.set_name("LibreOffice Calc (*.ods)")
    ods_filter.add_mime_type("application/vnd.oasis.opendocument.spreadsheet")
    dialog.add_filter(ods_filter)

    # Run dialog
    response = dialog.run()
    filename = dialog.get_filename() if response == Gtk.ResponseType.ACCEPT else None
    dialog.destroy()

    # Ensure .ods extension
    if filename and not filename.lower().endswith(".ods"):
        filename += ".ods"

    return filename

def show_message(text, title="Info"):
    dialog = Gtk.MessageDialog(
        transient_for=None,
        flags=0,
        message_type=Gtk.MessageType.INFO,
        buttons=Gtk.ButtonsType.OK,
        text=title,
    )
    dialog.format_secondary_text(text)
    dialog.run()
    dialog.destroy()

def show_error(text, title="Error"):
    dialog = Gtk.MessageDialog(
        transient_for=None,
        flags=0,
        message_type=Gtk.MessageType.ERROR,
        buttons=Gtk.ButtonsType.CLOSE,
        text=title,
    )
    dialog.format_secondary_text(text)
    dialog.run()
    dialog.destroy()

def open_ods_file(path):
    if os.path.isfile(path):
        try:
            subprocess.run(["libreoffice", path], check=True)
        except Exception as e:
            print(f"Failed to open file: {e}")

def parse_keff_values(filepath):
    analog_values = []
    implicit_values = []

    analog_pattern = re.compile(r"=\s*([0-9]*\.[0-9]+)")
    implicit_pattern = re.compile(r"=\s*([0-9]*\.[0-9]+)")

    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            head = line[:19]

            if "k-eff (analog)" in head:
                match = analog_pattern.search(line)
                if match:
                    analog_values.append(float(match.group(1)))

            elif "k-eff (implicit)" in head:
                match = implicit_pattern.search(line)
                if match:
                    implicit_values.append(float(match.group(1)))

    return analog_values, implicit_values


def create_spreadsheet(analog_values, implicit_values, output_path):
    doc = OpenDocumentSpreadsheet()
    table = Table(name="k-eff data")

    start_row = 9  # first data row (1-based)
    header_row_index = start_row - 1
    for _ in range(header_row_index - 1):
        table.addElement(TableRow())

    header_row = TableRow()

    # Column A (empty)
    header_row.addElement(TableCell())

    # Column B — Analog
    keffAnalogHeader = TableCell()
    keffAnalogHeader.addElement(P(text="k-eff (analog)"))
    header_row.addElement(keffAnalogHeader)

    # Column C — Implicit
    keffImplicitHeader = TableCell()
    keffImplicitHeader.addElement(P(text="k-eff (implicit)"))
    header_row.addElement(keffImplicitHeader)

    table.addElement(header_row)

    max_len = max(len(analog_values), len(implicit_values))

    for i in range(max_len):
        row = TableRow()

        # Column A (empty)
        row.addElement(TableCell())

        # Column B – Analog
        cell_b = TableCell()
        if i < len(analog_values):
            cell_b.addElement(P(text=str(analog_values[i])))
        row.addElement(cell_b)

        # Column C – Implicit
        cell_c = TableCell()
        if i < len(implicit_values):
            cell_c.addElement(P(text=str(implicit_values[i])))
        row.addElement(cell_c)

        table.addElement(row)

    doc.spreadsheet.addElement(table)
    doc.save(output_path)

def main():
    input_file = select_input_file()
    if not input_file:
        show_error("No input file selected.")
        return

    input_dir = os.path.dirname(input_file)
    '''if not input_file:
        messagebox.showerror("Error", "No input file selected.")
        return'''

    analog_values, implicit_values = parse_keff_values(input_file)

    if not analog_values and not implicit_values:
        show_error("No k-eff values found in the selected file.")
        return

    output_file = select_output_file(input_dir, input_file)
    if not output_file:
        #messagebox.showerror("Error", "No output file selected.")
        show_error("No output file selected.")
        return

    create_spreadsheet(analog_values, implicit_values, output_file)

    '''messagebox.showinfo(
        "Done",
        f"Spreadsheet successfully saved:\n{output_file}"
    )'''
    show_message(f"Spreadsheet successfully saved:\n{output_file}", "Done")
    #open_ods_file(output_file) #python runtime blocking
    #subprocess.Popen(["libreoffice", output_file])
    if shutil.which("libreoffice"):
        subprocess.Popen(["libreoffice", output_file])
    else:
        subprocess.Popen(["xdg-open", output_file])

if __name__ == "__main__":
    main()