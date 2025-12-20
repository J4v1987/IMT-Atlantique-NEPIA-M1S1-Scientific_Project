import re
import gi
<<<<<<< HEAD
import os

=======
>>>>>>> 6c00fa11987f3379014d7f0c281df1a36d8c1ad2
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
<<<<<<< HEAD
    
=======

>>>>>>> 6c00fa11987f3379014d7f0c281df1a36d8c1ad2
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

<<<<<<< HEAD
def select_output_file(default_dir):
=======
def select_output_file():
>>>>>>> 6c00fa11987f3379014d7f0c281df1a36d8c1ad2
    dialog = Gtk.FileChooserNative(
        title="Save LibreOffice Calc file",
        action=Gtk.FileChooserAction.SAVE
    )

    dialog.set_do_overwrite_confirmation(True)
    dialog.set_current_name("keff_results.ods")

<<<<<<< HEAD
    if default_dir and os.path.isdir(default_dir):
        dialog.set_current_folder(default_dir)

=======
>>>>>>> 6c00fa11987f3379014d7f0c281df1a36d8c1ad2
    ods_filter = Gtk.FileFilter()
    ods_filter.set_name("LibreOffice Calc (*.ods)")
    ods_filter.add_mime_type(
        "application/vnd.oasis.opendocument.spreadsheet"
    )
    dialog.add_filter(ods_filter)

    response = dialog.run()
    filename = dialog.get_filename() if response == Gtk.ResponseType.ACCEPT else None
    dialog.destroy()

    if filename and not filename.lower().endswith(".ods"):
        filename += ".ods"

    return filename

<<<<<<< HEAD
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
=======
>>>>>>> 6c00fa11987f3379014d7f0c281df1a36d8c1ad2

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

    start_row = 9  # LibreOffice rows are 1-based

    max_len = max(len(analog_values), len(implicit_values))

    for i in range(start_row - 1):
        table.addElement(TableRow())

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
<<<<<<< HEAD
        show_error("No input file selected.")
        return

    input_dir = os.path.dirname(input_file)
    '''if not input_file:
        messagebox.showerror("Error", "No input file selected.")
        return'''
=======
        messagebox.showerror("Error", "No input file selected.")
        return
>>>>>>> 6c00fa11987f3379014d7f0c281df1a36d8c1ad2

    analog_values, implicit_values = parse_keff_values(input_file)

    if not analog_values and not implicit_values:
<<<<<<< HEAD
        show_error("No k-eff values found in the selected file.")
        return

    output_file = select_output_file(input_dir)
    if not output_file:
        #messagebox.showerror("Error", "No output file selected.")
        show_error("No output file selected.")
=======
        messagebox.showerror("Error", "No k-eff values found.")
        return

    output_file = select_output_file()
    if not output_file:
        messagebox.showerror("Error", "No output file selected.")
>>>>>>> 6c00fa11987f3379014d7f0c281df1a36d8c1ad2
        return

    create_spreadsheet(analog_values, implicit_values, output_file)

<<<<<<< HEAD
    '''messagebox.showinfo(
        "Done",
        f"Spreadsheet successfully saved:\n{output_file}"
    )'''
    show_message(f"Spreadsheet successfully saved:\n{output_file}", "Done")
=======
    messagebox.showinfo(
        "Done",
        f"Spreadsheet successfully saved:\n{output_file}"
    )
>>>>>>> 6c00fa11987f3379014d7f0c281df1a36d8c1ad2


if __name__ == "__main__":
    main()