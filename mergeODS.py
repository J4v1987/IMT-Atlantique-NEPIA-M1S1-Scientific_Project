import re
import gi
import subprocess
import os
import copy
import shutil

gi.require_version("Gtk", "3.0")
from gi.repository import Gtk

from odf.opendocument import load, OpenDocumentSpreadsheet
from odf.table import Table

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


def select_input_files():
    dialog = Gtk.FileChooserDialog(
        title="Select ODS workbooks to merge",
        action=Gtk.FileChooserAction.OPEN,
        buttons=(
            Gtk.STOCK_CANCEL,
            Gtk.ResponseType.CANCEL,
            Gtk.STOCK_OPEN,
            Gtk.ResponseType.OK,
        ),
    )
    dialog.set_select_multiple(True)

    filter_ods = Gtk.FileFilter()
    filter_ods.set_name("ODS files")
    filter_ods.add_pattern("*.ods")
    dialog.add_filter(filter_ods)

    response = dialog.run()
    files = dialog.get_filenames() if response == Gtk.ResponseType.OK else []
    dialog.destroy()

    return files


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

def merge_ods_workbooks(input_files, output_path):
    # Step 2: create blank pivot workbook
    pivot_doc = OpenDocumentSpreadsheet()

    for ods_path in input_files:
        # Step 3: open data workbook
        doc = load(ods_path)

        # Step 4: get first (and only) sheet
        sheets = doc.spreadsheet.getElementsByType(Table)
        if not sheets:
            continue

        sheet = sheets[0]

        # Rename sheet to workbook name (without extension)
        base_name = os.path.splitext(os.path.basename(ods_path))[0]
        sheet.setAttribute("name", base_name)

        # Step 5: deep-copy sheet into pivot workbook
        pivot_doc.spreadsheet.addElement(copy.deepcopy(sheet))

        # Step 6: close source workbook
        doc = None  # explicit dereference (good habit)

    # Step 8: save pivot workbook
    pivot_doc.save(output_path)

def main():
    input_files = select_input_files()
    if not input_files:
        #print("No files selected.")
        show_message("No files selected.")
        return

    # Use directory of first selected workbook
    default_dir = os.path.dirname(input_files[0])

    output_path = select_output_file(
        default_dir=default_dir,
        input_file_path=input_files[0]
    )

    if not output_path:
        #print("No output file selected.")
        show_message("No output file selected.")
        return

    merge_ods_workbooks(input_files, output_path)
    #print(f"Merged workbook saved to: {output_path}")
    show_message(f"Merged workbook saved to: {output_path}")
    if shutil.which("libreoffice"):
        subprocess.Popen(["libreoffice", output_path])
    else:
        subprocess.Popen(["xdg-open", output_path])

if __name__ == "__main__":
    main()
