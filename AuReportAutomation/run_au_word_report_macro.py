"""Import and run the Au Word report VBA macro automatically.

This helper avoids manually importing `create_au_word_report_macro.bas` into Word
each time. It creates a temporary macro-enabled Word host document, imports the
`.bas` module into that document, runs `CreateAuWordReportFromWorkbook`, and then
closes the temporary host.

Important: Word must allow programmatic VBA access. In Word, enable:
File > Options > Trust Center > Trust Center Settings > Macro Settings >
"Trust access to the VBA project object model".
"""

from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
from tkinter import Tk, messagebox

MACRO_FILE_NAME = "create_au_word_report_macro.bas"
MACRO_PROCEDURE_NAME = "CreateAuWordReportFromWorkbook"
WD_FORMAT_XML_DOCUMENT_MACRO_ENABLED = 13
WD_ALERTS_NONE = 0
WD_ALERTS_ALL = -1


class MacroImportError(RuntimeError):
    """Raised when Word cannot import or run the VBA macro."""


def macro_path() -> Path:
    """Return the local VBA macro path."""
    return Path(__file__).with_name(MACRO_FILE_NAME)


def import_macro_module(host_document, module_path: Path) -> None:
    """Import the .bas macro module into a Word document's VBA project."""
    try:
        host_document.VBProject.VBComponents.Import(str(module_path.resolve()))
    except Exception as error:
        raise MacroImportError(
            "Word could not import the VBA macro module. Enable Word's 'Trust access to the VBA project object model' "
            "setting, then run this script again."
        ) from error


def run_word_macro(word_app, macro_name: str) -> None:
    """Run the imported Word VBA macro."""
    try:
        word_app.Run(macro_name)
    except Exception as error:
        raise MacroImportError(f"Word could not run the imported macro {macro_name!r}.") from error


def main() -> None:
    """Create a temporary Word macro host, import the Au macro, and run it."""
    root = Tk()
    root.withdraw()

    module_path = macro_path()
    if not module_path.exists():
        messagebox.showerror("Missing macro", f"Could not find VBA macro file:\n\n{module_path}")
        return

    import win32com.client

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    previous_alerts = getattr(word, "DisplayAlerts", WD_ALERTS_ALL)
    word.DisplayAlerts = WD_ALERTS_NONE

    try:
        with TemporaryDirectory(prefix="au_word_macro_") as temporary_directory:
            host_path = Path(temporary_directory) / "AuWordMacroHost.docm"
            host_document = word.Documents.Add()
            host_document.SaveAs2(str(host_path), FileFormat=WD_FORMAT_XML_DOCUMENT_MACRO_ENABLED)
            import_macro_module(host_document, module_path)
            run_word_macro(word, MACRO_PROCEDURE_NAME)
            host_document.Close(SaveChanges=False)
    except MacroImportError as error:
        messagebox.showerror("Au Word macro failed", str(error))
    finally:
        word.DisplayAlerts = previous_alerts
        word.Quit(SaveChanges=False)


if __name__ == "__main__":
    main()
