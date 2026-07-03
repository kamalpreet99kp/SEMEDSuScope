"""Import and run the Au Word report VBA macro automatically.

This helper avoids manually importing `create_au_word_report_macro.bas` into Word
each time. It creates a temporary macro-enabled Word host document, imports the
`.bas` module into that document, reads the actual module name that Word assigned
to the import, runs `CreateAuWordReportFromWorkbook` from that host, and then
closes the temporary host. The host is closed even if Word rejects one macro name
style so the temporary `.docm` file is not left locked.

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


def component_names(vb_components) -> set[str]:
    """Return the current VBA component names in a Word document."""
    names = set()
    try:
        component_count = vb_components.Count
    except Exception:
        return names
    for index in range(1, component_count + 1):
        try:
            names.add(str(vb_components(index).Name))
        except Exception:
            continue
    return names


def import_macro_module(host_document, module_path: Path) -> str:
    """Import the .bas macro module and return the module name Word assigned."""
    try:
        vb_components = host_document.VBProject.VBComponents
        before_names = component_names(vb_components)
        imported_component = vb_components.Import(str(module_path.resolve()))
        if imported_component is not None:
            return str(imported_component.Name)
        after_names = component_names(vb_components)
        added_names = sorted(after_names - before_names)
        if added_names:
            return added_names[0]
    except Exception as error:
        raise MacroImportError(
            "Word could not import the VBA macro module. Enable Word's 'Trust access to the VBA project object model' "
            "setting, then run this script again."
        ) from error
    raise MacroImportError("Word imported the VBA macro file, but the imported module name could not be identified.")


def document_full_name(host_document) -> str:
    """Return the host document full name when Word exposes it."""
    try:
        return str(host_document.FullName)
    except Exception:
        return ""


def qualified_macro_name(host_document, module_name: str) -> str:
    """Build the preferred document-qualified macro name for the imported module."""
    return f"'{host_document.Name}'!{module_name}.{MACRO_PROCEDURE_NAME}"


def macro_name_candidates(host_document, module_name: str) -> tuple[str, ...]:
    """Return Word macro name styles to try, most-specific first."""
    module_qualified_name = f"{module_name}.{MACRO_PROCEDURE_NAME}"
    host_name = str(host_document.Name)
    host_full_name = document_full_name(host_document)
    candidates = [
        qualified_macro_name(host_document, module_name),
        f"{host_name}!{module_qualified_name}",
        f"'{host_name}'!{MACRO_PROCEDURE_NAME}",
        f"{host_name}!{MACRO_PROCEDURE_NAME}",
    ]
    if host_full_name:
        candidates.insert(0, f"'{host_full_name}'!{module_qualified_name}")
        candidates.append(f"'{host_full_name}'!{MACRO_PROCEDURE_NAME}")
    return tuple(dict.fromkeys(candidates))


def run_word_macro(word_app, host_document, module_name: str) -> None:
    """Run the imported Word VBA macro."""
    errors = []
    for macro_name in macro_name_candidates(host_document, module_name):
        try:
            word_app.Run(macro_name)
            return
        except Exception as error:
            errors.append(f"{macro_name!r}: {error}")
    raise MacroImportError("Word could not run the imported macro. Tried:\n" + "\n".join(errors))


def close_document_without_saving(document) -> None:
    """Close a Word document, ignoring cleanup errors from Word COM."""
    if document is None:
        return
    try:
        document.Close(SaveChanges=False)
    except Exception:
        pass


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
            host_document = None
            try:
                host_document = word.Documents.Add()
                host_document.SaveAs2(str(host_path), FileFormat=WD_FORMAT_XML_DOCUMENT_MACRO_ENABLED)
                module_name = import_macro_module(host_document, module_path)
                host_document.Save()
                host_document.Activate()
                run_word_macro(word, host_document, module_name)
            finally:
                close_document_without_saving(host_document)
    except MacroImportError as error:
        messagebox.showerror("Au Word macro failed", str(error))
    finally:
        word.DisplayAlerts = previous_alerts
        word.Quit(SaveChanges=False)


if __name__ == "__main__":
    main()
