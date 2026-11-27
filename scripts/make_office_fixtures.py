"""Have real Excel and PowerPoint author the live-gate fixtures (Windows, Office installed).

Usage:
    python scripts/make_office_fixtures.py [--excel] [--powerpoint]

Writes, under tests/fixtures/ (gitignored, never committed):

    xlam/RealAddin.xlam          Excel: ModuleKeep, ModuleEdit (TEST_BUILD = "ORIGINAL"), ModuleDelete
    xlam/FormSource.xlam         Excel: UserForm frmDemo (13 controls: a Frame with option
                                 buttons, a MultiPage with a label on a page, a ListBox,
                                 an Image and a CommandButton with a picture) whose code
                                 calls FormHelpers
    ppam/RealAddin.ppam          PowerPoint, the same three modules
    pptm/RealPresentation.pptm   PowerPoint, the same three modules

Excel builds its files through the VBA project object model ("Trust access to
the VBA project object model" must be on for Excel). PowerPoint needs no such
access: pyOpenVBA writes a presentation with the modules, and PowerPoint opens
it and saves it again as .ppam and .pptm, so PowerPoint authors both packages
and their VBA projects. Office runs as a separate, hidden instance.
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "tests" / "fixtures"

MODULES = {
    "ModuleKeep": 'Public Function KeepProbe() As String\n    KeepProbe = "KEEP_OK"\nEnd Function\n',
    "ModuleEdit": 'Public Const TEST_BUILD As String = "ORIGINAL"\n\n'
    "Public Function EditProbe() As String\n    EditProbe = TEST_BUILD\nEnd Function\n",
    "ModuleDelete": "Public Sub DeleteMe()\nEnd Sub\n",
}

FORM_CODE = """Option Explicit

Private Sub UserForm_Initialize()
    txtName.Value = DefaultName()
    cboKind.AddItem "Alpha"
    cboKind.AddItem "Beta"
    cboKind.ListIndex = 0
End Sub

Private Sub cmdOK_Click()
    lblStatus.Caption = Greeting(txtName.Value)
    Me.Hide
End Sub
"""

HELPER_CODE = """Option Explicit

Public Function DefaultName() As String
    DefaultName = "World"
End Function

Public Function Greeting(ByVal who As String) As String
    Greeting = "Hello, " & who
End Function
"""

# Run from a second workbook: a project cannot edit its own designers while
# one of its macros runs. A picture cannot cross processes, so VBA loads it.
SET_PICTURES = """Public Function SetPictures() As String
    Dim d As Object
    Set d = Workbooks("@BOOK@").VBProject.VBComponents("frmDemo").Designer
    Set d.Controls("imgLogo").Picture = LoadPicture("@BMP@")
    Set d.Controls("cmdOK").Picture = LoadPicture("@BMP@")
    d.Controls("cmdOK").PicturePosition = 1
    SetPictures = "ok"
End Function
"""


def _crlf(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\n", "\r\n")


def _bitmap(folder: Path) -> Path:
    from PIL import Image, ImageDraw

    path = folder / "logo.bmp"
    image = Image.new("RGB", (48, 32), (30, 110, 200))
    ImageDraw.Draw(image).rectangle((8, 8, 40, 24), fill=(250, 200, 40))
    image.save(path)
    return path


def _replace(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()


def make_excel(scratch: Path) -> list[Path]:
    import win32com.client as com

    bitmap = _bitmap(scratch)
    app = com.DispatchEx("Excel.Application")
    made: list[Path] = []
    try:
        app.Visible = False
        app.DisplayAlerts = False
        app.AutomationSecurity = 1

        book = app.Workbooks.Add()
        for name, code in MODULES.items():
            module = book.VBProject.VBComponents.Add(1)
            module.Name = name
            module.CodeModule.AddFromString(code)
        target = FIXTURES / "xlam" / "RealAddin.xlam"
        _replace(target)
        book.SaveAs(str(target), FileFormat=55)
        book.Close(False)
        made.append(target)

        book = app.Workbooks.Add()
        project = book.VBProject
        helpers = project.VBComponents.Add(1)
        helpers.Name = "FormHelpers"
        helpers.CodeModule.AddFromString(HELPER_CODE)
        form = project.VBComponents.Add(3)
        form.Name = "frmDemo"
        form.Properties("Caption").Value = "Demo Form"
        form.Properties("Width").Value = 320
        form.Properties("Height").Value = 260
        designer = form.Designer

        def add(progid: str, name: str, box: tuple[float, float, float, float], parent=None):
            holder = parent.Controls if parent is not None else designer.Controls
            control = holder.Add(progid, name)
            control.Left, control.Top, control.Width, control.Height = box
            return control

        add("Forms.Label.1", "lblTitle", (6, 6, 200, 18)).Caption = "Enter your name:"
        designer.Controls("lblTitle").Font.Bold = True
        add("Forms.TextBox.1", "txtName", (6, 26, 150, 18))
        add("Forms.ComboBox.1", "cboKind", (162, 26, 100, 18))
        frame = add("Forms.Frame.1", "fraOptions", (6, 50, 150, 80))
        frame.Caption = "Options"
        add("Forms.OptionButton.1", "optOne", (6, 6, 120, 18), frame).Caption = "First"
        add("Forms.OptionButton.1", "optTwo", (6, 28, 120, 18), frame).Caption = "Second"
        add("Forms.CheckBox.1", "chkRemember", (162, 52, 120, 18)).Caption = "Remember me"
        pages = add("Forms.MultiPage.1", "mpgTabs", (162, 76, 140, 80))
        pages.Pages(0).Controls.Add("Forms.Label.1", "lblInner").Caption = "On page 1"
        add("Forms.ListBox.1", "lstItems", (6, 136, 150, 40))
        add("Forms.Image.1", "imgLogo", (162, 160, 48, 32))
        add("Forms.CommandButton.1", "cmdOK", (220, 196, 80, 30)).Caption = "OK"
        add("Forms.Label.1", "lblStatus", (6, 196, 200, 18)).Caption = ""
        form.CodeModule.AddFromString(FORM_CODE)

        helper_book = app.Workbooks.Add()
        setup = helper_book.VBProject.VBComponents.Add(1)
        setup.CodeModule.AddFromString(
            SET_PICTURES.replace("@BOOK@", book.Name).replace("@BMP@", str(bitmap))
        )
        result = app.Run(f"'{helper_book.Name}'!{setup.Name}.SetPictures")
        helper_book.Close(False)
        if result != "ok":
            raise RuntimeError(f"Excel could not set the pictures: {result}")
        target = FIXTURES / "xlam" / "FormSource.xlam"
        _replace(target)
        book.SaveAs(str(target), FileFormat=55)
        book.Close(False)
        made.append(target)
    finally:
        app.Quit()
    return made


def make_powerpoint(scratch: Path) -> list[Path]:
    import win32com.client as com
    from pyopenvba import PowerPointFile

    seed = scratch / "Seed.pptm"
    with PowerPointFile.create_new(seed) as host:
        project = host.vba_project()
        for name, code in MODULES.items():
            project.add_module(name, _crlf(code))
        host.save()

    app = com.DispatchEx("PowerPoint.Application")
    made: list[Path] = []
    try:
        app.DisplayAlerts = 1
        app.AutomationSecurity = 1
        deck = app.Presentations.Open(str(seed), ReadOnly=False, Untitled=False, WithWindow=False)
        try:
            for target, file_format in (
                (FIXTURES / "pptm" / "RealPresentation.pptm", 25),  # ppSaveAsOpenXMLPresentationMacroEnabled
                (FIXTURES / "ppam" / "RealAddin.ppam", 30),  # ppSaveAsOpenXMLAddin
            ):
                _replace(target)
                deck.SaveCopyAs(str(target), file_format)
                made.append(target)
        finally:
            deck.Close()
    finally:
        app.Quit()
    return made


def main(argv: list[str]) -> int:
    want_excel = "--excel" in argv or not argv
    want_powerpoint = "--powerpoint" in argv or not argv
    import pythoncom

    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from _office_procs import live_pids, stop_new

    if want_powerpoint and live_pids("POWERPNT.EXE"):
        print("Close PowerPoint first: it runs one instance, which this script would quit.")
        return 1
    excel_before, ppt_before = live_pids("EXCEL.EXE"), live_pids("POWERPNT.EXE")
    pythoncom.CoInitialize()
    scratch = Path(tempfile.mkdtemp(prefix="vbaae-fixtures-"))
    try:
        made: list[Path] = []
        if want_excel:
            made += make_excel(scratch)
        if want_powerpoint:
            made += make_powerpoint(scratch)
    finally:
        pythoncom.CoUninitialize()
        stop_new("EXCEL.EXE", excel_before)
        stop_new("POWERPNT.EXE", ppt_before)
        shutil.rmtree(scratch, ignore_errors=True)
    for path in made:
        print(f"wrote {path.relative_to(ROOT)}  ({path.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
