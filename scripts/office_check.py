"""Load an add-in in real Excel or PowerPoint and report on its UserForms.

Usage:
    python scripts/office_check.py PATH FORM [FORM ...]

The file must hold the report module ``VbaaeLiveCheck`` (``REPORT_MODULE``
below); the live UserForm test adds it through the editor, like any module,
so no access to the VBA project object model is needed. For every FORM the
report loads the form (``UserForms.Add``, which compiles its code and runs
its Initialize), then lists each control with its type, position, size,
caption, tip, font, whether it shows a picture and its value.

Excel opens the .xlam read only. PowerPoint opens a .pptm read only, and
loads a .ppam as an add-in for the run (``AddIns.Add``), removing it after.
Office is started as a separate, hidden instance and quit at the end. Prints
the report and ``OFFICE CHECK OK``; exits 1 when Office refuses the file or a
form, 2 for a usage error. Run it with a timeout: a VBA error dialog would
otherwise wait for a click.
"""

from __future__ import annotations

import sys
from pathlib import Path

REPORT_MODULE = "VbaaeLiveCheck"
# Every property read is guarded by the control's type, so the report raises
# no error even with "Break on All Errors" set in the VBA editor.
REPORT_CODE = '''Option Explicit

Private Function Has(ByVal o As Object, ByVal what As String) As Boolean
    Dim t As String
    t = TypeName(o)
    Select Case what
        Case "cap": Has = (t = "Label" Or t = "CommandButton" Or t = "CheckBox" Or t = "OptionButton" Or t = "ToggleButton" Or t = "Frame")
        Case "font": Has = Not (t = "Image" Or t = "SpinButton" Or t = "ScrollBar" Or t = "MultiPage")
        Case "pic": Has = (t = "Image" Or t = "CommandButton")
        Case "value": Has = (t = "TextBox" Or t = "ComboBox" Or t = "CheckBox" Or t = "OptionButton")
    End Select
End Function

Public Function VbaaeFormReport(ByVal formName As String) As String
    Dim f As Object, c As Object, s As String
    Set f = VBA.UserForms.Add(formName)
    s = "FORM " & formName & " caption=" & f.Caption & " controls=" & f.Controls.Count
    For Each c In f.Controls
        s = s & vbLf & "  " & c.Name & ":" & TypeName(c) & " @" & c.Left & "," & c.Top & " " & c.Width & "x" & c.Height
        If Has(c, "cap") Then s = s & " cap=[" & c.Caption & "]"
        If Len(c.ControlTipText) > 0 Then s = s & " tip=[" & c.ControlTipText & "]"
        If Has(c, "font") Then s = s & " font=" & c.Font.Name & "/" & c.Font.Size & IIf(c.Font.Bold, "/b", "")
        If Has(c, "pic") Then
            If c.Picture Is Nothing Then
                s = s & " pic=False"
            Else
                s = s & " pic=" & (c.Picture.Handle <> 0)
            End If
        End If
        If Has(c, "value") Then s = s & " value=" & c.Value
    Next
    Unload f
    VbaaeFormReport = s
End Function
'''


def _excel(path: Path, forms: list[str]) -> list[str]:
    import win32com.client as com

    app = com.DispatchEx("Excel.Application")
    try:
        app.Visible = False
        app.DisplayAlerts = False
        app.AutomationSecurity = 1  # msoAutomationSecurityLow: the file's macros run
        book = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
        try:
            return [app.Run(f"'{book.Name}'!{REPORT_MODULE}.VbaaeFormReport", name) for name in forms]
        finally:
            book.Close(False)
    finally:
        app.Quit()


def _powerpoint(path: Path, forms: list[str]) -> list[str]:
    import win32com.client as com

    app = com.DispatchEx("PowerPoint.Application")
    addin = None
    deck = None
    try:
        app.DisplayAlerts = 1  # ppAlertsNone
        app.AutomationSecurity = 1
        if path.suffix.lower() == ".ppam":
            addin = app.AddIns.Add(str(path))
            addin.Loaded = -1
            macro = f"{path.name}!{REPORT_MODULE}.VbaaeFormReport"
        else:
            deck = app.Presentations.Open(str(path), ReadOnly=True, Untitled=False, WithWindow=False)
            macro = f"'{deck.Name}'!{REPORT_MODULE}.VbaaeFormReport"
        return [app.Run(macro, name) for name in forms]
    finally:
        if deck is not None:
            deck.Close()
        if addin is not None:
            name = addin.Name
            addin.Loaded = 0
            app.AddIns.Remove(name)
        app.Quit()


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print(__doc__)
        return 2
    path = Path(argv[0]).resolve()
    forms = argv[1:]
    import pythoncom

    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from _office_procs import live_pids, stop_new

    excel = path.suffix.lower() == ".xlam"
    image = "EXCEL.EXE" if excel else "POWERPNT.EXE"
    before = live_pids(image)
    if before and not excel:
        # PowerPoint runs one instance: the check would use, and then quit, yours.
        print("OFFICE CHECK FAILED: close PowerPoint first.")
        return 1
    pythoncom.CoInitialize()
    try:
        reports = (_excel if excel else _powerpoint)(path, forms)
    except Exception as exc:  # noqa: BLE001 - reported to the caller
        print(f"OFFICE CHECK FAILED: {type(exc).__name__}: {exc}")
        return 1
    finally:
        pythoncom.CoUninitialize()
        stop_new(image, before)
    for report in reports:
        print(str(report).replace("\r", ""))
    print("OFFICE CHECK OK")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
