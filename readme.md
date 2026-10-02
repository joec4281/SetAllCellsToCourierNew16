I use LibreOffice Calc.

I often download files via Microsoft Edge,   
with the option to either open or save the file.

Even though I have LibreOffice Calc set to the default font of Courier New 16pt when creating a new spreadsheet,   
when I open a spreadsheet via Microsoft Edge, the font is not Courier New 16pt,   
and I have to press Ctrl-A to select everything in the spreadsheet,   
press Alt-1,   
then change the font to Courier New 16pt.  

This LibreOffice macro applies Courier New, 16 pt to all cells in the current Calc sheet:

Sub SetAllCellsToCourierNew16
    Dim document As Object
    Dim sheet As Object
    Dim cursor As Object
    
    document = ThisComponent
    sheet = document.CurrentController.ActiveSheet

    'Find the used range of the current sheet
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(True)

    'Apply the font and size
    cursor.CharFontName = "Courier New"
    cursor.CharHeight = 16
End Sub

**Install the macro**
* Open a spreadsheet in Calc.
* Press Alt+F11 to open the Basic Macros dialog.
* Select My Macros → Standard.
* Click New to create a module, if necessary.
* Give the module a name, such as Formatting.
* Paste the macro into the editor.
* Save it.
  
**Run it**
With a spreadsheet open:

* Press Alt+F11.
* Select SetAllCellsToCourierNew16.
* Click Run.

You can also assign it to a toolbar button or keyboard shortcut through Tools → Customize.
