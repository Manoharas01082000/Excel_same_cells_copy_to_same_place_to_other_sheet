# Excel_same_cells_copy_to_same_place_to_other_sheet
Copying the same value in  multiple cells in an Excel sheet to another Excel sheet in the same position
# CopyNPCells VBA Macro

## Description
This VBA macro copies cells containing the value "NP" from a specified range in one sheet to the same positions in another sheet within the same workbook. The range can be explicitly defined by setting the maximum row and column to be processed.

## Steps to Use the VBA Macro

### 1. Open Excel and Your Workbook
Ensure the workbook containing the source and target sheets is open.

### 2. Check Sheet Names
Make sure the names of the sheets in your workbook are exactly:
- `SourceSheetName`
- `TargetSheetName`

### 3. Open the VBA Editor
Press `Alt + F11` to open the Visual Basic for Applications editor.

### 4. Insert a New Module
In the VBA editor, go to `Insert > Module`. This will create a new module where you can write your VBA code.

### 5. Copy the VBA Code into the Module
Paste the provided VBA code into the module:

```vba
Sub CopyNPCells()
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim cell As Range
    Dim maxRow As Long
    Dim maxCol As Long
    Dim r As Long, c As Long

    ' Define your source and target sheets
    On Error GoTo ErrHandler
    Set sourceSheet = ThisWorkbook.Sheets("SourceSheetName") ' Replace with your source sheet name
    Set targetSheet = ThisWorkbook.Sheets("TargetSheetName") ' Replace with your target sheet name

    ' Specify the maximum row and column to be processed
    maxRow = 100 ' Replace with the maximum row you want to process
    maxCol = 50  ' Replace with the maximum column you want to process

    ' Turn off screen updating and automatic calculations to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Loop through each cell in the specified range of the source sheet
    For r = 1 To maxRow
        For c = 1 To maxCol
            If sourceSheet.Cells(r, c).Value = "NP" Then
                ' Copy the cell to the same position in the target sheet
                targetSheet.Cells(r, c).Value = sourceSheet.Cells(r, c).Value
            End If
        Next c
    Next r

    ' Turn on screen updating and automatic calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical

End Sub
###6. Set Maximum Rows and Columns
In the VBA code, set the maxRow and maxCol variables to define the maximum row and column to process. For example:


maxRow = 100 ' Replace with the maximum row you want to process
maxCol = 50  ' Replace with the maximum column you want to process

###7. Ensure Sheet Names are Correct
Make sure the sheet names in the VBA code exactly match the sheet names in your workbook.

###8. Run the Macro
Close the VBA editor by clicking the X or pressing Alt + Q.
Press Alt + F8 to open the "Macro" dialog box.
Select CopyNPCells from the list and click Run.
Explanation of the Code
Define Sheets

Set sourceSheet = ThisWorkbook.Sheets("SourceSheetName")
Set targetSheet = ThisWorkbook.Sheets("TargetSheetName")
Specify Range
Set the maximum row and column to process:

maxRow = 100 ' Replace with the maximum row you want to process
maxCol = 50  ' Replace with the maximum column you want to process
Optimize Performance
Turn off screen updating and automatic calculations to improve performance:

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Loop Through Cells
Loop through each cell in the specified range of the source sheet and copy cells with "NP" to the target sheet:

For r = 1 To maxRow
    For c = 1 To maxCol
        If sourceSheet.Cells(r, c).Value = "NP" Then
            targetSheet.Cells(r, c).Value = sourceSheet.Cells(r, c).Value
        End If
    Next c
Next r
Restore Settings
Turn on screen updating and automatic calculations:

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Testing the Code
Create a Test Workbook
Create a new Excel workbook.
Name the first sheet as SourceSheetName.
Name the second sheet as TargetSheetName.
Populate some cells in SourceSheetName with the value "NP".
Run the Macro
Follow the steps to run the macro as described above.

This updated code ensures that the macro processes cells up to the specified maximum row and column and correctly copies cells containing "NP" from one sheet to another in the specified range.

Feel free to copy and use this `README` file in your project.




