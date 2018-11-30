Sub MarkAdvanceEntry()
Application.CutCopyMode = False

iRow = ActiveCell.Row

planC = Cell(iRow,5) '5/E plan
employerC = Cell(iRow,6) '6/F employer
adminPersonC = Cell(iRow,7) '7/G adminPerson
effectiveDateC = Cell(iRow,10) '10/J effective date
descriptionC = Cell(iRow,12) '12/L Description
daysBetweenC = Cell(iRow,14) '14/N Days between
dayOfWeekC = Cell(iRow,16)'16/P day of week
firstDateNYC = Cell(iRow,17) '17/Q 1st 2019 date

iRow = ActiveCell.Row
ActiveCell.EntireRow.Select
Selection.Interior.Color = vbGreen

Do While Rows(iRow + 1).EntireRow.Hidden
    iRow = iRow + 1
Loop

Rows(iRow + 1).EntireRow.Select

End Sub
