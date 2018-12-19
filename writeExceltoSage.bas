Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub writeExceltoSage()
' writes a line from workbook to sage55, makes decisions based on supplied codes
'
' Keyboard Shortcut: Ctrl+e
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''DECLARATIONS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'macro name
functionName = "writeExceltoSage"

'important indicies
currentExcelRow = ActiveCell.Row
currentSageRow = Range("sage_50_row_number").Value
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

debugHead = Trim(Str(currentSageRow)) + "," + Trim(Str(currentExcelRow)) + ":" + functionName + ">    "
Debug.Print debugHead + "starting..."
ActiveWindow.ScrollRow = currentExcelRow

If currentSageRow > 1 And currentExcelRow = 2 Then
    Debug.Print debugHead + "sage/excel index mismatch! Aborting..."
    Exit Sub
End If

SetClipboard ("666")

'window titles
sageWindow = "Sales/Invoicing (BLRDCB1.capitalsg.local)"
excelWindow = "PeachtreeDirectBill_v2_macros.xlsm"
    
'sage vars          Sage Field
sQuant = ""         'Quantity
sItem = ""          'Item
sDesc = ""          'Description
sUnitP = ""         'Unit Price
sTax = ""           'Tax
sAmnt = ""          'Amount
sJob = ""           'Job
sGLcode = ""

'Providers for P case
serviceProvider1 = "BlueStar Retirement Services, Inc."
serviceProvider2 = "TD Ameritrade Trust Company"
'p case desc string
saaDesc = "401(k) Quarterly Administration Fee"
sacDesc = "401(k) Quarterly Per Account Fee"
spDesc = "401(k) Quarterly Fulfillment Services"
sa1Desc = "Asset Fee"
sa2Desc = "Custodian Fee"

saaGL = "41201"
sacGL = "41202"
spGL = "41202"
sa1GL = "41301"
sa2GL = "41301"


'excel vars             Excel Col
planId = ""             'A
planCode = ""           'B
serviceProvider = ""    'C
adminFeeDB = ""         'D
assetfeedb = ""         'E
accountCount = ""       'F
participantCount = ""   'G
perAcctPtFee = ""       'H
preTotal = ""           'I
revShare = ""           'J
invoiceTotal = ""       'K
noteField = ""          'L

    
'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''''

AppActivate (excelWindow)

'Debug.Print debugHead + "starting Excel copy operations..."


planId = Cells(currentExcelRow, 1)           'A
planCode = Cells(currentExcelRow, 2)         'B
serviceProvider = Cells(currentExcelRow, 3)  'C
adminFeeDB = Cells(currentExcelRow, 4)       'D
assetfeedb = Cells(currentExcelRow, 5)       'E
accountCount = Cells(currentExcelRow, 6)     'F
participantCount = Cells(currentExcelRow, 7) 'G
perAcctPtFee = Cells(currentExcelRow, 8)     'H
preTotal = Cells(currentExcelRow, 9)         'I
revShare = Cells(currentExcelRow, 10)        'J
invoiceTotal = Cells(currentExcelRow, 11)    'K
noteField = Cells(currentExcelRow, 12)       'L

'Debug.Print Str(curRow) + ", planId: " + planId 'A
'Debug.Print Str(curRow) + ", planCode: " + planCode 'B
'Debug.Print Str(curRow) + ", serviceProvider: " + serviceProvider 'C
'Debug.Print Str(curRow) + ", adminFeeDB: " + Str(adminFeeDB) 'D
'Debug.Print Str(curRow) + ", assetFeeDB: " + Str(assetfeedb) 'E
'Debug.Print Str(curRow) + ", accountCount: " + Str(accountCount) 'F
'Debug.Print Str(curRow) + ", participantCount: " + Str(participantCount) 'G
'Debug.Print Str(curRow) + ", perAcctPtFee: " + Str(perAcctPtFee) 'H
'Debug.Print Str(curRow) + ", preTotal: " + Str(preTotal) 'I
'Debug.Print Str(curRow) + ", revShare: " + Str(revShare) 'J
'Debug.Print Str(curRow) + ", invoiceTotal" + Str(invoiceTotal) 'K
'Debug.Print Str(curRow) + ", noteField" + Str(noteField) 'L
sQuant = 0
sUnitP = 0

If planCode = "AA" Then
    'Debug.Print debugHead + "case AA"
    sAmnt = Str(adminFeeDB)
    sDesc = saaDesc
    'Debug.Print debugHead + "sAmnt:   " + sAmnt
    sGLcode = saaGL
    
ElseIf planCode = "AC" Then
    'Debug.Print debugHead + "case AC"
    sUnitP = Str(perAcctPtFee)
    'Debug.Print debugHead + "sUnitP:   " + sUnitP
    sQuant = Str(accountCount)
    'Debug.Print debugHead + "sQuant:   " + sQuant
    sDesc = sacDesc
    'Debug.Print debugHead + "sDesc:   " + sDesc
    sGLcode = sacGL
    
ElseIf planCode = "P" Then 'fulfillment services
    'Debug.Print debugHead + "case P"
    sQuant = Str(participantCount)
    'Debug.Print debugHead + "sQuant:   " + sQuant
    sUnitP = Str(perAcctPtFee)
    'Debug.Print debugHead + "sUnitP:   " + sUnitP
    sDesc = spDesc
    'Debug.Print debugHead + "sDesc:   " + sDesc
    sGLcode = spGL

ElseIf planCode = "A" Then
    Debug.Print debugHead + "case A"
    If serviceProvider = serviceProvider1 Then 'bluestar provider
        sDesc = sa1Desc
        'Debug.Print debugHead + "sDesc:   " + sDesc
        sAmnt = Str(assetfeedb)
        'Debug.Print debugHead + "sAmnt:   " + sAmnt
        sGLcode = sa1GL
    ElseIf serviceProvider = serviceProvider2 Then 'tdprovider
        sDesc = sa2Desc
        'Debug.Print debugHead + "sDesc:   " + sDesc
        sAmnt = Str(assetfeedb)
        'Debug.Print debugHead + "sAmnt:   " + sAmnt
        sGLcode = sa2GL
    Else
        Debug.Print debugHead + "Error on code A case: " + serviceProvider + " is not a valid provider"
        Exit Sub
    End If
End If
    
'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'

'Debug.Print debugHead + "Starting Sage50 write Operations"

AppActivate (sageWindow)
waitTime = "00:00:01"
waitTimeLong = "00:00:03"

If sQuant <> "" Then
SetClipboard (sQuant)
'Application.Wait (Now() + TimeValue(waitTime))
SendKeys "{TAB}", True
SendKeys "+{TAB}", True
Application.Wait (Now() + TimeValue(waitTime))
'Debug.Print debugHead + "Set sQuant = " + sQuant
SendKeys "^v", True
Application.Wait (Now() + TimeValue(waitTime))
End If
SendKeys ("{TAB}"), True

If sItem <> "" Then
SetClipboard ("Test")
'Application.Wait (Now() + TimeValue(waitTime))
'Debug.Print debugHead + "Set sItem = " + sItem
SendKeys "^v", True
Application.Wait (Now() + TimeValue(waitTime))
End If
SendKeys ("{TAB}"), True

If sDesc <> "" Then
SetClipboard (sDesc)
'Application.Wait (Now() + TimeValue(waitTime))
'Debug.Print debugHead + "Set sDesc = " + sDesc
SendKeys "^v", True
Application.Wait (Now() + TimeValue(waitTime))
End If
SendKeys ("{TAB}"), True

If sGLcode <> "" Then
SetClipboard (sGLcode)
'Application.Wait (Now() + TimeValue(waitTime))
'Debug.Print debugHead + "Set sDesc = " + sDesc
SendKeys "^v", True
Application.Wait (Now() + TimeValue(waitTime))
End If
SendKeys ("{TAB}"), True

If sUnitP <> "" Then
SetClipboard (sUnitP)
'Application.Wait (Now() + TimeValue(waitTime))
'Debug.Print "Set sUnitP = " + sUnitP
SendKeys "^v", True
Application.Wait (Now() + TimeValue(waitTime))
End If
SendKeys ("{TAB}"), True

If sTax <> "" Then
'SetClipboard ("Test")
'Application.Wait (Now() + TimeValue(waitTime))
'Debug.Print debugHead + "ERROR: no taxes given"
'SendKeys "^v",True
End If
SendKeys ("{TAB}"), True

If sAmnt <> "" Then
SetClipboard (sAmnt)
'Application.Wait (Now() + TimeValue(waitTime))
'Debug.Print debugHead + "Set sAmnt = " + sAmnt
SendKeys "^v", True
Application.Wait (Now() + TimeValue(waitTime))
End If

SendKeys ("{TAB}"), True

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''MOVING TO NEXT ROW IN SAGE50''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SendKeys ("{TAB}"), True

'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''''
'Debug.Print debugHead + "returning to excel..."
AppActivate (excelWindow)

Debug.Print debugHead + "finished: " + Str(curRow) + "," + planId

Cells(currentExcelRow, "M").Value = currentSageRow

currentExcelRow = currentExcelRow + 1
Rows(currentExcelRow).Select
Range("current_excel_row").Value = currentExcelRow




       
End Sub

