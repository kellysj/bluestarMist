Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub NextInvoice()
' NextInvoice Macro
' Keyboard Shortcut: Ctrl+t
' Advances to next invoice, saves current invoice, selects 1st field (quantity)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'''''DECLARATIONS'''''DECLARATIONS'''''DECLARATIONS'''''DECLARATIONS'''''DECLARATIONS'''''DECLARATIONS'''''DECLARATIONS'''''DECLARATIONS'''''''''
'window titles
sageWindow1 = "Sales Invoice List"
sageWindow2 = "Sales/Invoicing (BLRDCB1.capitalsg.local)"
excelWindow = "PeachtreeDirectBill_v2_macros.xlsm"

waitTime1 = "00:00:01"
waitTime2 = "00:00:05"

'macro name
functionName = "NextInvoice"

'important indicies
currentExcelRow = ActiveCell.Row
currentSageRow = Range("sage_50_row_number").Value



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If currentExcelRow = 2 Then 'Catch start on first excel entry
    Range("sage_50_row_number").Value = 1
    currentSageRow = 1
End If
debugHead = Trim(Str(currentSageRow)) + "," + Trim(Str(currentExcelRow)) + ":" + functionName + ">    "
Debug.Print "--------------------------------------------------------------"
Debug.Print debugHead + "starting..."



'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'
AppActivate (sageWindow1), True
Application.Wait (Now() + TimeValue(waitTime1))
SendKeys "{HOME}", True
If currentExcelRow > 2 Then
    For i = 1 To currentSageRow - 1
        SendKeys "{DOWN}", True
    Next i
End If

SendKeys "{ENTER}", True
Application.Wait (Now() + TimeValue(waitTime1))
Application.Wait (Now() + TimeValue(waitTime1))
Application.Wait (Now() + TimeValue(waitTime1))
Application.Wait (Now() + TimeValue(waitTime1))

Application.Wait (Now() + TimeValue(waitTime1))
Application.Wait (Now() + TimeValue(waitTime1))
Application.Wait (Now() + TimeValue(waitTime1))
Application.Wait (Now() + TimeValue(waitTime1))


'AppActivate (sageWindow2)
Application.Wait (Now() + TimeValue(waitTime1))
For i = 1 To 19
    SendKeys "{TAB}", True
Next i

'Debug.Print debugHead + "Tabbed to Quantity Field"
    
'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''''
'Debug.Print debugHead + "returning to excel..."
Application.Wait (Now() + TimeValue(waitTime1))
AppActivate (excelWindow), True


Debug.Print debugHead + "finished"

End Sub

Sub SaveClosInv()
'
' Save_Close_Current_Invoice Macro
' Saves and Closes the current open invoice
'
' Keyboard Shortcut: Ctrl+r
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'''''DECLARATIONS'''''DECLARATIONS'''''DECLARATIONS'''''DECLARATIONS'''''DECLARATIONS'''''DECLARATIONS'''''DECLARATIONS'''''DECLARATIONS'''''''''


'window titles
sageWindow = "Sales/Invoicing (BLRDCB1.capitalsg.local)"
excelWindow = "PeachtreeDirectBill_v2_macros.xlsm"

waitTime = "00:00:01"
waitTimeLong = "00:00:03"

'macro name
functionName = "SaveClosInv"

'important indicies
currentExcelRow = ActiveCell.Row
currentSageRow = Range("sage_50_row_number").Value
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If currentExcelRow = 2 Then
    Range("sage_50_row_number").Value = 1
End If
debugHead = Trim(Str(currentSageRow)) + "," + Trim(Str(currentExcelRow)) + ":" + functionName + ">    "
Debug.Print debugHead + "starting..."

'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'''''RETURNING TO SAGE50'
AppActivate (sageWindow)

Application.Wait (Now() + TimeValue(waitTime))
Application.Wait (Now() + TimeValue(waitTime))
SendKeys "%{S}", True    'saving dialogue kotkey
Application.Wait (Now() + TimeValue(waitTimeLong))
SendKeys "{ENTER}" 'warning about credit
SendKeys "{DOWN}" 'change to apply to all future transactions
Application.Wait (Now() + TimeValue(waitTime))
SendKeys "{ENTER}" 'accept change
SendKeys "{ENTER}" 'accept again for some reason...

Application.Wait (Now() + TimeValue(waitTime))
SendKeys "{ESC}" 'leave the save window dialogue


'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''RETURNING TO EXCEL'''''''
'Debug.Print debugHead + "returning to excel..."
AppActivate (excelWindow)

Range("sage_50_row_number").Value = Range("sage_50_row_number").Value + 1
Debug.Print debugHead + "finished"

End Sub

Sub incrementSageInd()
'A macro to increment or decrement the sage index variable
'ctrl + i
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'macro name
functionName = "incrementSageInd"
currentSageRow = Range("sage_50_row_number")
currentExcelRow = ActiveCell.Row
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

debugHead = Trim(Str(currentSageRow)) + "," + Trim(Str(currentExcelRow)) + ":" + functionName + ">    "
Debug.Print debugHead + "starting..."
Range("sage_50_row_number").Value = currentSageRow + 1
currentSageRow = Range("sage_50_row_number")
Debug.Print debugHead + "sage index: " + Str(currentSageRow)
debugHead = Trim(Str(currentSageRow)) + "," + Trim(Str(currentExcelRow)) + ":" + functionName + ">    "
Debug.Print debugHead + "finished"

End Sub

Sub decrementSageInd()
'A macro to decrement or decrement the sage index variable
'ctrl + u
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'macro name
functionName = "decrementSageInd"
currentSageRow = Range("sage_50_row_number")
currentExcelRow = ActiveCell.Row
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

debugHead = Trim(Str(currentSageRow)) + "," + Trim(Str(currentExcelRow)) + ":" + functionName + ">    "
Debug.Print debugHead + "starting..."
Range("sage_50_row_number").Value = currentSageRow - 1
currentSageRow = Range("sage_50_row_number")
Debug.Print debugHead + "sage index: " + Str(currentSageRow)
debugHead = Trim(Str(currentSageRow)) + "," + Trim(Str(currentExcelRow)) + ":" + functionName + ">    "
Debug.Print debugHead + "finished"

End Sub

Sub noteForfeit()
'
'
'
''''''''''''''''''''''''
'macro name
functionName = "noteForfeit"
currentSageRow = Range("sage_50_row_number")
currentExcelRow = ActiveCell.Row
messageStr = "Forfeiture Applied Item"
''''''''''''''''''''''''
debugHead = Trim(Str(currentSageRow)) + "," + Trim(Str(currentExcelRow)) + ":" + functionName + ">    "
Debug.Print debugHead + "starting..."
Cells(currentExcelRow, "L").Value = messageStr


Debug.Print debugHead + "finished"
End Sub

