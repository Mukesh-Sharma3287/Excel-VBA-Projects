Attribute VB_Name = "ModCalendar"
Option Explicit


Sub MainCalendar()
    
    'Declare variables
    Dim lCounter As Long
    Dim lInnerCounter As Long
    Dim lLastCurRow As Long
    Dim lLastRow As Long
    Dim rngHoliDay As Range
    Dim iCounter As Integer
    Dim vRangeProvider As Variant
    
    'off screen and alert
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
    On Error GoTo Err_handling
    lLastRow = WksProvider.Cells(Rows.Count, "A").End(xlUp).Row
    lLastCurRow = WksCalendar.Cells(Rows.Count, "C").End(xlUp).Row
    
    'Apply formula
    WksProvider.Range("H2:H" & lLastRow).Formula = "=DATE(LEFT($A2,4),MID($A2,5,2),RIGHT($A2,2))"
    WksProvider.Range("H2:H" & lLastRow).Value = WksProvider.Range("H2:H" & lLastRow).Value
    
    'Create range
    vRangeProvider = WksProvider.Range("C2:H" & lLastRow)
    
    'Clear contents in Calendars
    WksCalendar.Range("C1:AN" & lLastCurRow).CurrentRegion.Offset(4, 0).ClearContents
     
    'Set holiday range for 1 year
    For iCounter = 1 To 35
      If WksHoliday.Cells(1, iCounter).Value = Year(Date) Then
        Set rngHoliDay = WksHoliday.Range(WksHoliday.Cells(3, iCounter), WksHoliday.Cells(10, iCounter))
        Exit For
      End If
    Next
    
    lLastCurRow = WksCalendar.Cells(Rows.Count, "C").End(xlUp).Row + 1
    For lCounter = LBound(vRangeProvider) To UBound(vRangeProvider)
        
        Call fnIncAccCalendar(lLastCurRow, VBA.CDate(vRangeProvider(lCounter, 6)), vRangeProvider(lCounter, 5), vRangeProvider(lCounter, 1), rngHoliDay)
    
        lLastCurRow = lLastCurRow + 1
    Next
      
    'Save Calendar file in Folder path
    WksCalendar.Copy
    ActiveWorkbook.SaveAs WksMacro.Range("E3").Value & "\Calendar " & Format(WksProvider.Range("H2").Value, "mmm-yy") & ".xlsx"
    ActiveWorkbook.Close True
    
    'Clear date column data
    WksProvider.Range("H:H").ClearContents
    
    'On screen and alert
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
    MsgBox "Done!!", vbInformation
    On Error GoTo 0
    Exit Sub
    
Err_handling:
    MsgBox Err.Description, vbCritical
    

End Sub

''*******************************************************************************************
'' Function Name             : fnIncCal
'' Description               : To create Inc calendar
'' Arguments                 : NA
'' Returns                   : NA
''*******************************************************************************************

Function fnIncAccCalendar(lLastCurRow As Long, dtDiaryDate As Date, vIncAcc As Variant, vFundCode As Variant, rngHoliDay As Range)
    
    Dim iCounter As Integer
    Dim iCol As Integer
    Dim iColMov As Integer
    Dim bDateOut As Boolean
    Dim iHoliDay As Integer
    Dim rng As Range
      
    On Error Resume Next
    For iCounter = 5 To 41
        If dtDiaryDate = WksCalendar.Cells(4, iCounter).Value Then
            iCol = iCounter
            Exit For
        End If
    Next
        
    iHoliDay = 0
    For Each rng In rngHoliDay
        If dtDiaryDate = rng.Value Then
            iHoliDay = iCounter
            Exit For
        End If
    Next
    
   If VBA.LCase(vIncAcc) = "inc" Then
        
        'Fund Code
        WksCalendar.Cells(lLastCurRow, 3).Value = vFundCode
        
        'Income/Acc
        WksCalendar.Cells(lLastCurRow, 4).Value = vIncAcc
   
        'PAY DATE CONFIRM PAYMENTS
        If Format(dtDiaryDate, "ddd") = "Sat" Or Format(dtDiaryDate, "ddd") = "Sun" Then
            iColMov = fnMoveCol(dtDiaryDate, -1, rngHoliDay)
            WksCalendar.Cells(lLastCurRow, iColMov).Value = "PAY DATE CONFIRM PAYMENTS"
            dtDiaryDate = Application.WorksheetFunction.WorkDay_Intl(dtDiaryDate, -1, 1, rngHoliDay)
        Else
            If iHoliDay > 0 Then
                iColMov = fnMoveCol(dtDiaryDate, -1, rngHoliDay)
                WksCalendar.Cells(lLastCurRow, iColMov).Value = "PAY DATE CONFIRM PAYMENTS"
                dtDiaryDate = Application.WorksheetFunction.WorkDay_Intl(dtDiaryDate, -1, 1, rngHoliDay)
            Else
                WksCalendar.Cells(lLastCurRow, iCol).Value = "PAY DATE CONFIRM PAYMENTS"
            End If
        End If
        
        'Pay Date+1
        iColMov = fnMoveCol(dtDiaryDate, 1, rngHoliDay)
        WksCalendar.Cells(lLastCurRow, iColMov).Value = "PAY DATE + 1"
       
        'PD-2 Z20, BACS and Tax Vouchers
        iColMov = fnMoveCol(dtDiaryDate, -2, rngHoliDay)
        WksCalendar.Cells(lLastCurRow, iColMov).Value = "Z20, BACS and Tax Vouchers"
        
        'PD-5 JPM instruction
        iColMov = fnMoveCol(dtDiaryDate, -5, rngHoliDay)
        WksCalendar.Cells(lLastCurRow, iColMov).Value = "JPM instruction"
        
        'PD-10 Rates, Allocation Run and  PR/RR
        bDateOut = False
       If VBA.Month(Application.WorksheetFunction.WorkDay_Intl(dtDiaryDate, -10, 1, rngHoliDay)) = Month(dtDiaryDate) Then
         iColMov = fnMoveCol(dtDiaryDate, -10, rngHoliDay)
         WksCalendar.Cells(lLastCurRow, iColMov).Value = "Rates, Allocation Run and  PR/RR"
       Else
         If Format(VBA.DateSerial(Year(dtDiaryDate), Month(dtDiaryDate), 1), "ddd") = "Sat" Or Format(VBA.DateSerial(Year(dtDiaryDate), Month(dtDiaryDate), 1), "ddd") = "Sun" Then
            iColMov = fnMoveCol(VBA.DateSerial(Year(dtDiaryDate), Month(dtDiaryDate), 1), 1, rngHoliDay)
            WksCalendar.Cells(lLastCurRow, iColMov).Value = "Rates, Allocation Run and  PR/RR"
            dtDiaryDate = Application.WorksheetFunction.WorkDay_Intl(VBA.DateSerial(Year(dtDiaryDate), Month(dtDiaryDate), 1), 1, 1, rngHoliDay)
            bDateOut = True
         Else
            
            If Application.WorksheetFunction.NetworkDays(WksCalendar.Range("E4").Value, WksCalendar.Range("E4").Value, rngHoliDay) = 0 Then
                iColMov = fnMoveCol(WksCalendar.Range("E4").Value, 1, rngHoliDay)
                WksCalendar.Cells(lLastCurRow, iColMov).Value = "Rates, Allocation Run and PR/RR"
                dtDiaryDate = Application.WorksheetFunction.WorkDay_Intl(WksCalendar.Range("E4").Value, 1, 1, rngHoliDay)
                 bDateOut = True
            Else
                WksCalendar.Cells(lLastCurRow, 5).Value = "Rates, Allocation Run and PR/RR"
            End If
           
         End If
      End If
        
    'PD-9 Send Rates and Differential Deals
    If bDateOut = False Then
        iColMov = fnMoveCol(dtDiaryDate, -9, rngHoliDay)
        WksCalendar.Cells(lLastCurRow, iColMov).Value = "Send Rates and Differential Deals"
    ElseIf bDateOut = True Then
        iColMov = fnMoveCol(dtDiaryDate, 1, rngHoliDay)
        WksCalendar.Cells(lLastCurRow, iColMov).Value = "Send Rates and Differential Deals"
    End If
    
        
    ElseIf VBA.LCase(vIncAcc) = "acc" Then
        
         'Fund Code
        WksCalendar.Cells(lLastCurRow, 3).Value = vFundCode
        
        'Income/Acc
        WksCalendar.Cells(lLastCurRow, 4).Value = vIncAcc
        
         'PAY DATE CONFIRM PAYMENTS
        If Format(dtDiaryDate, "ddd") = "Sat" Or Format(dtDiaryDate, "ddd") = "Sun" Then
            iColMov = fnMoveCol(dtDiaryDate, -1, rngHoliDay)
            WksCalendar.Cells(lLastCurRow, iColMov).Value = "PAY DATE CONFIRM PAYMENTS"
            dtDiaryDate = Application.WorksheetFunction.WorkDay_Intl(dtDiaryDate, -1, 1, rngHoliDay)
        Else
            If iHoliDay > 0 Then
                iColMov = fnMoveCol(dtDiaryDate, -1, rngHoliDay)
                WksCalendar.Cells(lLastCurRow, iColMov).Value = "PAY DATE CONFIRM PAYMENTS"
                dtDiaryDate = Application.WorksheetFunction.WorkDay_Intl(dtDiaryDate, -1, 1, rngHoliDay)
            Else
                WksCalendar.Cells(lLastCurRow, iCol).Value = "PAY DATE CONFIRM PAYMENTS"
            End If
        End If
        
        'Pay Date+1
        iColMov = fnMoveCol(dtDiaryDate, 1, rngHoliDay)
        WksCalendar.Cells(lLastCurRow, iColMov).Value = "PAY DATE + 1"
       
        'PD-2 Tax Vouchers
        iColMov = fnMoveCol(dtDiaryDate, -2, rngHoliDay)
        WksCalendar.Cells(lLastCurRow, iColMov).Value = "Tax Vouchers"
        
    
     'PD-10 Rates, Allocation Run and  PR/RR
     bDateOut = False
      If VBA.Month(Application.WorksheetFunction.WorkDay_Intl(dtDiaryDate, -10, 1, rngHoliDay)) = Month(dtDiaryDate) Then
         iColMov = fnMoveCol(dtDiaryDate, -10, rngHoliDay)
         WksCalendar.Cells(lLastCurRow, iColMov).Value = "Rates, Allocation Run and  PR/RR"
      Else
         If Format(VBA.DateSerial(Year(dtDiaryDate), Month(dtDiaryDate), 1), "ddd") = "Sat" Or Format(VBA.DateSerial(Year(dtDiaryDate), Month(dtDiaryDate), 1), "ddd") = "Sun" Then
            iColMov = fnMoveCol(dtDiaryDate, 1, rngHoliDay)
            WksCalendar.Cells(lLastCurRow, iColMov).Value = "Rates, Allocation Run and  PR/RR"
            dtDiaryDate = Application.WorksheetFunction.WorkDay_Intl(dtDiaryDate, 1, 1, rngHoliDay)
            bDateOut = True
         Else
            
            If Application.WorksheetFunction.NetworkDays(WksCalendar.Range("E4").Value, WksCalendar.Range("E4").Value, rngHoliDay) = 0 Then
                iColMov = fnMoveCol(WksCalendar.Range("E4").Value, 1, rngHoliDay)
                WksCalendar.Cells(lLastCurRow, iColMov).Value = "Rates, Allocation Run and  PR/RR"
                dtDiaryDate = Application.WorksheetFunction.WorkDay_Intl(WksCalendar.Range("E4").Value, 1, 1, rngHoliDay)
                bDateOut = True
            Else
                WksCalendar.Cells(lLastCurRow, 5).Value = "Rates, Allocation Run and PR/RR"
            End If
            
         End If
     End If
        
     'PD-9 Send Rates
    If bDateOut = False Then
        iColMov = fnMoveCol(dtDiaryDate, -9, rngHoliDay)
        WksCalendar.Cells(lLastCurRow, iColMov).Value = "Send Rates"
    ElseIf bDateOut = True Then
        iColMov = fnMoveCol(dtDiaryDate, 1, rngHoliDay)
        WksCalendar.Cells(lLastCurRow, iColMov).Value = "Send Rates"
    End If
    
End If

    On Error GoTo 0
    
End Function

''*******************************************************************************************
'' Function Name             : fnMoveCol
'' Description               : To return column no
'' Arguments                 : NA
'' Returns                   : NA
''*******************************************************************************************
Function fnMoveCol(dtDiaryDate As Date, iMov As Integer, rngHoliDay As Range) As Integer
    
    Dim iCounter As Integer
    Dim dtDateCal As Date
    
    dtDateCal = Application.WorksheetFunction.WorkDay_Intl(dtDiaryDate, iMov, 1, rngHoliDay)
    For iCounter = 5 To 41
        If dtDateCal = WksCalendar.Cells(4, iCounter).Value Then
            fnMoveCol = iCounter
            Exit For
        End If
    Next
    
End Function

''*******************************************************************************************
'' Function Name             : ClearProviderData
'' Description               : Clear contents
'' Arguments                 : NA
'' Returns                   : NA
''*******************************************************************************************

Sub ClearProviderData()

    WksProvider.Range("A1").CurrentRegion.Offset(1, 0).ClearContents

End Sub

Sub FolderBrowser()
    
    Dim objFileDialog As FileDialog
    Dim vSelected As Variant
    
    'Show dialog
    Set objFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    With objFileDialog
        .AllowMultiSelect = False
        .ButtonName = "Folder Path"
        .Show
        
        'Assign the selected folder path into text box
        For Each vSelected In .SelectedItems
            
           WksMacro.Range("E3").Value = vSelected
            
        Next
   End With
    
End Sub

