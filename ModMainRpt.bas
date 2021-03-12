Attribute VB_Name = "ModMainRpt"
Option Explicit

Sub MainRpt()
    
    'Declare variables
    Dim wkbRaw As Workbook
    Dim wkbTemp As Workbook
    Dim wksRaw As Worksheet
    Dim wksTemp As Worksheet
    Dim lLastRow As Long
    Dim rng As Range
    Dim lCounter As Long
    Dim iLastCol As Integer
    
    'turn offf alert and screen update
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
    
    'Open Files
    Set wkbRaw = Workbooks.Open(wksHome.Range("C4").Value, False, False)
    Set wksRaw = wkbRaw.Sheets(1)
    
    'Formula
    lLastRow = wksRaw.Cells(Rows.Count, "A").End(xlUp).Row
    iLastCol = wksRaw.Cells(2, Columns.Count).End(xlToLeft).Column

    wksRaw.Cells(1, iLastCol).Value = "Fatals_ Count"
    wksRaw.Range(Cells(2, iLastCol), Cells(lLastRow, iLastCol)).Formula = "=IF($J2>0,1,0)"
    
    
    Set wkbTemp = Workbooks.Open(wksHome.Range("C5").Value, False, False)
    Set wksTemp = wkbTemp.Sheets("Raw Data")
    
    wksTemp.Range("A1").CurrentRegion.Offset(1, 0).ClearContents
    
    Set rng = wksRaw.Range("A1").CurrentRegion
    rng.Copy
    wksTemp.Range("A1").PasteSpecial (xlPasteAll)
    
    'Formula
    
    
    'Close Raw data file
    wkbRaw.Close False
     
    wkbTemp.RefreshAll
    
    'Paste data Agent wise
    Set rng = wkbTemp.Sheets("Pivot Table").Range("A3").CurrentRegion.Offset(2, 0)
    rng.Copy
    With wkbTemp.Sheets("Agent Wise")
        .Activate
        .Range("A2").PasteSpecial xlPasteValues
         lLastRow = .Cells(Rows.Count, "A").End(xlUp).Row
        .Range("E2:E" & lLastRow).Formula = "=$D2/$B2"
        .Range("A2:E2").Copy
        .Range("A3:E" & lLastRow).PasteSpecial xlPasteFormats
        .Range("A" & lLastRow & ":E" & lLastRow).Interior.ColorIndex = 55
        .Range("A" & lLastRow & ":E" & lLastRow).Font.Color = vbWhite
    End With
    
    'Paste data Date wise
    Set rng = wkbTemp.Sheets("Pivot Table").Range("G3").CurrentRegion.Offset(2, 0)
    rng.Copy
    With wkbTemp.Sheets("Date Wise")
        .Activate
        .Range("A2").PasteSpecial xlPasteValues
         lLastRow = .Cells(Rows.Count, "A").End(xlUp).Row
        .Range("E2:E" & lLastRow).Formula = "=$D2/$B2"
        .Range("A2:E2").Copy
        .Range("A3:E" & lLastRow).PasteSpecial xlPasteFormats
        .Range("A" & lLastRow & ":E" & lLastRow).Interior.ColorIndex = 55
        .Range("A" & lLastRow & ":E" & lLastRow).Font.Color = vbWhite
    End With
    
    'Paste data week wise
    Set rng = wkbTemp.Sheets("Pivot Table").Range("N3").CurrentRegion.Offset(2, 0)
    rng.Copy
    With wkbTemp.Sheets("Week Wise")
        .Activate
        .Range("A2").PasteSpecial xlPasteValues
         lLastRow = .Cells(Rows.Count, "A").End(xlUp).Row
        .Range("E2:E" & lLastRow).Formula = "=$D2/$B2"
        .Range("A2:E2").Copy
        .Range("A3:E" & lLastRow).PasteSpecial xlPasteFormats
        .Range("A" & lLastRow & ":E" & lLastRow).Interior.ColorIndex = 55
        .Range("A" & lLastRow & ":E" & lLastRow).Font.Color = vbWhite
    End With
    
    'Paste data Tenure wise
    Set rng = wkbTemp.Sheets("Pivot Table").Range("U3").CurrentRegion.Offset(2, 0)
    rng.Copy
    With wkbTemp.Sheets("Tenure Wise")
        .Activate
        .Range("A2").PasteSpecial xlPasteValues
         lLastRow = .Cells(Rows.Count, "A").End(xlUp).Row
        .Range("E2:E" & lLastRow).Formula = "=$D2/$B2"
        .Range("A2:E2").Copy
        .Range("A3:E" & lLastRow).PasteSpecial xlPasteFormats
        .Range("A" & lLastRow & ":E" & lLastRow).Interior.ColorIndex = 55
        .Range("A" & lLastRow & ":E" & lLastRow).Font.Color = vbWhite
    End With
    
    'Paste data TL Wise
    Set rng = wkbTemp.Sheets("Pivot Table").Range("AB3").CurrentRegion.Offset(2, 0)
    rng.Copy
    With wkbTemp.Sheets("TL Wise")
        .Activate
        .Range("A2").PasteSpecial xlPasteValues
         lLastRow = .Cells(Rows.Count, "A").End(xlUp).Row
        .Range("E2:E" & lLastRow).Formula = "=$D2/$B2"
        .Range("A2:E2").Copy
        .Range("A3:E" & lLastRow).PasteSpecial xlPasteFormats
        .Range("A" & lLastRow & ":E" & lLastRow).Interior.ColorIndex = 55
        .Range("A" & lLastRow & ":E" & lLastRow).Font.Color = vbWhite
    End With
    
    'QA Wise
    Set rng = wkbTemp.Sheets("Pivot Table").Range("AI3").CurrentRegion.Offset(2, 0)
    rng.Copy
    With wkbTemp.Sheets("QA Wise")
        .Activate
        .Range("A2").PasteSpecial xlPasteValues
         lLastRow = .Cells(Rows.Count, "A").End(xlUp).Row
        .Range("E2:E" & lLastRow).Formula = "=$D2/$B2"
        .Range("A2:E2").Copy
        .Range("A3:E" & lLastRow).PasteSpecial xlPasteFormats
        .Range("A" & lLastRow & ":E" & lLastRow).Interior.ColorIndex = 55
        .Range("A" & lLastRow & ":E" & lLastRow).Font.Color = vbWhite
    End With

    Application.CutCopyMode = False
    
    wkbTemp.RefreshAll
     
    wkbTemp.SaveAs wksHome.Range("C6").Value & "\Dashboard" & Format(Now(), "dd-mmm-yyyy") & ".xlsx"
    wkbTemp.Close False
    
     'turn offf alert and screen update
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
    
    MsgBox "Done!!", vbInformation
    
End Sub


Public Function fnAgentwiseRpt(wksPvt As Worksheet, wksRpt As Worksheet, wksChart As Worksheet)
    
    'Declare variables
    Dim lCounter As Long
    Dim iLastRow As Integer
    
    wksRpt.UsedRange.ClearContents

    wksPvt.Range("A3").CurrentRegion.Copy

    wksRpt.Activate
    wksRpt.Range("A1").PasteSpecial (xlPasteValues)
    wksRpt.Range("A1").PasteSpecial (xlPasteFormats)


    For lCounter = 2 To 32
        
       If Trim(wksRpt.Cells(3, lCounter).Value) = "Total" Or Trim(wksRpt.Cells(3, lCounter).Value) = "Fatal Count" Then
            wksRpt.Columns(lCounter).Delete shift:=xlToLeft
            lCounter = lCounter - 1
        End If
    Next
     
     'To create
     iLastRow = wksRpt.Cells(Rows.Count, "A").End(xlUp).Row
     wksRpt.Range("B" & iLastRow & ":D" & iLastRow).Copy wksChart.Range("B3")
     wksRpt.Range("E" & iLastRow & ":G" & iLastRow).Copy wksChart.Range("B4")
     wksRpt.Range("H" & iLastRow & ":J" & iLastRow).Copy wksChart.Range("B5")
     wksRpt.Range("K" & iLastRow & ":M" & iLastRow).Copy wksChart.Range("B6")
     wksRpt.Range("N" & iLastRow & ":P" & iLastRow).Copy wksChart.Range("B7")
     
    
End Function

