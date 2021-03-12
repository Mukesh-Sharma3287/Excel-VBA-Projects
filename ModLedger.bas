Attribute VB_Name = "ModLedger"
Option Explicit

''*******************************************************************************************
'' Function Name             : Main
'' Description               : To Create Report
'' Arguments                 : NA
'' Returns                   : NA
''*******************************************************************************************

Sub Main()
    
    Dim strProviderPath As String
    Dim strFldPath As String
    Dim wkbTemplate As Workbook
    Dim vArrayData As Variant
    Dim lLastrow As Long
    Dim lCounter As Long
    Dim wkbRpt As Workbook
    Dim wkbProviderData As Workbook
    Dim dteStartDate As Date
    Dim dteEndDate As Date
    Dim wkstemp As Worksheet
    
    strProviderPath = wksMacro.Range("C5").Value
    strFldPath = wksMacro.Range("C7").Value
    
    
    If strFldPath = "" Then
        MsgBox "Please Browse folder path to save report", vbInformation
        Exit Sub
    End If
    
    'turn off the screen,events,alert
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    On Error GoTo Err_handling:
    'open provider file
    Set wkbProviderData = Workbooks.Open(strProviderPath)
    
    lLastrow = wkbProviderData.Sheets(1).Cells(Rows.Count, "A").End(xlUp).Row
    
    vArrayData = wkbProviderData.Sheets(1).Range("A2:Z" & lLastrow)
    
    
    Set wkbTemplate = Workbooks.Add
    
    Set wkstemp = wkbTemplate.Sheets(1)
    
    'MPU
    wkstemp.Name = "MPU"
    wkstemp.Range("A1").Value = "MPU"
    wkstemp.Range("A1").Font.Bold = True
    wkstemp.Range("A1").Interior.Color = vbYellow
    wkstemp.Range("B2:E10000").ColumnWidth = 20
    wkstemp.Range("B2:E10000").HorizontalAlignment = xlCenter
      
     Set wkstemp = wkbTemplate.Sheets.Add(after:=wkbTemplate.Sheets("MPU"))
     'BBM
    wkstemp.Name = "BBM"
    wkstemp.Range("A1").Value = "BBM"
    wkstemp.Range("B2").Value = "ISM"
    wkstemp.Range("C2").Value = "FASL"
    
    wkstemp.Range("H1").Value = "BBM"
    wkstemp.Range("I2").Value = "FASL"
    wkstemp.Range("J2").Value = "ISM"
    
    wkstemp.Range("A1").Font.Bold = True
    wkstemp.Range("A1").Interior.Color = vbYellow
    wkstemp.Range("H1").Font.Bold = True
    wkstemp.Range("H1").Interior.Color = vbYellow
    
    wkstemp.Range("H2:E10000").ColumnWidth = 20
    wkstemp.Range("K2:E10000").HorizontalAlignment = xlCenter
    
    Set wkstemp = wkbTemplate.Sheets.Add(after:=wkbTemplate.Sheets("BBM"))
    
     'JDE Keying
    wkstemp.Name = "JDE"
    wkstemp.Range("A1").Value = "JDE"
    wkstemp.Range("A1").Font.Bold = True
    wkstemp.Range("A1").Interior.Color = vbYellow
    wkstemp.Range("B2:E10000").ColumnWidth = 20
    wkstemp.Range("B2:E10000").HorizontalAlignment = xlCenter
    
    dteStartDate = wksMacro.Range("C9").Value
    dteEndDate = wksMacro.Range("C11").Value
    
    For lCounter = LBound(vArrayData) To UBound(vArrayData)
        
        If lCounter = 18 Then
            Debug.Print lCounter
        End If
        
        If vArrayData(lCounter, 4) <> Empty And vArrayData(lCounter, 11) <> Empty Then
            If VBA.IsNumeric(vArrayData(lCounter, 4)) = True And VBA.IsNumeric(vArrayData(lCounter, 11)) = True And vArrayData(lCounter, 5) = Empty Then
                If CDate(vArrayData(lCounter, 2)) >= dteStartDate And CDate(vArrayData(lCounter, 2)) <= dteEndDate Then
                    Call fnCreateLedgerCode(wkbTemplate, vArrayData(lCounter, 3), vArrayData(lCounter, 4), vArrayData(lCounter, 11), vArrayData(lCounter, 18), vArrayData(lCounter, 24), vArrayData(lCounter, 7))
                End If
            End If
        
       End If
        
    Next
        
    'Copy past data FASl
    wkbTemplate.Activate
    With wkbTemplate.Sheets("BBM")
       .Activate
        lLastrow = .Cells(Rows.Count, "B").End(xlUp).Row + 2
       .Range("H1").CurrentRegion.Copy .Range("A" & lLastrow)
       .Range("H1").CurrentRegion.Delete shift:=xlToLeft
        
    End With
    
    'Save File
    wkbTemplate.SaveAs strFldPath & "\Keying Report " & Format(Now(), "dd-mmm-yyyy h.mm.ss") & ".xlsx"
   
    'turn on the screen,events,alert
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    
     MsgBox "Done!!", vbInformation
     
     Exit Sub
     
Err_handling:
     MsgBox Err.Description, vbCritical
    
End Sub


''*******************************************************************************************
'' Function Name             : Main
'' Description               : To Create Report
'' Arguments                 : NA
'' Returns                   : NA
''*******************************************************************************************

Public Function fnCreateLedgerCode(wkbTemp As Workbook, vFund As Variant, vDist As Variant, vFidValue As Variant, vProvideValue As Variant, vDiff As Variant, vRounding As Variant)
        
     Dim lCurRow As Long
      
    '1-Normal Keying
    'On Error Resume Next
    If Round(vProvideValue) = Round(vFidValue) And vDiff = 0 Then
    'On Error GoTo 0
        'MPU
        With wkbTemp.Sheets("MPU")
            .Activate
            lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row
            If lCurRow = 1 Then
                lCurRow = 3
            Else
                lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row + 1
            End If
            
            .Range("B" & lCurRow).Value = 4902.10004
            .Range("C" & lCurRow).Value = 4902.33094
            .Range("D" & lCurRow).Value = vProvideValue
            .Range("E" & lCurRow).Value = vFund
        End With
    End If
    
    If vFidValue <> Empty And vProvideValue <> Empty Then
            '2-Underpayment if yes
            If LCase(vRounding) = "yes" And vFidValue > vProvideValue Then
                 'MPU
                With wkbTemp.Sheets("MPU")
                    .Activate
                     lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row
                   If lCurRow = 1 Then
                        lCurRow = 3
                    Else
                        lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row + 1
                    End If
                    
                    .Range("B" & lCurRow).Value = 4902.10004
                    .Range("C" & lCurRow).Value = 4902.33094
                    .Range("D" & lCurRow).Value = vProvideValue
                    .Range("E" & lCurRow).Value = vFund
                    
                    .Range("B" & lCurRow + 1).Value = 4902.10004
                    .Range("C" & lCurRow + 1).Value = 4902.33094
                    .Range("D" & lCurRow + 1).Value = Abs(vDiff)
                    .Range("E" & lCurRow + 1).Value = vFund
                End With
                
                'BBM
                With wkbTemp.Sheets("BBM")
                    .Activate
                    lCurRow = .Cells(Rows.Count, "I").End(xlUp).Row + 1
                    
                    'FASL
                    .Range("I" & lCurRow).Value = 90546801
                    .Range("J" & lCurRow).Value = 60686298
                    .Range("K" & lCurRow).Value = Abs(vDiff)
                    .Range("L" & lCurRow).Value = vFund & vDist
                     
                End With
                
                'JDE
                 With wkbTemp.Sheets("JDE")
                    .Activate
                     lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row
                   If lCurRow = 1 Then
                        lCurRow = 3
                    Else
                        lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row + 1
                    End If
                    
                    .Range("B" & lCurRow).Value = 4025000.69523
                    .Range("B" & lCurRow).NumberFormat = "0.00000"
                    .Range("C" & lCurRow).Value = Abs(vDiff)
                    .Range("E" & lCurRow).Value = vFund & vDist
                    .Range("F" & lCurRow).Value = "03UMUF"
                    .Range("G" & lCurRow).Value = "C"
                    
                    .Range("B" & lCurRow + 1).Value = 402.10001
                    .Range("D" & lCurRow + 1).Value = Abs(vDiff)
                    .Range("E" & lCurRow + 1).Value = vFund & vDist
                End With
            
            End If
            
             '3-Underpayment if No/cross
            If (LCase(vRounding) = "no" Or LCase(vRounding) = "cross") And vFidValue > vProvideValue Then
                 'MPU
                With wkbTemp.Sheets("MPU")
                    .Activate
                     lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row
                   If lCurRow = 1 Then
                        lCurRow = 3
                    Else
                        lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row + 1
                    End If
                    
                    .Range("B" & lCurRow).Value = 4902.10004
                    .Range("C" & lCurRow).Value = 4902.33094
                    .Range("D" & lCurRow).Value = vProvideValue
                    .Range("E" & lCurRow).Value = vFund
                    
                    .Range("B" & lCurRow + 1).Value = 4902.10004
                    .Range("C" & lCurRow + 1).Value = 4902.33094
                    .Range("D" & lCurRow + 1).Value = Abs(vDiff)
                    .Range("E" & lCurRow + 1).Value = vFund
                End With
                
                'BBM
                With wkbTemp.Sheets("BBM")
                    .Activate
                    lCurRow = .Cells(Rows.Count, "I").End(xlUp).Row + 1
                    
                    'FASL
                    .Range("I" & lCurRow).Value = 90546801
                    .Range("J" & lCurRow).Value = 60686298
                    .Range("K" & lCurRow).Value = Abs(vDiff)
                    .Range("L" & lCurRow).Value = vFund & vDist
                     
                End With
                
                'JDE
                 With wkbTemp.Sheets("JDE")
                    .Activate
                     lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row
                   If lCurRow = 1 Then
                        lCurRow = 3
                    Else
                        lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row + 1
                    End If
                    
                    .Range("B" & lCurRow).Value = 402.33094
                    .Range("C" & lCurRow).Value = Abs(vDiff)
                    .Range("E" & lCurRow).Value = vFund & vDist
                    
                    .Range("B" & lCurRow + 1).Value = 402.10001
                    .Range("D" & lCurRow + 1).Value = Abs(vDiff)
                    .Range("E" & lCurRow + 1).Value = vFund & vDist
                End With
            
            End If
            
                
            '4-OverPayment if yes
            If vProvideValue > vFidValue And LCase(vRounding) = "yes" Then
                 'MPU
                With wkbTemp.Sheets("MPU")
                    .Activate
                     lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row
                    If lCurRow = 1 Then
                        lCurRow = 3
                    Else
                        lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row + 1
                    End If
                    
                    .Range("B" & lCurRow).Value = 4902.10004
                    .Range("C" & lCurRow).Value = 4902.33094
                    .Range("D" & lCurRow).Value = vProvideValue
                    .Range("E" & lCurRow).Value = vFund
                    
                    .Range("B" & lCurRow + 1).Value = 4902.33094
                    .Range("C" & lCurRow + 1).Value = 4902.10004
                    .Range("D" & lCurRow + 1).Value = Abs(vDiff)
                    .Range("E" & lCurRow + 1).Value = vFund
                End With
                
                'BBM
                With wkbTemp.Sheets("BBM")
                    .Activate
                    lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row + 1
                    
                    'ISM
                    .Range("B" & lCurRow).Value = 60686298
                    .Range("C" & lCurRow).Value = 90546801
                    .Range("D" & lCurRow).Value = Abs(vDiff)
                    .Range("E" & lCurRow).Value = vFund & vDist
                     
                End With
                
                'JDE
                 With wkbTemp.Sheets("JDE")
                    .Activate
                    lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row
                    If lCurRow = 1 Then
                        lCurRow = 3
                    Else
                        lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row + 1
                    End If
                    
                    .Range("B" & lCurRow).Value = 402.10001
                    .Range("C" & lCurRow).Value = Abs(vDiff)
                    .Range("E" & lCurRow).Value = vFund & vDist
                   
                    
                    .Range("B" & lCurRow + 1).Value = 4025000.69523
                    .Range("D" & lCurRow + 1).Value = Abs(vDiff)
                    .Range("E" & lCurRow + 1).Value = vFund & vDist
                    .Range("F" & lCurRow + 1).Value = "03UMUF"
                    .Range("G" & lCurRow + 1).Value = "C"
                End With
            
            End If
            
            '5- Overpayment if no
            If vProvideValue > vFidValue And LCase(vRounding) = "no" Or LCase(vRounding) = "cross" Then
                
                'MPU
                With wkbTemp.Sheets("MPU")
                    .Activate
                     lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row
                    If lCurRow = 1 Then
                        lCurRow = 3
                    Else
                        lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row + 1
                    End If
                    
                    .Range("B" & lCurRow).Value = 4902.10004
                    .Range("C" & lCurRow).Value = 4902.33094
                    .Range("D" & lCurRow).Value = vProvideValue
                    .Range("E" & lCurRow).Value = vFund
                    
                    .Range("B" & lCurRow + 1).Value = 4902.33094
                    .Range("C" & lCurRow + 1).Value = "4902.33099.DS"
                    .Range("D" & lCurRow + 1).Value = Abs(vDiff)
                    .Range("E" & lCurRow + 1).Value = vFund
                End With
            End If
        End If
    
    '6-Payment not received
    If (vProvideValue = 0 And vFidValue > 0) Then
         'MPU
        With wkbTemp.Sheets("MPU")
            .Activate
             lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row
            If lCurRow = 1 Then
                lCurRow = 3
            Else
                lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row + 1
            End If
            
            .Range("B" & lCurRow).Value = 4902.10004
            .Range("C" & lCurRow).Value = 4902.33094
            .Range("D" & lCurRow).Value = vFidValue
            .Range("E" & lCurRow).Value = vFund
        End With
        
        'BBM
        With wkbTemp.Sheets("BBM")
            .Activate
             lCurRow = .Cells(Rows.Count, "I").End(xlUp).Row + 1
            
            'FASL
            .Range("I" & lCurRow).Value = 90546801
            .Range("J" & lCurRow).Value = 60686298
            .Range("K" & lCurRow).Value = vFidValue
            .Range("L" & lCurRow).Value = vFund & vDist
             
        End With
        
        'JDE
         With wkbTemp.Sheets("JDE")
            .Activate
            lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row
            If lCurRow = 1 Then
                lCurRow = 3
            Else
                lCurRow = .Cells(Rows.Count, "B").End(xlUp).Row + 1
            End If
            
            .Range("B" & lCurRow).Value = 402.33094
            .Range("C" & lCurRow).Value = vFidValue
            .Range("E" & lCurRow).Value = vFund & vDist
            
            .Range("B" & lCurRow + 1).Value = 402.10001
            .Range("D" & lCurRow + 1).Value = vFidValue
            .Range("E" & lCurRow + 1).Value = vFund & vDist
        End With
    
    End If

End Function



''*******************************************************************************************
'' Function Name             : cmdBrowseRawPath_Click
'' Description               : To Browse the file path
'' Arguments                 : NA
'' Returns                   : NA
''*******************************************************************************************

Sub BrowseRawPath2()
    
    Dim objFileDialog As FileDialog
    Dim vSelected As Variant
    
    'Show dialog
    Set objFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With objFileDialog
        .AllowMultiSelect = False
        .ButtonName = "File Path"
        .Filters.Add "Excel File", "*.xlsx,*.xls"
        .Show
        
        'Assign the selected file into text box
        For Each vSelected In .SelectedItems
            
           wksMacro.Range("C5").Value = vSelected
            
        Next
        
    End With
    
End Sub


''*******************************************************************************************
'' Function Name             : FolderPath_Browse
'' Description               : To Browse the Folder path
'' Arguments                 : NA
'' Returns                   : NA
''*******************************************************************************************

Sub FolderPath_Browse()
    
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
            
           wksMacro.Range("C7").Value = vSelected
            
        Next
        
    End With
    
End Sub

''****************************************************************************************
'' Procedure          : fnValidPath
'' Description        : This function will validate whether selected file path is valid or not
''
'' Arguments          : strPath As String
'' Return             : Boolean- If selected file path is valid then it's return true else false
''****************************************************************************************

Public Function fnValidPath(strPath As String) As Boolean
    
    'Check whether file path is blank or not
    If strPath = "" Then
        fnValidPath = False
        MsgBox "Please Browse Raw file path to proceed", vbInformation
        Exit Function
    End If
    
    'Check whether file path is exist or not
    If Dir(strPath, vbDirectory) = "" Then
        fnValidPath = False
        MsgBox "Raw file does not exit to proceed", vbInformation
        Exit Function
    Else
        fnValidPath = True
    End If
    
End Function


