Attribute VB_Name = "ModFunction"
Option Explicit
''*******************************************************************************************
'' Module Name               : ModFunction
'' Description               : This module contains function to use in Report
''
''    Date                         Auther               Action              Remarks
'' 23-Oct-2020                  Mukesh Sharma         First Created
''*******************************************************************************************

''*******************************************************************************************
'' Function Name             : fnDataEntry
'' Description               : To store data into Excel sheet from User form
'' Arguments                 : NA
'' Returns                   : NA
''*******************************************************************************************

Function fnDataEntry(bSave As Boolean)
    
    'Declare variable
    Dim lLastRow As Long
    Dim lCounter As Long
    Dim lCurRow As Long
    Dim lTrackerRow As Long
    Dim rng As Range
    Dim rngTracker As Range
    Dim bExist As Boolean
    
    wksTracker.Unprotect "test@123"
       
    If bSave = True Then
        'Count last row
        lCurRow = wksData.Cells(Rows.Count, "A").End(xlUp).Row + 1
        
'        lLastRow = wksData.Cells(Rows.Count, "A").End(xlUp).Row
'        bExist = True
'        While bExist = True
'            bExist = False
'            For lCounter = 2 To lLastRow
'                If wksData.Range("A" & lCounter).Value = lCurRow Then
'                    lCurRow = lCounter + 1
'                    bExist = True
'                    Exit For
'                End If
'            Next
'        Wend
        
    Else
        lLastRow = wksData.Cells(Rows.Count, "A").End(xlUp).Row
        For lCounter = 2 To lLastRow
            If wksData.Range("A" & lCounter).Value = Helpdesk_Tracker.txtTicketNo.Value Then
                lCurRow = lCounter
                Exit For
            End If
        Next
        
    End If
        
    'Validation
    If lCurRow = 0 Then
        MsgBox "Unable to seach data in database", vbInformation
        Exit Function
    End If
        
    Set rng = wksData.Range("A" & lCurRow)
    
    lTrackerRow = wksTracker.Cells(Rows.Count, "A").End(xlUp).Row + 1
 
    
    Set rngTracker = wksTracker.Range("A" & lTrackerRow)
    
    'Store data in Excel sheet
   If Helpdesk_Tracker.Cmb_status.Value = "Open" Then
      wksData.Activate
      wksData.AutoFilterMode = False
      With rng
          
             'Ticket No
            .Offset(0, 0).Value = Format(Date, "mmm") & "_" & lCurRow - 1
            
            'Emp Id
            .Offset(0, 1).Value = Helpdesk_Tracker.txt_Id.Value
            
            'Emp Name
            .Offset(0, 2).Value = Helpdesk_Tracker.txt_Name.Value
            
            'Call Severity
            .Offset(0, 3).Value = Helpdesk_Tracker.cmb_Call_Severity.Value
            
            'Informed to
            .Offset(0, 4).Value = Helpdesk_Tracker.Cmb_Infromed_To.Value
            
             'Service
            .Offset(0, 5).Value = Helpdesk_Tracker.Cmb_Services.Value
            
             'Allign to
            .Offset(0, 6).Value = Helpdesk_Tracker.Cmb_Aligned_To.Value
            
             'Call rec.date
            .Offset(0, 7).Value = Helpdesk_Tracker.txt_CallRec_Date.Value
            
             'Status
            .Offset(0, 8).Value = Helpdesk_Tracker.Cmb_status.Value
            
              'Closing date
            .Offset(0, 9).Value = Helpdesk_Tracker.txt_Closing_Date.Value
            
            'Description
            .Offset(0, 10).Value = Helpdesk_Tracker.txtDesc.Value
            
            'Comments
            .Offset(0, 11).Value = Helpdesk_Tracker.txtComment.Value
            
            'Current user
            .Offset(0, 12).Value = VBA.Environ("UserName")
            
            'Date
            .Offset(0, 13).Value = Date
            
        End With
    
    ElseIf Helpdesk_Tracker.Cmb_status.Value = "Closed" Then
     wksTracker.Activate
      With rngTracker
          
             'Ticket No
           .Offset(0, 0).Value = Format(Date, "mmm") & "_" & lTrackerRow - 1
            
            'Emp Id
            .Offset(0, 1).Value = Helpdesk_Tracker.txt_Id.Value
            
            'Emp Name
            .Offset(0, 2).Value = Helpdesk_Tracker.txt_Name.Value
            
            'Call Severity
            .Offset(0, 3).Value = Helpdesk_Tracker.cmb_Call_Severity.Value
            
            'Informed to
            .Offset(0, 4).Value = Helpdesk_Tracker.Cmb_Infromed_To.Value
            
             'Service
            .Offset(0, 5).Value = Helpdesk_Tracker.Cmb_Services.Value
            
             'Allign to
            .Offset(0, 6).Value = Helpdesk_Tracker.Cmb_Aligned_To.Value
            
             'Call rec.date
            .Offset(0, 7).Value = Helpdesk_Tracker.txt_CallRec_Date.Value
            
             'Status
            .Offset(0, 8).Value = Helpdesk_Tracker.Cmb_status.Value
            
              'Closing date
            .Offset(0, 9).Value = Helpdesk_Tracker.txt_Closing_Date.Value
            
            'Description
            .Offset(0, 10).Value = Helpdesk_Tracker.txtDesc.Value
            
            'Comments
            .Offset(0, 11).Value = Helpdesk_Tracker.txtComment.Value
            
             'Current user
            .Offset(0, 12).Value = VBA.Environ("UserName")
            
            'Date
            .Offset(0, 13).Value = Date
            
            wksTracker.Cells.Columns.AutoFit
             
             'Delete Record
            Application.DisplayAlerts = False
            wksData.Range("A" & lCurRow).EntireRow.Delete shift:=xlToLeft
            Application.DisplayAlerts = True
  
        End With
    End If
      
       Call UploadData
      wksTracker.Protect "test@123"
      wksTracker.Activate
End Function

Sub UploadData()
    
    Dim lCurRow As Long
    Dim vArray As Variant
        
    wksData.AutoFilterMode = False
    lCurRow = wksData.Cells(Rows.Count, "A").End(xlUp).Row
      
     vArray = wksData.Range("A1:L" & lCurRow)

    Helpdesk_Tracker.lstData.RowSource = ""
    Helpdesk_Tracker.lstData.List = vArray
    
End Sub
''*******************************************************************************************
'' Procedure Name            : ShowUserForm
'' Description               : To Show User Form
'' Arguments                 : NA
'' Returns                   : NA
''*******************************************************************************************

Sub ShowUserForm()
    
 'Show user form
 Helpdesk_Tracker.Show vbModeless
 
End Sub


Sub Resetcontrols()
    
    Dim ctrl As Control
    Dim lCounter As Long
    
    For Each ctrl In Helpdesk_Tracker.Controls
        
       If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Then
            ctrl.Value = ""
        ElseIf TypeName(ctrl) = "OptionButton" Or TypeName(ctrl) = "CheckBox" Then
           ctrl.Value = False
        End If
    Next
    
    'reset list items
     For lCounter = 0 To Helpdesk_Tracker.lstData.ListCount - 1
         If Helpdesk_Tracker.lstData.Selected(lCounter) = True Then
            Helpdesk_Tracker.lstData.Selected(lCounter) = False
        End If
    Next

End Sub



