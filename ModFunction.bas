Attribute VB_Name = "ModFunction"
Option Explicit
    
    Sub Browse1()
    
        Dim objFileDialog As FileDialog
        Dim vSelected As Variant
        
        'Show dialog
        Set objFileDialog = Application.FileDialog(msoFileDialogFilePicker)
        
        With objFileDialog
            .AllowMultiSelect = False
            .ButtonName = "File Path"
            .Filters.Add "Excel File", "*.xlsx,*.xls,*.xlsm"
            .Show
            
            'Assign the selected file into text box
            For Each vSelected In .SelectedItems
                
               wksHome.Range("C4").Value = vSelected
                
            Next
            
        End With
    End Sub

 Sub Browse2()
    
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
                
               wksHome.Range("C5").Value = vSelected
                
            Next
            
        End With
    End Sub
    
     Sub Browse3()
    
        Dim objFileDialog As FileDialog
        Dim vSelected As Variant
        
        'Show dialog
        Set objFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
        
        With objFileDialog
            .AllowMultiSelect = False
            .ButtonName = "File Path"
            .Show
            
            'Assign the selected file into text box
            For Each vSelected In .SelectedItems
                
                wksHome.Range("C6").Value = vSelected
                
            Next
            
        End With
    End Sub
