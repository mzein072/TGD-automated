Dim CB_Val As Variant
Private Sub ComboBox1_change()
'DROP-DOWN USE CATEGORY BOX ///////////////////////////////////////////////////////

Worksheets("PEC Calculator").Unprotect "eau"

If Me.ComboBox1.Value <> CB_Val Then
       
    CB_Val = Me.ComboBox1.Value

    If ComboBox1.Value <> "" Then
        ComboBox1.ListFillRange = "UC_List"
        Me.ComboBox1.DropDown
    End If
    
    If ComboBox1.Value = "" Then
        ComboBox1.ListFillRange = "UC_List"
    End If

End If

Worksheets("PEC Calculator").Protect "eau"

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

Worksheets("PEC Calculator").Unprotect "eau"

Application.EnableEvents = False



If Worksheets("PEC Calculator").Range("D12").Value = "" Then
Worksheets("PEC Calculator").Range("D12").Value = 0
End If


If Worksheets("PEC Calculator").Range("D13").Value = "" Then
Worksheets("PEC Calculator").Range("D13").Value = 100
End If

'///////////////////////////////////////SHEET INITIALIZATION ///////////////////////////////////////////////////////////////////
    If Target.Address = "$D$3" Then
        If Range("D3").Value = "Public Domain" Or Range("D3").Value = "Agriculture" Or Range("D3").Value = "Chemical Industry (Basic)" Or _
        Range("D3").Value = "Electrical Industry" Or Range("D3").Value = "Paints, Lacquers, Varnishes" Or Range("D3").Value = "Mineral Oil Fuel Industry" Then
            Range("D7").Value = ""
            Range("D6").Value = ""
            Range("D4").Value = ""
            Range("D5").Value = ""
            ComboBox1.Value = ""
            ComboBox1.ListFillRange = "UC_List"
            Range("C23").Value = "Please Select"
            Range("C26").Value = "Please Select"
            Range("C29").Value = "Not Applicable"
            'Reset prod values too
            Range("D11").Value = ""
            Range("D12").Value = "0"
            Range("D13").Value = "100"
            Range("D14").Value = ""
            Range("D15").Value = ""
            Range("D17").Value = ""
            Range("D18").Value = ""
            Range("D19").Value = ""
        Else
            Range("D7").Value = ""
            Range("D6").Value = ""
            Range("D4").Value = ""
            Range("D5").Value = ""
            ComboBox1.Value = ""
            ComboBox1.ListFillRange = "UC_List"
            Range("C23").Value = "Please Select"
            Range("C26").Value = "Please Select"
            Range("C29").Value = "Please Select"
            'Reset prod values too
            Range("D11").Value = ""
            Range("D12").Value = "0"
            Range("D13").Value = "100"
            Range("D14").Value = ""
            Range("D15").Value = ""
            Range("D17").Value = ""
            Range("D18").Value = ""
            Range("D19").Value = ""
        End If
    End If
 
If Worksheets("PEC Calculator").Range("D11").Value = "1 facility" Then
Call TableSources

Call BTables_1F

Call ATables_1F
Call ATables_1FB


ElseIf Worksheets("PEC Calculator").Range("D11").Value = "Multiple facilities" Then
Call TableSources

Call BTables_Multiple

Call ATables_Multiple

End If


Application.EnableEvents = True

Worksheets("PEC Calculator").Protect "eau"

End Sub

Private Sub Worksheet_Calculate()
'https://stackoverflow.com/questions/40741967/run-macro-when-linked-cell-changes-value-excel-vba
Worksheets("PEC Calculator").Unprotect "eau"


Application.EnableEvents = False

    If Me.Range("D9").Value = "Cosmetics" Then
        
            If Range("D3").Value = "Public Domain" Then
            Range("C23").Value = "Not Applicable"
            Range("C26").Value = "Please Select"
            Range("C29").Value = "Not Applicable"
            End If

                    
        ElseIf Me.Range("D9").Value = "Cleaning/washing agents and additives" Then
        
            If Range("D3").Value = "Public Domain" Then
            Range("C23").Value = "Not Applicable"
            Range("C26").Value = "Not Applicable"
            Range("C29").Value = "Not Applicable"
            End If

                   
        ElseIf Me.Range("D9").Value = "Colouring agents" Then
        
            If Range("D3").Value = "Leather Processing" Then
            Range("C23").Value = "Please Select"
            Range("C26").Value = "Not Applicable"
            Range("C29").Value = "Not Applicable"
            End If
    
        
        ElseIf Me.Range("D9").Value <> "Cleaning/washing agents and additives" Or _
            Me.Range("D9").Value <> "Cosmetics" Or Me.Range("D9").Value <> "Colouring agents" Then
            
            If Range("D3").Value = "Public Domain" Then
            Range("C23").Value = "Please Select"
            Range("C26").Value = "Please Select"
            Range("C29").Value = "Not Applicable"
            End If
            
        End If
    
If Me.Range("D9").Value <> CB_Val Then
        'Reset prod values too
            Range("D11").Value = ""
            Range("D12").Value = 0
            Range("D13").Value = 100
            Range("D14").Value = ""
            Range("D15").Value = ""
            Range("D17").Value = ""
            Range("D18").Value = ""
            Range("D19").Value = ""
            
    If ComboBox1.Value <> "" Then
    ComboBox1.ListFillRange = "UC_List"
    Me.ComboBox1.DropDown
    End If
    
    If ComboBox1.Value = "" Then
    ComboBox1.ListFillRange = "UC_List"
    End If

    CB_Val = Me.Range("D9").Value
    
End If


Application.EnableEvents = True

Worksheets("PEC Calculator").Protect "eau"


End Sub


