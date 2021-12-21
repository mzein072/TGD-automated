Dim CB_Val As Variant
Private Sub ComboBox1_change()
'DROP-DOWN USE CATEGORY BOX ///////////////////////////////////////////////////////

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
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Application.EnableEvents = False

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
            Range("C22").Value = "Please Select"
            Range("C25").Value = "Please Select"
            Range("C28").Value = "Not Applicable"
            'Reset prod values too
            Range("D11").Value = ""
            Range("D12").Value = ""
            Range("D13").Value = ""
            Range("D15").Value = ""
            Range("D16").Value = ""
            Range("D17").Value = ""
            Range("D18").Value = ""
        Else
            Range("D7").Value = ""
            Range("D6").Value = ""
            Range("D4").Value = ""
            Range("D5").Value = ""
            ComboBox1.Value = ""
            ComboBox1.ListFillRange = "UC_List"
            Range("C22").Value = "Please Select"
            Range("C25").Value = "Please Select"
            Range("C28").Value = "Please Select"
            'Reset prod values too
            Range("D11").Value = ""
            Range("D12").Value = ""
            Range("D13").Value = ""
            Range("D15").Value = ""
            Range("D16").Value = ""
            Range("D17").Value = ""
            Range("D18").Value = ""
        End If
    End If
    
Call TableSources

Call BTables

Call ATables

Call ATables2




Application.EnableEvents = True
End Sub

Private Sub Worksheet_Calculate()
'https://stackoverflow.com/questions/40741967/run-macro-when-linked-cell-changes-value-excel-vba

Application.EnableEvents = False

    If Me.Range("D9").Value = "Cosmetics" Then
        
            If Range("D3").Value = "Public Domain" Then
            Range("C22").Value = "Not Applicable"
            Range("C25").Value = "Please Select"
            Range("C28").Value = "Not Applicable"
            End If

                    
        ElseIf Me.Range("D9").Value = "Cleaning/washing agents and additives" Then
        
            If Range("D3").Value = "Public Domain" Then
            Range("C22").Value = "Not Applicable"
            Range("C25").Value = "Not Applicable"
            Range("C28").Value = "Not Applicable"
            End If

                   
        ElseIf Me.Range("D9").Value = "Colouring agents" Then
        
            If Range("D3").Value = "Leather Processing" Then
            Range("C22").Value = "Please Select"
            Range("C25").Value = "Not Applicable"
            Range("C28").Value = "Not Applicable"
            End If
    
        
        ElseIf Me.Range("D9").Value <> "Cleaning/washing agents and additives" Or _
            Me.Range("D9").Value <> "Cosmetics" Or Me.Range("D9").Value <> "Colouring agents" Then
            
            If Range("D3").Value = "Public Domain" Then
            Range("C22").Value = "Please Select"
            Range("C25").Value = "Please Select"
            Range("C28").Value = "Not Applicable"
            End If
            
        End If
    
If Me.Range("D9").Value <> CB_Val Then
        'Reset prod values too
            Range("D11").Value = ""
            Range("D12").Value = ""
            Range("D13").Value = ""
            Range("D15").Value = ""
            Range("D16").Value = ""
            Range("D17").Value = ""
            Range("D18").Value = ""
            
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
End Sub


