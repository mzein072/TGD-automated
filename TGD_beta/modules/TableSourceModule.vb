Sub TableSources()

Dim tblA As ListObject
Dim nRows As Long
Dim nCols As Long
Dim i As Long
Dim VP As Long
Dim WS As Long
Dim Tprod As Long
Dim T As Long
Dim Use As String
Dim Ind As String
Dim MC1 As String
Dim MC2 As String
Dim MC3 As String

VP = Worksheets("PEC Calculator").Range("D15").Value
WS = Worksheets("PEC Calculator").Range("D16").Value
Use = Worksheets("PEC Calculator").Range("D9").Value
Ind = Worksheets("PEC Calculator").Range("D3").Value
Tprod = Worksheets("PEC Calculator").Range("D11").Value
T = Worksheets("PEC Calculator").Range("D14").Value
MC1 = Worksheets("PEC Calculator").Range("E22").Value
MC2 = Worksheets("PEC Calculator").Range("E25").Value
MC3 = Worksheets("PEC Calculator").Range("E28").Value

Set tblA = Worksheets("PEC Calculator").ListObjects("ATableINPUT")

'A Table Sources ------------------------------------------------------------
'Production
If Ind = "Public Domain" Then
    If (Use = "Cleaning/washing agents and additives" Or Use = "Cosmetics") And Range("D11").Value < 1000 Then
    tblA.Range(2, 2) = "TableA1"
    Else
    tblA.Range(2, 2) = "TableA1.1"
    End If
ElseIf Ind = "Chemical Industry (Synthesis)" And Use = "Intermediates" Then
tblA.Range(2, 2) = "TableA1.2"
ElseIf (Ind = "Leather Processing" And Use = "Colouring agents") Then
tblA.Range(2, 2) = "TableA1.3"
Else
tblA.Range(2, 2) = "TableA1.1"
End If

'Formulation
If Ind = "Public Domain" Then
    If Use = "Cleaning/washing agents and additives" Then
    tblA.Range(2, 3) = "TableA2"
    Else
    tblA.Range(2, 3) = "TableA2.1"
    End If
ElseIf Ind = "Metal Industry" And (Use = "Heat transferring agents" Or Use = "Lubricants and additives") Then
tblA.Range(2, 3) = "TableA2.2"
ElseIf Ind = "Photographic Industry" And Use = "Photochemicals" Then
tblA.Range(2, 3) = "TableA2.3"
Else
tblA.Range(2, 3) = "TableA2.1"
End If

'Industrial
If Ind = "Agriculture" Then
tblA.Range(2, 4) = "TableA3.1"
ElseIf Ind = "Paints, Lacquers, Varnishes" Then
tblA.Range(2, 4) = "TableA3.15"
ElseIf Ind = "Engineering Industry" Or Ind = "Other" Then
tblA.Range(2, 4) = "TableA3.16"
ElseIf Ind = "Chemical Industry (Basic)" Then
tblA.Range(2, 4) = "TableA3.2"
ElseIf Ind = "Chemical Industry (Synthesis)" Then
tblA.Range(2, 4) = "TableA3.3"
ElseIf Ind = "Electrical Industry" Then
tblA.Range(2, 4) = "TableA3.4"
ElseIf Ind = "Public Domain" Then
tblA.Range(2, 4) = "TableA3.5"
ElseIf Ind = "Leather Processing" Then
tblA.Range(2, 4) = "TableA3.6"
ElseIf Ind = "Metal Industry" Then
tblA.Range(2, 4) = "TableA3.7"
ElseIf Ind = "Mineral Oil Fuel Industry" Then
tblA.Range(2, 4) = "TableA3.8"
ElseIf Ind = "Photographic Industry" Then
tblA.Range(2, 4) = "TableA3.9"
End If


'Private Use
If Ind = "Mineral Oil Fuel Industry" Then
tblA.Range(2, 5) = "TableA4.2"
ElseIf Ind = "Photographic Industry" And Use = "Photochemicals" And Worksheets("PEC Calculator").Range("D6").Value = "No" Then
tblA.Range(2, 5) = "TableA4.3"
ElseIf Ind = "Paints, Lacquers, Varnishes" Then
tblA.Range(2, 5) = "TableA4.5"
ElseIf Ind = "Engineering Industry" Then
tblA.Range(2, 5) = "TableA3.16"
Else
tblA.Range(2, 5) = "Not Applicable"
End If

'Waste Treatment
If Ind = "Photographic Industry" And Use = "Photochemicals" And Worksheets("PEC Calculator").Range("D6").Value = "No" Then
tblA.Range(2, 6) = "TableA5.1"
Else
tblA.Range(2, 6) = "Not Applicable"
End If



'////////// B TABLES //////////////
Dim tblB As ListObject
Dim nRowsB As Long
Dim nColsB As Long
Dim iB As Long

Set tblB = Worksheets("PEC Calculator").ListObjects("BTableINPUT")

'B Table Sources ------------------------------------------------------------
'Production
If Ind = "Leather Processing" Then
    If T > 5000 Then
        If Use <> "Colouring agents" And Use <> "Bleaching agents" And Use <> "Cleaning/washing agents and additivies" And _
        Use <> "Impregnation agents" Then
        tblB.Range(2, 2) = "TableB1.4"
        End If
    ElseIf T > 2500 Then
        If Use = "Colouring agents" Or Use = "Bleaching agents" Or Use = "Cleaning/washing agents and additivies" Or _
        Use = "Impregnation agents" Then
        tblB.Range(2, 2) = "TableB1.4"
        End If
    ElseIf T < 5000 Then
        If Use <> "Colouring agents" And Use <> "Bleaching agents" And Use <> "Cleaning/washing agents and additivies" And _
        Use <> "Impregnation agents" Then
        tblB.Range(2, 2) = "TableB1.8"
        End If
    ElseIf T < 2500 Then
        If Use = "Colouring agents" Or Use = "Bleaching agents" Or Use = "Cleaning/washing agents and additivies" Or _
        Use = "Impregnation agents" Then
        tblB.Range(2, 2) = "TableB1.9"
        End If
    End If
ElseIf Ind = "Public Domain" Or Ind = "Electrical Industry" Then
    If T < 7000 Then
    tblB.Range(2, 2) = "TableB1.7"
    Else
    tblB.Range(2, 2) = "TableB1.6"
    End If
ElseIf Ind = "Chemical Industry (Synthesis)" Then
    If T < 7000 Then
    tblB.Range(2, 2) = "TableB1.2"
    ElseIf T > 7000 Then
    tblB.Range(2, 2) = "TableB1.6"
    End If
ElseIf Ind = "Agriculture" Then
    If (Range("D5").Value = "Yes" Or Range("B6").Value = "Yes") Then
        If Range("D4").Value = "Yes" Then
        tblB.Range(2, 2) = "TableB1.4"
        Else
            If T > 10000 Then
            tblB.Range(2, 2) = "TableB1.3"
            Else
            tblB.Range(2, 2) = "TableB1.2"
            End If
        End If
    Else
    tblB.Range(2, 2) = "TableB1.1"
    End If
ElseIf (Ind = "Paints, Lacquers, Varnishes" Or Ind = "Engineering Industry" Or Ind = "Others") Then
    If T < 7000 Then
    tblB.Range(2, 2) = "TableB1.2"
    ElseIf T >= 7000 Then
    tblB.Range(2, 2) = "TableB1.6"
    End If
ElseIf Ind = "Photographic Industry" Then
    If Tprod < 4000 Then
    tblB.Range(2, 2) = "TableB1.12"
    ElseIf Tprod >= 4000 Then
    tblB.Range(2, 2) = "TableB1.4"
    End If
ElseIf Ind = "Mineral Oil Fuel Industry" Then
    If Use = "Fuels" Then
        If T < 25000 Then
        tblB.Range(2, 2) = "TableB1.1"
        ElseIf T >= 25000 Then
        tblB.Range(2, 2) = "TableB1.11"
        End If
    ElseIf Use <> "Fuels" Then
        If T < 3000 Then
        tblB.Range(2, 2) = "TableB1.2"
        ElseIf T >= 3000 Then
        tblB.Range(2, 2) = "TableB1.4"
        End If
    End If
ElseIf Ind = "Metal Industry" Then
    If Use = "Heat transferring agents" Or Use = "Lubricants and additives" Then
        If T < 2500 Then
        tblB.Range(2, 2) = "TableB1.10"
        ElseIf T >= 2500 Then
        tblB.Range(2, 2) = "TableB1.4"
        End If
    Else
        If T < 7000 Then
        tblB.Range(2, 2) = "TableB1.2"
        ElseIf T >= 7000 Then
        tblB.Range(2, 2) = "TableB1.6"
        End If
    End If
ElseIf T >= 10000 Then
tblB.Range(2, 2) = "TableB1.5"
Else
tblB.Range(2, 2) = "TableB1.1"
End If

'Formulation
If Ind = "Leather Processing" Then
    If T > 5000 Then
        If (Use = "Cleaning/washing agents and additives" Or _
        Use = "Colouring agents" Or _
        Use = "Impregnation agents") Then
        tblB.Range(2, 3) = "TableB2.6"
        Else
        tblB.Range(2, 3) = "TableB2.3"
        End If
    Else
        tblB.Range(2, 3) = "TableB2.4"
    End If
ElseIf Ind = "Public Domain" Then
    If T < 7000 Then
    tblB.Range(2, 3) = "TableB2.1"
    Else
    tblB.Range(2, 3) = "TableB2.3"
    End If
ElseIf Ind = "Chemical Industry (Synthesis)" Or Ind = "Electrical Industry" Then
    If T < 7000 Then
    tblB.Range(2, 3) = "TableB2.4"
    Else
    tblB.Range(2, 3) = "TableB2.3"
    End If
ElseIf Ind = "Agriculture" Then
    If (Range("D5").Value = "Yes" Or Range("D6").Value = "Yes") Then
    tblB.Range(2, 3) = "TableB2.3"
    ElseIf Range("D4").Value = "Yes" Then
    tblB.Range(2, 3) = "TableB2.3"
    Else
    tblB.Range(2, 3) = "TableB2.1"
    End If
ElseIf Ind = "Engineering Industry" Or Ind = "Other" Then
    If T < 7000 Then
    tblB.Range(2, 3) = "TableB2.8"
    ElseIf T >= 7000 Then
    tblB.Range(2, 3) = "TableB2.3"
    End If
ElseIf Ind = "Paints, Lacquers, Varnishes" Then
    If T < 7000 Then
    tblB.Range(2, 3) = "TableB2.10"
    ElseIf T >= 7000 Then
    tblB.Range(2, 3) = "TableB2.3"
    End If
ElseIf Ind = "Photographic Industry" Then
    If T < 4000 Then
    tblB.Range(2, 3) = "TableB2.8"
    ElseIf T >= 4000 Then
    tblB.Range(2, 3) = "TableB2.3"
    End If
ElseIf Ind = "Mineral Oil Fuel Industry" Then
    If Use = "Fuels" Then
        If T < 25000 Then
        tblB.Range(2, 3) = "TableB2.7"
        ElseIf T >= 25000 Then
        tblB.Range(2, 3) = "TableB2.6"
        End If
    Else
        If T < 3000 Then
        tblB.Range(2, 3) = "TableB2.8"
        ElseIf T >= 3000 Then
        tblB.Range(2, 3) = "TableB2.6"
        End If
    End If
ElseIf Ind = "Metal Industry" Then
    If T < 10000 Then
    tblB.Range(2, 3) = "TableB2.4"
    ElseIf T >= 10000 Then
    tblB.Range(2, 3) = "TableB2.3"
    End If
ElseIf T >= 10000 Then
tblB.Range(2, 3) = "TableB2.5"
Else
tblB.Range(2, 3) = "TableB2.4"
End If

'Industrial
If Ind = "Agriculture" Then
tblB.Range(2, 4) = "TableB3.1"
ElseIf Ind = "Public Domain" Then
tblB.Range(2, 4) = "TableB3.3"
ElseIf Ind = "Leather Processing" Then
tblB.Range(2, 4) = "TableB3.4"
ElseIf Ind = "Metal Industry" Then
tblB.Range(2, 4) = "TableB3.6"
ElseIf Ind = "Mineral Oil Fuel Industry" Then
tblB.Range(2, 4) = "TableB3.7"
ElseIf Ind = "Photographic Industry" Then
tblB.Range(2, 4) = "TableB3.8"
ElseIf Ind = "Paints, Lacquers, Varnishes" Then
tblB.Range(2, 4) = "TableB3.13"
ElseIf Ind = "Engineering Industry" Or Ind = "Other" Then
tblB.Range(2, 4) = "TableB3.14"
Else
tblB.Range(2, 4) = "TableB3.2"
End If

'Private Use
If Ind = "Engineering Industry" Then
    tblB.Range(2, 5) = "TableB4.5"
ElseIf Ind = "Paints, Lacquers, Varnishes" Then
    If Worksheets("PEC Calculator").Range("D5").Value = "Construction/maintenance" Then
    tblB.Range(2, 5) = "TableB4.5"
    Else
    tblB.Range(2, 5) = "TableB4.4"
    End If
ElseIf Ind = "Photographic Industry" Then
    tblB.Range(2, 5) = "TableB4.2"
ElseIf Ind = "Mineral Oil Fuel Industry" Then
    tblB.Range(2, 5) = "TableB4.1"
Else
    tblB.Range(2, 5) = "Not Applicable"
End If

'Wastewater
If Ind = "Photographic Industry" Then
tblB.Range(2, 6) = "TableB5.1"
Else
tblB.Range(2, 6) = "Not Applicable"
End If

End Sub
