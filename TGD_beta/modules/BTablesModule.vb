Sub BTables()

Dim tblA As ListObject
Dim nRows As Long
Dim nCols As Long
Dim i As Long
Dim VP As Long
Dim WS As Long
Dim Use As String
Dim Ind As String
Dim MC1 As String
Dim MC2 As String
Dim MC3 As String

VP = Worksheets("PEC Calculator").Range("D15").Value
WS = Worksheets("PEC Calculator").Range("D16").Value
Use = Worksheets("PEC Calculator").Range("D9").Value
Ind = Worksheets("PEC Calculator").Range("D3").Value

MC1 = Worksheets("PEC Calculator").Range("E22").Value
MC2 = Worksheets("PEC Calculator").Range("E25").Value
MC3 = Worksheets("PEC Calculator").Range("E28").Value

Dim tblB As ListObject
Dim nRowsB As Long
Dim nColsB As Long
Dim iB As Long


Set tblB = Worksheets("PEC Calculator").ListObjects("BTableINPUT")

'FMAIN SOURCE AND NO DAYS /////////////////////////////////////////////////////
' Production -------------------------------------------------------------------
Dim Tprod As Long
Tprod = Worksheets("PEC Calculator").Range("D11").Value

If tblB.Range(2, 2) = "TableB1.1" Then
    If Tprod < 1000 Then
    tblB.Range(3, 2) = 1
    tblB.Range(4, 2) = 0.1 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 1000 And Tprod < 2000 Then
    tblB.Range(3, 2) = 0.9
    tblB.Range(4, 2) = 0.1 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 2000 And Tprod < 4000 Then
    tblB.Range(3, 2) = 0.75
    tblB.Range(4, 2) = 0.1 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 4000 Then
    tblB.Range(3, 2) = 0.7
    tblB.Range(4, 2) = 300
    End If
ElseIf tblB.Range(2, 2) = "TableB1.2" Then
    If Tprod < 10 Then
    tblB.Range(3, 2) = 1
    tblB.Range(4, 2) = Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 10 And Tprod < 50 Then
    tblB.Range(3, 2) = 0.9
    tblB.Range(4, 2) = Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 50 And Tprod < 100 Then
    tblB.Range(3, 2) = 0.8
    tblB.Range(4, 2) = 0.6667 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 100 And Tprod < 1000 Then
    tblB.Range(3, 2) = 0.75
    tblB.Range(4, 2) = 0.4 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 1000 And Tprod < 2500 Then
    tblB.Range(3, 2) = 0.6
    tblB.Range(4, 2) = 0.2 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 2500 Then
    tblB.Range(3, 2) = 0.6
    tblB.Range(4, 2) = 300
    End If
ElseIf tblB.Range(2, 2) = "TableB1.3" Then
    If Tprod < 25000 Then
    tblB.Range(3, 2) = 1
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 25000 And Tprod < 100000 Then
    tblB.Range(3, 2) = 0.75
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 100000 Then
    tblB.Range(3, 2) = 0.6
    tblB.Range(4, 2) = 300
    End If
ElseIf tblB.Range(2, 2) = "TableB1.4" Then
    If Tprod < 5000 Then
    tblB.Range(3, 2) = 1
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 5000 And Tprod < 25000 Then
    tblB.Range(3, 2) = 0.8
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 25000 And Tprod < 100000 Then
    tblB.Range(3, 2) = 0.6
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 100000 Then
    tblB.Range(3, 2) = 0.4
    tblB.Range(4, 2) = 300
    End If
ElseIf tblB.Range(2, 2) = "TableB1.5" Then
    If Tprod < 25000 Then
    tblB.Range(3, 2) = 1
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 25000 And Tprod < 100000 Then
    tblB.Range(3, 2) = 0.75
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 100000 And Tprod < 500000 Then
    tblB.Range(3, 2) = 0.6
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 500000 Then
    tblB.Range(3, 2) = 0.5
    tblB.Range(4, 2) = 300
    End If
ElseIf tblB.Range(2, 2) = "TableB1.6" Then
    If Tprod < 10000 Then
    tblB.Range(3, 2) = 1
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 10000 And Tprod < 50000 Then
    tblB.Range(3, 2) = 0.75
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 50000 And Tprod < 250000 Then
    tblB.Range(3, 2) = 0.6
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 250000 Then
    tblB.Range(3, 2) = 0.5
    tblB.Range(4, 2) = 300
    End If
ElseIf tblB.Range(2, 2) = "TableB1.7" Then
    If Tprod < 100 Then
    tblB.Range(3, 2) = 1
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 100 And Tprod < 1000 Then
    tblB.Range(3, 2) = 0.9
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 1000 And Tprod < 2500 Then
    tblB.Range(3, 2) = 0.8
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 2500 Then
    tblB.Range(3, 2) = 0.8
    tblB.Range(4, 2) = 300
    End If
ElseIf tblB.Range(2, 2) = "TableB1.8" Then
    If Tprod < 1000 Then
    tblB.Range(3, 2) = 1
    tblB.Range(4, 2) = 0.1 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 1000 And Tprod < 4000 Then
    tblB.Range(3, 2) = 0.9
    tblB.Range(4, 2) = 0.1 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 4000 Then
    tblB.Range(3, 2) = 0.75
    tblB.Range(4, 2) = 300
    End If
ElseIf tblB.Range(2, 2) = "TableB1.9" Then
    If Tprod < 10 Then
    tblB.Range(3, 2) = 1
    tblB.Range(4, 2) = Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 10 And Tprod < 50 Then
    tblB.Range(3, 2) = 0.9
    tblB.Range(4, 2) = Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 50 And Tprod < 500 Then
    tblB.Range(3, 2) = 0.5
    tblB.Range(4, 2) = Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 500 And Tprod < 1500 Then
    tblB.Range(3, 2) = 0.2
    tblB.Range(4, 2) = Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 1500 Then
    tblB.Range(3, 2) = 0.2
    tblB.Range(4, 2) = 300
    End If
ElseIf tblB.Range(2, 2) = "TableB1.10" Then
    If Tprod < 10 Then
    tblB.Range(3, 2) = 1
    tblB.Range(4, 2) = Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 10 And Tprod < 50 Then
    tblB.Range(3, 2) = 0.9
    tblB.Range(4, 2) = Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 50 And Tprod < 500 Then
    tblB.Range(3, 2) = 0.8
    tblB.Range(4, 2) = 0.6667 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 500 And Tprod < 1500 Then
    tblB.Range(3, 2) = 0.5
    tblB.Range(4, 2) = 0.4 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 1500 Then
    tblB.Range(3, 2) = 0.5
    tblB.Range(4, 2) = 300
    End If
ElseIf tblB.Range(2, 2) = "TableB1.11" Then
    If Tprod < 100000 Then
    tblB.Range(3, 2) = 1
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 100000 And Tprod < 500000 Then
    tblB.Range(3, 2) = 0.75
    tblB.Range(4, 2) = 300
    ElseIf Tprod >= 500000 Then
    tblB.Range(3, 2) = 0.5
    tblB.Range(4, 2) = 300
    End If
ElseIf tblB.Range(2, 2) = "TableB1.12" Then
    If Tprod < 5 Then
    tblB.Range(3, 2) = 1
    tblB.Range(4, 2) = Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 5 And Tprod < 50 Then
    tblB.Range(3, 2) = 1
    tblB.Range(4, 2) = 0.5 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 50 And Tprod < 250 Then
    tblB.Range(3, 2) = 0.75
    tblB.Range(4, 2) = 0.4 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 250 And Tprod < 3000 Then
    tblB.Range(3, 2) = 0.5
    tblB.Range(4, 2) = 0.2 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 3000 Then
    tblB.Range(3, 2) = 0.5
    tblB.Range(4, 2) = 300
    End If
ElseIf tblB.Range(2, 2) = "TableB1.13" Then
    If Tprod < 50 Then
    tblB.Range(3, 2) = 0.9
    tblB.Range(4, 2) = 0.4 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 50 And Tprod < 500 Then
    tblB.Range(3, 2) = 0.75
    tblB.Range(4, 2) = 0.2 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 500 And Tprod < 5000 Then
    tblB.Range(3, 2) = 0.6
    tblB.Range(4, 2) = 0.1 * Val(tblB.Range(3, 2)) * Val(Tprod)
    ElseIf Tprod >= 5000 And Tprod < 25000 Then
    tblB.Range(3, 2) = 0.75
    tblB.Range(4, 2) = 200
    ElseIf Tprod >= 25000 Then
    tblB.Range(3, 2) = 0.5
    tblB.Range(4, 2) = 300
    End If
End If

'Formulation -------------------------------------------------------------
Dim T As Long
T = Worksheets("PEC Calculator").Range("D14").Value

If tblB.Range(2, 3) = "TableB2.1" Then
    If T < 100 Then
    tblB.Range(3, 3) = 1
    tblB.Range(4, 3) = 2 * Val(tblB.Range(3, 3)) * Val(T)
    ElseIf T >= 100 And T < 500 Then
    tblB.Range(3, 3) = 0.6
    tblB.Range(4, 3) = Val(tblB.Range(3, 3)) * Val(T)
    ElseIf T >= 500 And T < 1000 Then
    tblB.Range(3, 3) = 0.6
    tblB.Range(4, 3) = 0.5 * Val(tblB.Range(3, 3)) * Val(T)
    ElseIf T >= 1000 Then
    tblB.Range(3, 3) = 0.4
    tblB.Range(4, 3) = 300
    End If
ElseIf tblB.Range(2, 3) = "TableB2.2" Then
    If T < 15000 Then
    tblB.Range(3, 3) = 1
    tblB.Range(4, 3) = 300
    ElseIf T >= 15000 And T < 50000 Then
    tblB.Range(3, 3) = 0.75
    tblB.Range(4, 3) = 300
    ElseIf T >= 50000 Then
    tblB.Range(3, 3) = 0.6
    tblB.Range(4, 3) = 300
    End If
ElseIf tblB.Range(2, 3) = "TableB2.3" Then
    If T < 3500 Then
    tblB.Range(3, 3) = 1
    tblB.Range(4, 3) = 300
    ElseIf T >= 3500 And T < 10000 Then
    tblB.Range(3, 3) = 0.8
    tblB.Range(4, 3) = 300
    ElseIf T >= 10000 And T < 25000 Then
    tblB.Range(3, 3) = 0.7
    tblB.Range(4, 3) = 300
    ElseIf T >= 25000 And T < 50000 Then
    tblB.Range(3, 3) = 0.6
    tblB.Range(4, 3) = 300
    ElseIf T >= 50000 Then
    tblB.Range(3, 3) = 0.4
    tblB.Range(4, 3) = 300
    End If
ElseIf tblB.Range(2, 3) = "TableB2.4" Then
    If T < 10 Then
    tblB.Range(3, 3) = 1
    tblB.Range(4, 3) = 2 * Val(tblB.Range(3, 3)) * Val(T)
    ElseIf T >= 10 And T < 50 Then
    tblB.Range(3, 3) = 0.9
    tblB.Range(4, 3) = Val(tblB.Range(3, 3)) * Val(T)
    ElseIf T >= 50 And T < 500 Then
    tblB.Range(3, 3) = 0.8
    tblB.Range(4, 3) = 0.4 * Val(tblB.Range(3, 3)) * Val(T)
    ElseIf T >= 500 And T < 2000 Then
    tblB.Range(3, 3) = 0.75
    tblB.Range(4, 3) = 0.2 * Val(tblB.Range(3, 3)) * Val(T)
    ElseIf T >= 2000 Then
    tblB.Range(3, 3) = 0.65
    tblB.Range(4, 3) = 300
    End If
ElseIf tblB.Range(2, 3) = "TableB2.5" Then
    If T < 25000 Then
    tblB.Range(3, 3) = 1
    tblB.Range(4, 3) = 300
    ElseIf T >= 25000 And T < 50000 Then
    tblB.Range(3, 3) = 0.75
    tblB.Range(4, 3) = 300
    ElseIf T >= 50000 Then
    tblB.Range(3, 3) = 0.4
    tblB.Range(4, 3) = 300
    End If
ElseIf tblB.Range(2, 3) = "TableB2.6" Then
    If T < 100000 Then
    tblB.Range(3, 3) = 1
    tblB.Range(4, 3) = 300
    ElseIf T >= 100000 And T < 250000 Then
    tblB.Range(3, 3) = 0.7
    tblB.Range(4, 3) = 300
    ElseIf T >= 250000 Then
    tblB.Range(3, 3) = 0.4
    tblB.Range(4, 3) = 300
    End If
ElseIf tblB.Range(2, 3) = "TableB2.7" Then
    If T < 1000 Then
    tblB.Range(3, 3) = 1
    tblB.Range(4, 3) = 100
    ElseIf T >= 1000 And T < 2000 Then
    tblB.Range(3, 3) = 0.8
    tblB.Range(4, 3) = 200
    ElseIf T >= 2000 Then
    tblB.Range(3, 3) = 0.6
    tblB.Range(4, 3) = 300
    End If
ElseIf tblB.Range(2, 3) = "TableB2.8" Then
    If T < 5 Then
    tblB.Range(3, 3) = 1
    tblB.Range(4, 3) = 20
    ElseIf T >= 5 And T < 50 Then
    tblB.Range(3, 3) = 1
    tblB.Range(4, 3) = 60
    ElseIf T >= 50 And T < 100 Then
    tblB.Range(3, 3) = 1
    tblB.Range(4, 3) = 2 * Val(tblB.Range(3, 3)) * Val(T)
    ElseIf T >= 100 And T < 500 Then
    tblB.Range(3, 3) = 0.8
    tblB.Range(4, 3) = Val(tblB.Range(3, 3)) * Val(T)
    ElseIf T >= 500 And T < 1000 Then
    tblB.Range(3, 3) = 0.6
    tblB.Range(4, 3) = 0.5 * Val(tblB.Range(3, 3)) * Val(T)
    ElseIf T >= 1000 Then
    tblB.Range(3, 3) = 0.4
    tblB.Range(4, 3) = 300
    End If
ElseIf tblB.Range(2, 3) = "TableB2.9" Then
    If T < 25000 Then
    tblB.Range(3, 3) = 1
    tblB.Range(4, 3) = 300
    ElseIf T >= 25000 And T < 50000 Then
    tblB.Range(3, 3) = 0.75
    tblB.Range(4, 3) = 300
    ElseIf T >= 50000 Then
    tblB.Range(3, 3) = 0.4
    tblB.Range(4, 3) = 300
    End If
ElseIf tblB.Range(2, 3) = "TableB2.10" Then
    If T < 3500 Then
    tblB.Range(3, 3) = 1
    tblB.Range(4, 3) = 300
    ElseIf T >= 3500 And T < 10000 Then
    tblB.Range(3, 3) = 0.8
    tblB.Range(4, 3) = 300
    ElseIf T >= 10000 And T < 25000 Then
    tblB.Range(3, 3) = 0.7
    tblB.Range(4, 3) = 300
    ElseIf T >= 25000 And T < 50000 Then
    tblB.Range(3, 3) = 0.6
    tblB.Range(4, 3) = 300
    ElseIf T >= 50000 Then
    tblB.Range(3, 3) = 0.4
    tblB.Range(4, 3) = 300
    End If
End If
    
'Industrial -----------------------------------------------------------------------
If tblB.Range(2, 4) = "TableB3.1" Then
    If T < 10 Then
    tblB.Range(3, 4) = 0.05
    ElseIf T >= 10 And T < 100 Then
    tblB.Range(3, 4) = 0.01
    ElseIf T >= 100 And T < 1000 Then
    tblB.Range(3, 4) = 0.005
    ElseIf T >= 1000 And T < 10000 Then
    tblB.Range(3, 4) = 0.001
    ElseIf T >= 10000 And T < 50000 Then
    tblB.Range(3, 4) = 0.0005
    ElseIf T >= 50000 Then
    tblB.Range(3, 4) = 0.00001
    End If
If Use = "Pharmaceuticals" Then
    tblB.Range(4, 4) = 10
    
    If T < 10 Then
    tblB.Range(3, 4) = 0.05
    ElseIf T >= 10 And T < 100 Then
    tblB.Range(3, 4) = 0.01
    ElseIf T >= 100 And T < 1000 Then
    tblB.Range(3, 4) = 0.005
    ElseIf T >= 1000 And T < 10000 Then
    tblB.Range(3, 4) = 0.001
    ElseIf T >= 10000 And T < 50000 Then
    tblB.Range(3, 4) = 0.0005
    ElseIf T >= 50000 Then
    tblB.Range(3, 4) = 0.00001
    End If
    
    ElseIf Use = "Odour agents" Or Use = "Cleaning/washing agents and additives" Or Use = "Colouring agents" Then
    tblB.Range(4, 4) = 50
    
    If T < 10 Then
    tblB.Range(3, 4) = 0.05
    ElseIf T >= 10 And T < 100 Then
    tblB.Range(3, 4) = 0.01
    ElseIf T >= 100 And T < 1000 Then
    tblB.Range(3, 4) = 0.005
    ElseIf T >= 1000 And T < 10000 Then
    tblB.Range(3, 4) = 0.001
    ElseIf T >= 10000 And T < 50000 Then
    tblB.Range(3, 4) = 0.0005
    ElseIf T >= 50000 Then
    tblB.Range(3, 4) = 0.00001
    End If
    
    ElseIf Use = "Food/feedstuff additives" Then
    tblB.Range(4, 4) = 300
    
    If T < 10 Then
    tblB.Range(3, 4) = 0.05
    ElseIf T >= 10 And T < 100 Then
    tblB.Range(3, 4) = 0.01
    ElseIf T >= 100 And T < 1000 Then
    tblB.Range(3, 4) = 0.005
    ElseIf T >= 1000 And T < 10000 Then
    tblB.Range(3, 4) = 0.001
    ElseIf T >= 10000 And T < 50000 Then
    tblB.Range(3, 4) = 0.0005
    ElseIf T >= 50000 Then
    tblB.Range(3, 4) = 0.00001
    End If
    
    ElseIf Use = "Aerosol Propellants" Or Use = "Fertilisers" Or Use = "Biocides, non-agricultural" Or Use = "Solvents" Or Use = "Surface-actives agents" Then tblB.Range(4, 4) = 2
    tblB.Range(4, 4) = 2
    
    If T < 10 Then
    tblB.Range(3, 4) = 0.05
    ElseIf T >= 10 And T < 100 Then
    tblB.Range(3, 4) = 0.01
    ElseIf T >= 100 And T < 1000 Then
    tblB.Range(3, 4) = 0.005
    ElseIf T >= 1000 And T < 10000 Then
    tblB.Range(3, 4) = 0.001
    ElseIf T >= 10000 And T < 50000 Then
    tblB.Range(3, 4) = 0.0005
    ElseIf T >= 50000 Then
    tblB.Range(3, 4) = 0.00001
    End If
    
    Else
    tblB.Range(3, 4) = "Choose appropriate use"
    tblB.Range(4, 4) = "Choose appropriate use"
    
    End If
ElseIf tblB.Range(2, 4) = "TableB3.2" Then
    If T < 10 Then
    tblB.Range(3, 4) = 0.8
    tblB.Range(4, 4) = 2 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 10 And T < 50 Then
    tblB.Range(3, 4) = 0.65
    tblB.Range(4, 4) = Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 50 And T < 500 Then
    tblB.Range(3, 4) = 0.5
    tblB.Range(4, 4) = 0.4 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 500 And T < 2000 Then
    tblB.Range(3, 4) = 0.4
    tblB.Range(4, 4) = 0.25 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 2000 And T < 5000 Then
    tblB.Range(3, 4) = 0.3
    tblB.Range(4, 4) = 0.2 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 5000 And T < 25000 Then
    tblB.Range(3, 4) = 0.25
    tblB.Range(4, 4) = 300
    ElseIf T >= 25000 And T < 75000 Then
    tblB.Range(3, 4) = 0.2
    tblB.Range(4, 4) = 300
    ElseIf T >= 75000 Then
    tblB.Range(3, 4) = 0.15
    tblB.Range(4, 4) = 300
    End If
ElseIf tblB.Range(2, 4) = "TableB3.3" Then
    tblB.Range(3, 4) = 0.002
    
    If Use = "Biocides, non-agricultural" Then
    tblB.Range(4, 4) = 15
    ElseIf Use = "Cleaning/washing agents and additives" Then
    tblB.Range(4, 4) = 200
    Else
    tblB.Range(4, 4) = 50
    End If
ElseIf tblB.Range(2, 4) = "TableB3.4" Then
    If T < 10 Then
    tblB.Range(3, 4) = 0.8
    tblB.Range(4, 4) = 2 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 10 And T < 50 Then
    tblB.Range(3, 4) = 0.75
    tblB.Range(4, 4) = 2 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 50 And T < 500 Then
    tblB.Range(3, 4) = 0.6
    tblB.Range(4, 4) = Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 500 And T < 1500 Then
    tblB.Range(3, 4) = 0.5
    tblB.Range(4, 4) = 0.4 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 1500 And T < 5000 Then
    tblB.Range(3, 4) = 0.35
    tblB.Range(4, 4) = 300
    ElseIf T >= 5000 And T < 25000 Then
    tblB.Range(3, 4) = 0.2
    tblB.Range(4, 4) = 300
    ElseIf T >= 25000 Then
    tblB.Range(3, 4) = 0.1
    tblB.Range(4, 4) = 300
    End If
ElseIf tblB.Range(2, 4) = "TableB3.6" Then
    If T < 10 Then
    tblB.Range(3, 4) = 1
    tblB.Range(4, 4) = 2 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 10 And T < 50 Then
    tblB.Range(3, 4) = 1
    tblB.Range(4, 4) = 0.5 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 50 And T < 500 Then
    tblB.Range(3, 4) = 0.9
    tblB.Range(4, 4) = 0.4 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 500 And T < 2000 Then
    tblB.Range(3, 4) = 0.8
    tblB.Range(4, 4) = 0.1875 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 2000 And T < 10000 Then
    tblB.Range(3, 4) = 0.7
    tblB.Range(4, 4) = 300
    ElseIf T >= 10000 And T < 50000 Then
    tblB.Range(3, 4) = 0.6
    tblB.Range(4, 4) = 300
    ElseIf T >= 50000 Then
    tblB.Range(3, 4) = 0.5
    tblB.Range(4, 4) = 300
    End If
ElseIf tblB.Range(2, 4) = "TableB3.7" Then
    If T < 50 Then
    tblB.Range(3, 4) = 0.5
    tblB.Range(4, 4) = 350
    ElseIf T >= 50 And T < 500 Then
    tblB.Range(3, 4) = 0.4
    tblB.Range(4, 4) = 350
    ElseIf T >= 500 And T < 5000 Then
    tblB.Range(3, 4) = 0.3
    tblB.Range(4, 4) = 350
    ElseIf T >= 5000 And T < 25000 Then
    tblB.Range(3, 4) = 0.2
    tblB.Range(4, 4) = 350
    ElseIf T >= 25000 And T < 100000 Then
    tblB.Range(3, 4) = 0.05
    tblB.Range(4, 4) = 350
    ElseIf T >= 100000 Then
    tblB.Range(3, 4) = 0.02
    tblB.Range(4, 4) = 350
    End If
ElseIf tblB.Range(2, 4) = "TableB3.8" Then
    If Worksheets("PEC Calculator").Range("B5").Value = "1 company" Then
    tblB.Range(3, 4) = 1
    tblB.Range(4, 4) = 300
    ElseIf Worksheets("PEC Calculator").Range("B5").Value = "Small companies" Then
    tblB.Range(3, 4) = 0.333
    tblB.Range(4, 4) = 300
    ElseIf Worksheets("PEC Calculator").Range("B5").Value = "Large companies" Then
    tblB.Range(3, 4) = 0.05
    tblB.Range(4, 4) = 300
    End If
ElseIf tblB.Range(2, 4) = "TableB3.13" Then
    If T < 10 Then
    tblB.Range(3, 4) = 0.9
    tblB.Range(4, 4) = 20 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 10 And T < 50 Then
    tblB.Range(3, 4) = 0.6
    tblB.Range(4, 4) = 6.667 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 50 And T < 300 Then
    tblB.Range(3, 4) = 0.3
    tblB.Range(4, 4) = 3.333 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 300 And T < 5000 Then
    tblB.Range(3, 4) = 0.15
    tblB.Range(4, 4) = 300
    ElseIf T >= 5000 And T < 25000 Then
    tblB.Range(3, 4) = 0.1
    tblB.Range(4, 4) = 300
    ElseIf T >= 25000 Then
    tblB.Range(3, 4) = 0.05
    tblB.Range(4, 4) = 300
    End If
ElseIf tblB.Range(2, 4) = "TableB3.14" Then
    If T < 10 Then
    tblB.Range(3, 4) = 1
    tblB.Range(4, 4) = 2 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 10 And T < 50 Then
    tblB.Range(3, 4) = 0.9
    tblB.Range(4, 4) = Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 50 And T < 500 Then
    tblB.Range(3, 4) = 0.8
    tblB.Range(4, 4) = 0.4 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 500 And T < 2000 Then
    tblB.Range(3, 4) = 0.75
    tblB.Range(4, 4) = 0.2 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 2000 And T < 5000 Then
    tblB.Range(3, 4) = 0.6
    tblB.Range(4, 4) = 0.1 * Val(tblB.Range(3, 4)) * Val(T)
    ElseIf T >= 5000 And T < 25000 Then
    tblB.Range(3, 4) = 0.5
    tblB.Range(4, 4) = 300
    ElseIf T >= 25000 Then
    tblB.Range(3, 4) = 0.3
    tblB.Range(4, 4) = 300
    End If
End If

'Private Use -----------------------------------------------------------------------
If tblB.Range(2, 5) = "TableB4.1" Then
    tblB.Range(3, 5) = 0.002
    tblB.Range(4, 5) = 365
ElseIf tblB.Range(2, 5) = "TableB4.2" Then
    If Worksheets("PEC Calculator").Range("D5").Value = "Small companies" Then
        If T < 10 Then
        tblB.Range(3, 5) = 0
        tblB.Range(4, 5) = 200
        ElseIf T >= 10 And T < 50 Then
        tblB.Range(3, 5) = 0.00000004
        tblB.Range(4, 5) = 200
        ElseIf T >= 50 And T < 500 Then
        tblB.Range(3, 5) = 0.0000002
        tblB.Range(4, 5) = 200
        ElseIf T >= 500 And T < 5000 Then
        tblB.Range(3, 5) = 0.000001
        tblB.Range(4, 5) = 200
        ElseIf T >= 5000 Then
        tblB.Range(3, 5) = 0.00001
        tblB.Range(4, 5) = 200
        End If
    Else
    tblB.Range(3, 5) = 0
    tblB.Range(4, 5) = ""
    End If
ElseIf tblB.Range(2, 5) = "TableB4.3" Then
    If T < 50 Then
    tblB.Range(3, 5) = 0
    tblB.Range(4, 5) = ""
    ElseIf T >= 50 And T < 500 Then
    tblB.Range(3, 5) = 0.000004
    tblB.Range(4, 5) = 300
    ElseIf T >= 500 Then
    tblB.Range(3, 5) = 0.00002
    tblB.Range(4, 5) = 300
    End If
ElseIf tblB.Range(2, 5) = "TableB4.4" Then
    If T < 500 Then
    tblB.Range(3, 5) = 0.002
    tblB.Range(4, 5) = 150
    ElseIf T >= 500 Then
    tblB.Range(3, 5) = 0.002
    tblB.Range(4, 5) = 300
    End If
ElseIf tblB.Range(2, 5) = "TableB4.5" Then
    If T < 50 Then
    tblB.Range(3, 5) = 0
    tblB.Range(4, 5) = ""
    ElseIf T >= 50 And T < 500 Then
    tblB.Range(3, 5) = 0.00000004
    tblB.Range(4, 5) = 200
    ElseIf T >= 500 And T < 2500 Then
    tblB.Range(3, 5) = 0.0000008
    tblB.Range(4, 5) = 300
    ElseIf T >= 2500 And T < 10000 Then
    tblB.Range(3, 5) = 0.000004
    tblB.Range(4, 5) = 300
    ElseIf T >= 10000 And T < 50000 Then
    tblB.Range(3, 5) = 0.00002
    tblB.Range(4, 5) = 300
    ElseIf T >= 50000 Then
    tblB.Range(3, 5) = 0.0001
    tblB.Range(4, 5) = 300
    End If
Else
tblB.Range(3, 5) = ""
tblB.Range(4, 5) = ""
End If

'Waste Treatment ---------------------------------------------------------------------
If tblB.Range(2, 6) = "TableB5.1" Then
    If Worksheets("PEC Calculator").Range("D5").Value = "1 company" Then
        If T < 10 Then
        tblB.Range(3, 6) = 1
        tblB.Range(4, 6) = 150
        ElseIf T >= 10 Then
        tblB.Range(3, 6) = 1
        tblB.Range(4, 6) = 300
        End If
    ElseIf Worksheets("PEC Calculator").Range("D5").Value = "Large companies" Then
        If T < 10 Then
        tblB.Range(3, 6) = 0.333
        tblB.Range(4, 6) = 150
        ElseIf T >= 10 Then
        tblB.Range(3, 6) = 0.333
        tblB.Range(4, 6) = 300
        End If
    ElseIf Worksheets("PEC Calculator").Range("D5").Value = "Small companies" Then
        If T < 10 Then
        tblB.Range(3, 6) = 0.2
        tblB.Range(4, 6) = 150
        ElseIf T >= 10 Then
        tblB.Range(3, 6) = 0.2
        tblB.Range(4, 6) = 300
        End If
    End If
ElseIf tblB.Range(2, 6) = "TableB5.2" Then
If T < 100 Then
    tblB.Range(3, 6) = 0.5
    tblB.Range(4, 6) = 150
    ElseIf T >= 100 And T < 1000 Then
    tblB.Range(3, 6) = 0.4
    tblB.Range(4, 6) = 200
    ElseIf T >= 1000 And T < 10000 Then
    tblB.Range(3, 6) = 0.3
    tblB.Range(4, 6) = 250
    ElseIf T >= 10000 And T < 100000 Then
    tblB.Range(3, 6) = 0.2
    tblB.Range(4, 6) = 300
    ElseIf T >= 100000 Then
    tblB.Range(3, 6) = 0.1
    tblB.Range(4, 6) = 300
    End If
Else
tblB.Range(3, 6) = ""
tblB.Range(4, 6) = ""
End If

End Sub
