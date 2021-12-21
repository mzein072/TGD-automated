Sub ATables_1F()

Worksheets("PEC Calculator").Unprotect "eau"


'Initiate all variables -----
Dim tblA As ListObject
Dim nRows As Long
Dim nCols As Long
Dim i As Long
Dim VP As Long
Dim WS As Long
Dim HLC As Long
Dim SQ As Long
Dim Import As Long
Dim Export As Long
Dim Tprod As Long
Dim Conc As Long
Dim T As Long
Dim Use As String
Dim Ind As String
Dim MC1 As String
Dim MC2 As String
Dim MC3 As String

'Assign values to all variables ------
Use = Worksheets("PEC Calculator").Range("D9").Value 'Use category
Ind = Worksheets("PEC Calculator").Range("D3").Value 'Main industry
VP = Worksheets("PEC Calculator").Range("D17").Value 'Vapour pressure
WS = Worksheets("PEC Calculator").Range("D18").Value 'Water solubility
HLC = Worksheets("PEC Calculator").Range("D19").Value 'Log Henry (UC 29, 35 only)
SQ = Worksheets("PEC Calculator").Range("D12").Value 'Substance quantity at 1 facility
Conc = Worksheets("PEC Calculator").Range("D13").Value 'Max concentration of substance in products
Import = Worksheets("PEC Calculator").Range("D14").Value 'Import volume
Export = Worksheets("PEC Calculator").Range("D15").Value 'Export volume

'For production, total volume at 1 facility simply equals product volume (no import/export)
Tprod = CLng(SQ) / (CLng(Conc) / 100)
Debug.Print Tprod
'Total volume at 1 facility = product volume + import volume - export volume
T = (CLng(SQ) / (CLng(Conc) / 100)) + Val(Import) - Val(Export)
Debug.Print T

MC1 = Worksheets("PEC Calculator").Range("E23").Value
MC2 = Worksheets("PEC Calculator").Range("E26").Value
MC3 = Worksheets("PEC Calculator").Range("E29").Value

Set tblA = Worksheets("PEC Calculator").ListObjects("ATableINPUT")

' PRODUCTION //////////////////////////////////////////////////////////////
'Air ----------------------------------------------------------------------
If tblA.Range(2, 2).Value = "TableA1" Then
    If Range("D4").Value = "Batch" Then
    tblA.Range(3, 2) = 0.000001
    Else
    tblA.Range(3, 2) = 0.000001
    End If
ElseIf tblA.Range(2, 2).Value = "TableA1.1" Then
    If MC1 = "1b" Then
        If VP < 1 Then
        tblA.Range(3, 2) = 0
        ElseIf VP >= 1 And VP < 10 Then
        tblA.Range(3, 2) = 0
        ElseIf VP >= 10 And VP < 100 Then
        tblA.Range(3, 2) = 0.00001
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 2) = 0.0001
        ElseIf VP >= 1000 And VP < 10000 Then
        tblA.Range(3, 2) = 0.001
        ElseIf VP >= 10000 Then
        tblA.Range(3, 2) = 0.005
        End If
    ElseIf MC1 = "1c" Then
        If VP < 1 Then
        tblA.Range(3, 2) = 0
        ElseIf VP >= 1 And VP < 10 Then
        tblA.Range(3, 2) = 0.00001
        ElseIf VP >= 10 And VP < 100 Then
        tblA.Range(3, 2) = 0.0001
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 2) = 0.001
        ElseIf VP >= 1000 And VP < 10000 Then
        tblA.Range(3, 2) = 0.005
        ElseIf VP >= 10000 Then
        tblA.Range(3, 2) = 0.01
        End If
    ElseIf MC1 = "3" Then
        If VP < 1 Then
        tblA.Range(3, 2) = 0.00001
        ElseIf VP >= 1 And VP < 10 Then
        tblA.Range(3, 2) = 0.0001
        ElseIf VP >= 10 And VP < 100 Then
        tblA.Range(3, 2) = 0.001
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 2) = 0
        ElseIf VP >= 1000 And VP < 10000 Then
        tblA.Range(3, 2) = 0.05
        ElseIf VP >= 10000 Then
        tblA.Range(3, 2) = 0.05
        End If
    ElseIf Worksheets("PEC Calculator").Range("C23").Value = "Please Select" Then
    tblA.Range(3, 2) = "Select main category for production"
    End If
ElseIf tblA.Range(2, 2).Value = "TableA1.2" Then 'For UC=33
    If MC1 = "1a" Then
        If VP < 1 Then
        tblA.Range(3, 2) = 0
        ElseIf VP >= 1 And VP < 10 Then
        tblA.Range(3, 2) = 0
        ElseIf VP >= 10 And VP < 100 Then
        tblA.Range(3, 2) = 0
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 2) = 0.00001
        ElseIf VP >= 1000 And VP < 10000 Then
        tblA.Range(3, 2) = 0.0001
        ElseIf VP >= 10000 Then
        tblA.Range(3, 2) = 0.001
        End If
    ElseIf MC1 = "1b" Then
        If VP < 1 Then
        tblA.Range(3, 2) = 0
        ElseIf VP >= 1 And VP < 10 Then
        tblA.Range(3, 2) = 0
        ElseIf VP >= 10 And VP < 100 Then
        tblA.Range(3, 2) = 0.00001
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 2) = 0.0001
        ElseIf VP >= 1000 And VP < 10000 Then
        tblA.Range(3, 2) = 0.001
        ElseIf VP >= 10000 Then
        tblA.Range(3, 2) = 0.01
        End If
    ElseIf MC1 = "1c" Then
        If VP < 1 Then
        tblA.Range(3, 2) = 0
        ElseIf VP >= 1 And VP < 10 Then
        tblA.Range(3, 2) = 0.00001
        ElseIf VP >= 10 And VP < 100 Then
        tblA.Range(3, 2) = 0.0001
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 2) = 0.001
        ElseIf VP >= 1000 And VP < 10000 Then
        tblA.Range(3, 2) = 0.01
        ElseIf VP >= 10000 Then
        tblA.Range(3, 2) = 0.025
        End If
    ElseIf Worksheets("PEC Calculator").Range("C23").Value = "Please Select" Or Worksheets("PEC Calculator").Range("C23").Value = "" Then
    tblA.Range(3, 2) = "Select main category for production"
    End If
ElseIf tblA.Range(2, 2).Value = "TableA1.3" And Use = "Colouring agents" Then
    tblA.Range(3, 2) = 0.0008
End If
        
'Wastewater ---------------------------------------------------------------------
If tblA.Range(2, 2).Value = "TableA1.1" Then
    If Tprod < 1000 Then
    tblA.Range(4, 2) = 0.02
    ElseIf Tprod >= 1000 Then
    tblA.Range(4, 2) = 0.003
    End If
ElseIf tblA.Range(2, 2).Value = "TableA1" Then
    If Range("D4").Value = "Batch" Then
    tblA.Range(4, 2) = 0.003
    Else
    tblA.Range(4, 2) = 0.001
    End If
ElseIf tblA.Range(2, 2).Value = "TableA1.2" Then
    If Range("D4").Value = "Wet" Then
        If Tprod < 1000 Then
            tblA.Range(4, 2) = 0.02
        ElseIf Tprod >= 1000 Then
            tblA.Range(4, 2) = 0.003
        End If
    Else
    tblA.Range(4, 2) = 0
    End If
ElseIf tblA.Range(2, 2).Value = "TableA1.3" Then
    If WS < 2000 Then
    tblA.Range(4, 2) = 0.015
    ElseIf WS < 10000 And WS >= 2000 Then
    tblA.Range(4, 2) = 0.02
    ElseIf WS >= 10000 And WS < 100000 Then
    tblA.Range(4, 2) = 0.03
    ElseIf WS >= 100000 And WS < 500000 Then
    tblA.Range(4, 2) = 0.05
    ElseIf WS >= 500000 Then
    tblA.Range(4, 2) = 0.06
    End If
End If

'Soil ------------------------------------------------------------------------------
If tblA.Range(2, 2).Value = "TableA1" Then
    tblA.Range(5, 2) = 0
ElseIf tblA.Range(2, 2).Value = "TableA1.2" Then
    If MC1 = "1a" Then
    tblA.Range(5, 2) = 0
    ElseIf MC1 = "1b" Then
    tblA.Range(5, 2) = 0.00001
    ElseIf MC1 = "1c" Then
    tblA.Range(5, 2) = 0.0001
    End If
Else
tblA.Range(5, 2) = 0.0001
End If


' FORMULATION //////////////////////////////////////////////////////////////
'Air -----------------------------------------------------------------------
If tblA.Range(2, 3).Value = "TableA2" Then
    If Range("D5").Value = "Regular powder" Then
    tblA.Range(3, 3) = 0.0002
    ElseIf Range("D5").Value = "Compact powder" Then
    tblA.Range(3, 3) = 0.0002
    ElseIf Range("D5").Value = "Liquid" Then
    tblA.Range(3, 3) = 0.00002
    Else
    tblA.Range(3, 3) = 0.0002
    End If
ElseIf tblA.Range(2, 3).Value = "TableA2.1" Then
    If MC2 = "1b" Then
        If VP < 10 Then
        tblA.Range(3, 3) = 0.0005
        ElseIf VP >= 10 And VP < 100 Then
        tblA.Range(3, 3) = 0.001
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 3) = 0.0025
        ElseIf VP >= 1000 Then
        tblA.Range(3, 3) = 0.005
        End If
    ElseIf MC2 = "1c" Then
        If VP < 10 Then
        tblA.Range(3, 3) = 0.001
        ElseIf VP >= 10 And VP < 100 Then
        tblA.Range(3, 3) = 0.0025
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 3) = 0.005
        ElseIf VP >= 1000 Then
        tblA.Range(3, 3) = 0.01
        End If
    ElseIf MC2 = "3" Then
        If VP < 10 Then
        tblA.Range(3, 3) = 0.0025
        ElseIf VP >= 10 And VP < 100 Then
        tblA.Range(3, 3) = 0.005
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 3) = 0.01
        ElseIf VP >= 1000 Then
        tblA.Range(3, 3) = 0.025
        End If
    ElseIf Worksheets("PEC Calculator").Range("C26").Value = "Please Select" Or Worksheets("PEC Calculator").Range("C26").Value = "" Then
    tblA.Range(3, 3) = "Select main category for formulation"
    End If
ElseIf tblA.Range(2, 3).Value = "TableA2.2" Then
    If VP < 1 Then
    tblA.Range(3, 3) = 0.00005
    ElseIf VP >= 1 And VP < 10 Then
    tblA.Range(3, 3) = 0.00001
    ElseIf VP >= 10 And VP < 100 Then
    tblA.Range(3, 3) = 0.0005
    ElseIf VP >= 100 And VP < 1000 Then
    tblA.Range(3, 3) = 0.0025
    ElseIf VP >= 1000 Then
    tblA.Range(3, 3) = 0.025
    End If
ElseIf tblA.Range(2, 3).Value = "Table A2.3" Then
    If VP < 1 Then
    tblA.Range(3, 3) = 0.0001
    ElseIf VP >= 1 And VP < 10 Then
    tblA.Range(3, 3) = 0.001
    ElseIf VP >= 10 And VP < 100 Then
    tblA.Range(3, 3) = 0.3
    ElseIf VP >= 100 And VP < 1000 Then
    tblA.Range(3, 3) = 0.7
    ElseIf VP >= 1000 Then
    tblA.Range(3, 3) = 1
    End If
End If

'Wastewater -------------------------------------------------------
If tblA.Range(2, 3).Value = "TableA2" Then
    If Range("D5").Value = "Regular powder" Then
    tblA.Range(4, 3) = 0.0001
    ElseIf Range("D5").Value = "Compact powder" Then
    tblA.Range(4, 3) = 0.00001
    ElseIf Range("D5").Value = "Liquid" Then
    tblA.Range(4, 3) = 0.0009
    Else
    tblA.Range(4, 3) = 0.0009
    End If
ElseIf tblA.Range(2, 3).Value = "TableA2.2" Then
    tblA.Range(4, 3) = 0.002
ElseIf tblA.Range(2, 3).Value = "TableA2.3" Then
    If Worksheets("PEC Calculator").Range("D4").Value = "Control of crystal growth" Then
    tblA.Range(4, 3) = 0.99
    Else
    tblA.Range(4, 3) = 0.002
    End If
ElseIf Tprod < 1000 Then
    tblA.Range(4, 3) = 0.02
    ElseIf Tprod >= 1000 Then
    tblA.Range(4, 3) = 0.003
End If

'Soil ---------------------------------------------------------------------
If tblA.Range(2, 3).Value = "TableA2" Then
    tblA.Range(5, 3) = 0
ElseIf tblA.Range(2, 3).Value = "TableA2.2" Then
    tblA.Range(5, 3) = 0.00001
ElseIf tblA.Range(2, 2).Value = "TableA1.3" Then
    tblA.Range(5, 3) = 0.00025
Else
tblA.Range(5, 3) = 0.0001
End If

' INDUSTRIAL PROCESSING //////////////////////////////////////////////////////////////
' Air ---------------------------------------------------------------------
If tblA.Range(2, 4).Value = "TableA3.1" Then 'Agriculture only
    If Use = "Aerosol Propellants" Or Use = "Solvents" Then
    tblA.Range(3, 4) = 1
    ElseIf Use = "Cleaning/washing agents and additives" Or _
    Use = "Colouring agents" Or Use = "Odour agents" Or Use = "Fertilisers" Or _
    Use = "Food/feedstuff additives" Or Use = "Pharmaceuticals" Then
    tblA.Range(3, 4) = 0
    ElseIf Use = "Plant protection products-agricultural" Or Use = "Surface-active agents" Then
    tblA.Range(3, 4) = 0.05
    Else
    tblA.Range(3, 4) = 0.1
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.15" Then
    If Worksheets("PEC Calculator").Range("D4").Value = "Water-based" Then
        If Use = "Colouring agents" Or Use = "Corrosion inhibitors" Or Use = "Fillers" Or Use = "Surface-active agents" Then
        tblA.Range(3, 4) = 0
        ElseIf Use = "Softeners" Or Use = "Viscosity adjustors" Or Use = "Other" Then
            If VP < 10 Then
            tblA.Range(3, 4) = 0
            ElseIf VP >= 10 And VP < 500 Then
            tblA.Range(3, 4) = 0
            ElseIf VP >= 500 And VP < 5000 Then
            tblA.Range(3, 4) = 0.01
            ElseIf VP >= 5000 Then
            tblA.Range(3, 4) = 0.05
            End If
        ElseIf Use = "Solvents" Then
        tblA.Range(3, 4) = 0.8
        Else
        tblA.Range(3, 4) = "Select an appropriate use"
        End If
    ElseIf Worksheets("PEC Calculator").Range("D4").Value = "Solvent-based" Then
        If Use = "Aerosol Propellants" Then
        tblA.Range(3, 4) = 1
        ElseIf Use = "Colouring agents" Or Use = "Corrosion inhibitors" Or Use = "Fillers" Or Use = "Surface-active agents" Then
        tblA.Range(3, 4) = 0
        ElseIf Use = "Softeners" Or Use = "Viscosity adjustors" Or Use = "Other" Then
            If VP < 10 Then
            tblA.Range(3, 4) = 0
            ElseIf VP >= 10 And VP < 500 Then
            tblA.Range(3, 4) = 0.001
            ElseIf VP >= 500 And VP < 5000 Then
            tblA.Range(3, 4) = 0.05
            ElseIf VP >= 5000 Then
            tblA.Range(3, 4) = 0.15
            End If
        ElseIf Use = "Solvents" Then
        tblA.Range(3, 4) = 0.9
        Else
        tblA.Range(3, 4) = "Select an appropriate use"
        End If
    Else
    tblA.Range(3, 4) = "Specify water or solvent-based"
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.16" Then
         If WS < 100 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0.0001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.01
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.01
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.1
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0.01
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.1
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.25
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0.1
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.5
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.7
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0.5
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.75
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.9
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            End If
        ElseIf WS >= 100 And WS < 1000 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0.00001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.0001
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.001
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0.0001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.05
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.05
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.1
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0.05
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.1
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.5
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0.25
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.5
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.75
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            End If
        ElseIf WS >= 1000 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.00001
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.0001
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0.00001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.0001
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.001
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0.0001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.01
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.01
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.1
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(3, 4) = 0.01
                ElseIf MC3 = "3" Then
                tblA.Range(3, 4) = 0.1
                ElseIf MC3 = "4" Then
                tblA.Range(3, 4) = 0.5
                Else
                tblA.Range(3, 4) = "Select main category for processing"
                End If
            End If
        End If
ElseIf tblA.Range(2, 4).Value = "TableA3.2" Then 'Chem industry basic only
    If WS < 100 Then
        If VP < 100 Then
        tblA.Range(3, 4) = 0.65
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 4) = 0.8
        ElseIf VP >= 1000 Then
        tblA.Range(3, 4) = 0.95
        End If
    ElseIf WS >= 100 And WS < 1000 Then
        If VP < 100 Then
        tblA.Range(3, 4) = 0.4
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 4) = 0.55
        ElseIf VP >= 1000 Then
        tblA.Range(3, 4) = 0.65
        End If
    ElseIf WS >= 1000 And WS < 10000 Then
        If VP < 100 Then
        tblA.Range(3, 4) = 0.25
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 4) = 0.35
        ElseIf VP >= 1000 Then
        tblA.Range(3, 4) = 0.5
        End If
    ElseIf WS >= 10000 Then
        If VP < 100 Then
        tblA.Range(3, 4) = 0.05
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 4) = 0.1
        ElseIf VP >= 1000 Then
        tblA.Range(3, 4) = 0.25
        End If
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.3" Then 'Chemical Industry only
    If MC3 = "1b" Then
        If VP < 1 Then
        tblA.Range(3, 4) = 0
        ElseIf VP >= 1 And VP < 10 Then
        tblA.Range(3, 4) = 0
        ElseIf VP >= 10 And VP < 100 Then
        tblA.Range(3, 4) = 0
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 4) = 0.00001
        ElseIf VP >= 1000 And VP < 10000 Then
        tblA.Range(3, 4) = 0.0001
        ElseIf VP >= 10000 Then
        tblA.Range(3, 4) = 0.001
        End If
    ElseIf MC3 = "1c" Then
        If VP < 1 Then
        tblA.Range(3, 4) = 0
        ElseIf VP >= 1 And VP < 10 Then
        tblA.Range(3, 4) = 0
        ElseIf VP >= 10 And VP < 100 Then
        tblA.Range(3, 4) = 0.00001
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 4) = 0.0001
        ElseIf VP >= 1000 And VP < 10000 Then
        tblA.Range(3, 4) = 0.001
        ElseIf VP >= 10000 Then
        tblA.Range(3, 4) = 0.005
        End If
    ElseIf MC3 = "3" Then
        If VP < 1 Then
        tblA.Range(3, 4) = 0.00001
        ElseIf VP >= 1 And VP < 10 Then
        tblA.Range(3, 4) = 0.0001
        ElseIf VP >= 10 And VP < 100 Then
        tblA.Range(3, 4) = 0.001
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(3, 4) = 0.01
        ElseIf VP >= 1000 And VP < 10000 Then
        tblA.Range(3, 4) = 0.025
        ElseIf VP >= 10000 Then
        tblA.Range(3, 4) = 0.05
        End If
    ElseIf Worksheets("PEC Calculator").Range("C29").Value = "Please Select" Or Worksheets("PEC Calculator").Range("C29").Value = "" Then
        tblA.Range(3, 4) = "Select main category for processing"
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.4" Then 'Electrical Industry only
    If WS < 100 Then
        If MC3 = "2" Or MC3 = "3" Then
        tblA.Range(3, 4) = 0.0005
        ElseIf Worksheets("PEC Calculator").Range("C29").Value = "Please Select" Or Worksheets("PEC Calculator").Range("C29").Value = "" Then
        tblA.Range(3, 4) = "Select main category for processing"
        End If
    ElseIf WS >= 100 Then
        If MC3 = "2" Then
        tblA.Range(3, 4) = 0.0005
        ElseIf MC3 = "3" Then
        tblA.Range(3, 4) = 0.001
        ElseIf Worksheets("PEC Calculator").Range("C29").Value = "Please Select" Or Worksheets("PEC Calculator").Range("C29").Value = "" Then
        tblA.Range(3, 4) = "Select main category for processing"
        End If
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.5" Then
    If Use = "Biocides, non-agricultural" Then
    tblA.Range(3, 4) = 0.1
    ElseIf Use = "Cleaning/washing agents and additives" Then
        If T <= 1000 Then
        tblA.Range(3, 4) = 0.0025
        Else
        tblA.Range(3, 4) = 0
        End If
    Else
    tblA.Range(3, 4) = 0.05
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.6" Then
    If WS < 100 And VP < 100 Then
    tblA.Range(3, 4) = 0.001
    ElseIf WS < 100 And VP >= 100 Then
    tblA.Range(3, 4) = 0.01
    ElseIf WS >= 100 Then
    tblA.Range(3, 4) = 0
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.7" Then
    If Use = "Heat transferring agents" Or Use = "Lubricants and additives" Then
        If HLC < 2 Then
        tblA.Range(3, 4) = 0.0002
        ElseIf HLC >= 2 Then
        tblA.Range(3, 4) = 0.002
        End If
    Else
        If MC3 = "2" Then
        tblA.Range(3, 4) = 0
        ElseIf MC3 = "3" Then
        tblA.Range(3, 4) = 0.25
        Else
        tblA.Range(3, 4) = "Select main category for processing"
        End If
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.8" Then
    If VP < 1 Then
    tblA.Range(3, 4) = 0.0001
    ElseIf VP >= 1 And VP < 10 Then
    tblA.Range(3, 4) = 0.0005
    ElseIf VP >= 10 And VP < 100 Then
    tblA.Range(3, 4) = 0.001
    ElseIf VP >= 100 And VP < 1000 Then
    tblA.Range(3, 4) = 0.005
    ElseIf VP >= 1000 Then
    tblA.Range(3, 4) = 0.01
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.9" Then
    If Worksheets("PEC Calculator").Range("D6").Value = "Yes" Then
    tblA.Range(3, 4) = 0
    Else
        If MC3 = "3" Then
            If VP < 1 Then
            tblA.Range(3, 4) = 0.000035
            ElseIf VP >= 1 And VP < 10 Then
            tblA.Range(3, 4) = 0.00025
            ElseIf VP >= 10 And VP < 100 Then
            tblA.Range(3, 4) = 0.0075
            ElseIf VP >= 100 And VP < 1000 Then
            tblA.Range(3, 4) = 0.025
            ElseIf VP >= 1000 Then
            tblA.Range(3, 4) = 0.075
            End If
        Else
        tblA.Range(3, 4) = "Select category 3 for processing"
        End If
    End If
End If

'Wastewater -----------------------------------------------------------
If tblA.Range(2, 4).Value = "TableA3.1" Then
    If Use = "Aerosol Propellants" Or Use = "Solvents" Or Use = "Pharmaceuticals" Then
    tblA.Range(4, 4) = 0
    ElseIf Use = "Fertilisers" Then
    tblA.Range(4, 4) = 0.05
    Else
    tblA.Range(4, 4) = 0.1
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.15" Then
    If Worksheets("PEC Calculator").Range("D4").Value = "Water-based" Then
        If Use = "Colouring agents" Or Use = "Corrosion inhibitors" Or Use = "Fillers" Then
        tblA.Range(4, 4) = 0.005
        ElseIf Use = "Surface-active agents" Then
             If WS < 10 Then
            tblA.Range(4, 4) = 0.005
            ElseIf WS >= 10 And WS < 100 Then
            tblA.Range(4, 4) = 0.01
            ElseIf WS >= 100 Then
            tblA.Range(4, 4) = 0.05
            End If
        ElseIf Use = "Softeners" Or Use = "Viscosity adjustors" Or Use = "Other" Then
            If WS < 10 Then
            tblA.Range(4, 4) = 0.005
            ElseIf WS >= 10 And WS < 100 Then
            tblA.Range(4, 4) = 0.01
            ElseIf WS >= 100 Then
            tblA.Range(4, 4) = 0.05
            End If
        ElseIf Use = "Solvents" Then
        tblA.Range(4, 4) = 0.1
        Else
        tblA.Range(4, 4) = "Select appropriate use"
        End If
    ElseIf Worksheets("PEC Calculator").Range("D4").Value = "Solvent-based" Then
        If Use = "Aerosol Propellants" Then
        tblA.Range(4, 4) = 0
        ElseIf Use = "Colouring agents" Or Use = "Corrosion inhibitors" Or Use = "Fillers" Then
        tblA.Range(4, 4) = 0.001
        ElseIf Use = "Softeners" Or Use = "Viscosity adjustors" Or Use = "Other" Then
            If WS < 10 Then
            tblA.Range(4, 4) = 0.001
            ElseIf WS >= 10 And WS < 100 Then
            tblA.Range(4, 4) = 0.005
            ElseIf WS >= 100 Then
            tblA.Range(4, 4) = 0.01
            End If
        ElseIf Use = "Solvents" Then
        tblA.Range(4, 4) = 0.02
        Else
        tblA.Range(4, 4) = "Select appropriate use"
        End If
    Else
    tblA.Range(4, 4) = "Specify water or solvent-based"
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.16" Then
         If WS < 100 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0.01
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.1
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.5
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.01
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.1
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0.0001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.01
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0.00001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.0001
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.001
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.00001
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.0001
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            End If
        ElseIf WS >= 100 And WS < 1000 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0.25
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.5
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.75
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0.05
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.1
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.5
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.01
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.1
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0.0001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.05
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0.00001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.0001
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.001
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            End If
        ElseIf WS >= 1000 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0.5
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.75
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.9
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0.1
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.5
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.7
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0.01
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.1
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.25
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.01
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.1
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(4, 4) = 0.0001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 4) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(4, 4) = 0.01
                Else
                tblA.Range(4, 4) = "Select main category for processing"
                End If
            End If
        End If
ElseIf tblA.Range(2, 4).Value = "TableA3.2" Then
    If WS < 100 Then
        If VP < 100 Then
        tblA.Range(4, 4) = 0.25
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(4, 4) = 0.1
        ElseIf VP >= 1000 Then
        tblA.Range(4, 4) = 0.05
        End If
    ElseIf WS >= 100 And WS < 1000 Then
        If VP < 100 Then
        tblA.Range(4, 4) = 0.5
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(4, 4) = 0.35
        ElseIf VP >= 1000 Then
        tblA.Range(4, 4) = 0.25
        End If
    ElseIf WS >= 1000 And WS < 10000 Then
        If VP < 100 Then
        tblA.Range(4, 4) = 0.65
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(4, 4) = 0.55
        ElseIf VP >= 1000 Then
        tblA.Range(4, 4) = 0.4
        End If
    ElseIf WS >= 10000 Then
        If VP < 100 Then
        tblA.Range(4, 4) = 0.85
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(4, 4) = 0.8
        ElseIf VP >= 1000 Then
        tblA.Range(4, 4) = 0.65
        End If
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.3" Then
    If Range("D4").Value = "Wet" Then
        If T < 1000 Then
        tblA.Range(4, 4) = 0.02
        ElseIf T >= 1000 Then
        tblA.Range(4, 4) = 0.007
        End If
    ElseIf Range("D4").Value = "Dry" Then
    tblA.Range(4, 4) = 0
    Else
    tblA.Range(4, 4) = "Specify wet or dry"
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.4" Then
    If MC3 = "2" Then
    tblA.Range(4, 4) = 0.0001
    ElseIf MC3 = "3" Then
    tblA.Range(4, 4) = 0.005
    ElseIf Worksheets("PEC Calculator").Range("C29").Value = "Please Select" Or Worksheets("PEC Calculator").Range("C29").Value = "" Then
    tblA.Range(4, 4) = "Select main category for processing"
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.5" Then
    If Use = "Biocides, non-agricultural" Then
    tblA.Range(4, 4) = 1
    ElseIf Use = "Cleaning/washing agents and additives" Then
        If T <= 1000 Then
        tblA.Range(4, 4) = 0.9
        Else
        tblA.Range(4, 4) = 1
        End If
    Else
    tblA.Range(4, 4) = 0.45
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.6" Then
    If MC3 = "2" Then
        If WS < 100 Then
        tblA.Range(4, 4) = 0.05
        ElseIf WS >= 100 And WS < 1000 Then
        tblA.Range(4, 4) = 0.15
        ElseIf WS >= 1000 Then
        tblA.Range(4, 4) = 0.25
        End If
    ElseIf MC3 = "3" Then
        If WS < 100 Then
        tblA.Range(4, 4) = 0.9
        ElseIf WS >= 100 And WS < 1000 Then
        tblA.Range(4, 4) = 0.99
        ElseIf WS >= 1000 Then
        tblA.Range(4, 4) = 0.99
        End If
    ElseIf Worksheets("PEC Calculator").Range("C29").Value = "Please Select" Or Worksheets("PEC Calculator").Range("C29").Value = "" Then
        tblA.Range(4, 4) = "Select main category for processing"
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.7" Then
    If Use = "Heat transferring agents" Or Use = "Lubricants and additives" Then
        If Worksheets("PEC Calculator").Range("D4").Value = "Pure oils" Then
        tblA.Range(4, 4) = 0.185
        ElseIf Worksheets("PEC Calculator").Range("D4").Value = "Water-based + unknown" Then
        tblA.Range(4, 4) = 0.316
        Else
        tblA.Range(4, 4) = "Specify condition"
        End If
    Else
        If WS < 100 Then
            If MC3 = "2" Then
            tblA.Range(4, 4) = 0.05
            ElseIf MC3 = "3" Then
            tblA.Range(4, 4) = 0.5
            Else
            tblA.Range(4, 4) = "Select main category for processing"
            End If
        ElseIf WS >= 100 And WS < 1000 Then
            If MC3 = "2" Then
            tblA.Range(4, 4) = 0.1
            ElseIf MC3 = "3" Then
            tblA.Range(4, 4) = 0.5
            Else
            tblA.Range(4, 4) = "Select main category for processing"
            End If
        ElseIf WS >= 1000 Then
            If MC3 = "2" Then
            tblA.Range(4, 4) = 0.25
            ElseIf MC3 = "3" Then
            tblA.Range(4, 4) = 0.5
            Else
            tblA.Range(4, 4) = "Select main category for processing"
            End If
        End If
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.8" Then
    tblA.Range(4, 4) = 0.0005
ElseIf tblA.Range(2, 4).Value = "TableA3.9" Then
    If Worksheets("PEC Calculator").Range("D6").Value = "Yes" Then
    tblA.Range(4, 4) = 0
    ElseIf Worksheets("PEC Calculator").Range("D6").Value = "No" Then
        If Worksheets("PEC Calculator").Range("D7").Value = "Coupler of dye" Then
        tblA.Range(4, 4) = 0.15
        Else
        tblA.Range(4, 4) = 0.8
        End If
    Else
    tblA.Range(4, 4) = "Specify whether solid or aqueous"
    End If
End If

'Soil --------------------------------------------------------------------
If tblA.Range(2, 4).Value = "TableA3.1" Then
    If Use = "Aerosol Propellants" Or Use = "Solvents" Then
    tblA.Range(5, 4) = 0
    ElseIf Use = "Cleaning/washing agents and additives" Or Use = "Colouring agents" Or Use = "Odour agents" Then
    tblA.Range(5, 4) = 0.4
    ElseIf Use = "Food/feedstuff additives" Then
    tblA.Range(5, 4) = 0.05
    ElseIf Use = "Plant protectionproducts-agricultural" Or Use = "Surface-active agents" Then
    tblA.Range(5, 4) = 0.85
    ElseIf Use = "Fertilisers" Then
    tblA.Range(5, 4) = 0.95
    Else
    tblA.Range(5, 4) = 0.8
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.15" Then
    If Worksheets("PEC Calculator").Range("D4").Value = "Water-based" Then
        If Use = "Colouring agents" Or Use = "Corrosion inhibitors" Or Use = "Fillers" Or _
        Use = "Surface-active agents" Or Use = "Softeners" Or Use = "Viscosity adjustors" Or Use = "Other" Then
        tblA.Range(5, 4) = 0.005
        ElseIf Use = "Solvents" Then
        tblA.Range(5, 4) = 0.001
        Else
        tblA.Range(5, 4) = "Select appropriate use"
        End If
    ElseIf Worksheets("PEC Calculator").Range("D4").Value = "Solvent-based" Then
        If Use = "Aerosol Propellants" Then
        tblA.Range(5, 4) = 0
        ElseIf Use = "Colouring agents" Or Use = "Corrosion inhibitors" Or Use = "Fillers" Or _
        Use = "Softeners" Or Use = "Viscosity adjustors" Or Use = "Other" Then
        tblA.Range(5, 4) = 0.005
        ElseIf Use = "Solvents" Then
        tblA.Range(5, 4) = 0.001
        Else
        tblA.Range(5, 4) = "Select appropriate use"
        End If
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.16" Then
         If WS < 100 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0.005
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0.01
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0.05
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0.005
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0.01
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0.0005
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0.005
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0.0005
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0.001
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0.0005
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            End If
        ElseIf WS >= 100 And WS < 1000 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0.005
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0.01
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0.0005
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0.005
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0.0005
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0.001
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0.0005
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0.0001
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            End If
        ElseIf WS >= 1000 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0.0005
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0.005
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0.0005
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0.001
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0.0005
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0.0001
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 4) = 0
                ElseIf MC3 = "4" Then
                tblA.Range(5, 4) = 0
                Else
                tblA.Range(5, 4) = "Select main category for processing"
                End If
            End If
        End If
ElseIf tblA.Range(2, 4).Value = "TableA3.2" Then
    If WS < 100 Then
        If VP < 100 Then
        tblA.Range(5, 4) = 0.0005
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(5, 4) = 0.0025
        ElseIf VP >= 1000 Then
        tblA.Range(5, 4) = 0.001
        End If
    ElseIf WS >= 100 And WS < 1000 Then
        If VP < 100 Then
        tblA.Range(5, 4) = 0.005
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(5, 4) = 0.002
        ElseIf VP >= 1000 Then
        tblA.Range(5, 4) = 0.001
        End If
    ElseIf WS >= 1000 And WS < 10000 Then
        If VP < 100 Then
        tblA.Range(5, 4) = 0.005
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(5, 4) = 0.002
        ElseIf VP >= 1000 Then
        tblA.Range(5, 4) = 0.001
        End If
    ElseIf WS >= 10000 Then
        If VP < 100 Then
        tblA.Range(5, 4) = 0.005
        ElseIf VP >= 100 And VP < 1000 Then
        tblA.Range(5, 4) = 0.002
        ElseIf VP >= 1000 Then
        tblA.Range(5, 4) = 0.001
        End If
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.4" Then
    If MC3 = "2" Then
    tblA.Range(5, 4) = 0.0001
    ElseIf MC3 = "3" Then
    tblA.Range(5, 4) = 0.01
    ElseIf Worksheets("PEC Calculator").Range("C29").Value = "Please Select" Or Worksheets("PEC Calculator").Range("C29").Value = "" Then
    tblA.Range(4, 4) = "Select main category for processing"
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.5" Then
    If Use = "Biocides, non-agricultural" Then
    tblA.Range(5, 4) = 8
    ElseIf Use = "Cleaning/washing agents and additives" Then
        If Worksheets("PEC Calculator").Range("B14").Value <= 1000 Then
        tblA.Range(5, 4) = 0.05
        Else
        tblA.Range(5, 4) = 0
        End If
    Else
    tblA.Range(5, 4) = 0.45
    End If
ElseIf tblA.Range(2, 4).Value = "TableA3.6" Then
    tblA.Range(5, 4) = 0.01
ElseIf tblA.Range(2, 4).Value = "TableA3.7" Then
    tblA.Range(5, 4) = 0.0001
ElseIf tblA.Range(2, 4).Value = "TableA3.8" Then
    tblA.Range(5, 4) = 0.001
ElseIf tblA.Range(2, 4).Value = "TableA3.9" Then
 If Worksheets("PEC Calculator").Range("D6").Value = "Yes" Then
    tblA.Range(5, 4) = 0
    Else
        If MC3 = "3" Then
        tblA.Range(5, 4) = 0.00025
        Else
        tblA.Range(5, 4) = "Select main category for processing"
        End If
    End If
End If

Worksheets("PEC Calculator").Protect "eau"


End Sub


