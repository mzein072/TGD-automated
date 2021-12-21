Sub ATables_1FB()

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

' PRIVATE USE //////////////////////////////////////////////////////////////
' Air ---------------------------------------------------------------------
If tblA.Range(2, 5).Value = "TableA3.16" Then
        If WS < 100 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0.0001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.01
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.01
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.1
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0.01
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.1
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.25
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0.1
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.5
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.7
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0.5
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.75
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.9
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            End If
        ElseIf WS >= 100 And WS < 1000 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0.00001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.0001
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.001
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0.0001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.05
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.05
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.1
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0.05
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.1
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.5
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0.25
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.5
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.75
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            End If
        ElseIf WS >= 1000 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.00001
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.0001
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0.00001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.0001
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.001
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0.0001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.01
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.01
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.1
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(3, 5) = 0.01
                ElseIf MC3 = "3" Then
                tblA.Range(3, 5) = 0.1
                ElseIf MC3 = "4" Then
                tblA.Range(3, 5) = 0.5
                Else
                tblA.Range(3, 5) = "Select main category for processing"
                End If
            End If
        End If
ElseIf tblA.Range(2, 5).Value = "TableA4.5" Then
    If Worksheets("PEC Calculator").Range("D4").Value = "Water-based" Then
        If Use = "Colouring agents" Or Use = "Corrosion inhibitors" Or Use = "Fillers" Or Use = "Surface-active agents" Then
        tblA.Range(3, 5) = 0
        ElseIf Use = "Softeners" Or Use = "Viscosity adjustors" Or Use = "Other" Then
            If VP < 10 Then
            tblA.Range(3, 5) = 0
            ElseIf VP >= 10 And VP < 500 Then
            tblA.Range(3, 5) = 0
            ElseIf VP >= 500 And VP < 5000 Then
            tblA.Range(3, 5) = 0.01
            ElseIf VP >= 5000 Then
            tblA.Range(3, 5) = 0.05
            End If
        ElseIf Use = "Solvents" Then
        tblA.Range(3, 5) = 0.8
        Else
        tblA.Range(3, 5) = "Select an appropriate use"
        End If
    ElseIf Worksheets("PEC Calculator").Range("D4").Value = "Solvent-based" Then
        If Use = "Aerosol Propellants" Then
        tblA.Range(3, 5) = 1
        ElseIf Use = "Colouring agents" Or Use = "Corrosion inhibitors" Or Use = "Fillers" Or Use = "Surface-active agents" Then
        tblA.Range(3, 5) = 0
        ElseIf Use = "Softeners" Or Use = "Viscosity adjustors" Or Use = "Other" Then
            If VP < 10 Then
            tblA.Range(3, 5) = 0
            ElseIf VP >= 10 And VP < 500 Then
            tblA.Range(3, 5) = 0.001
            ElseIf VP >= 500 And VP < 5000 Then
            tblA.Range(3, 5) = 0.05
            ElseIf VP >= 5000 Then
            tblA.Range(3, 5) = 0.15
            End If
        ElseIf Use = "Solvents" Then
        tblA.Range(3, 5) = 0.95
        Else
        tblA.Range(3, 5) = "Select an appropriate use"
        End If
    Else
    tblA.Range(3, 5) = "Specify water or solvent-based"
    End If
ElseIf tblA.Range(2, 5).Value = "TableA4.2" Then
    If VP < 10 Then
    tblA.Range(3, 5) = 0.005
    ElseIf VP >= 10 And VP < 100 Then
    tblA.Range(3, 5) = 0.015
    ElseIf VP >= 100 And VP < 1000 Then
    tblA.Range(3, 5) = 0.15
    ElseIf VP >= 1000 And VP < 10000 Then
    tblA.Range(3, 5) = 0.4
    ElseIf VP >= 10000 Then
    tblA.Range(3, 5) = 0.6
    End If
ElseIf tblA.Range(2, 5).Value = "TableA4.3" Then
tblA.Range(3, 5) = 0
Else
tblA.Range(3, 5) = ""
End If

' Wastewater ---------------------------------------------------------------------
If tblA.Range(2, 5).Value = "TableA3.16" Then
        If WS < 100 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0.01
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.1
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.5
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.01
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.1
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0.0001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.01
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0.00001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.0001
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.001
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.00001
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.0001
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            End If
        ElseIf WS >= 100 And WS < 1000 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0.25
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.5
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.75
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0.05
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.1
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.5
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.01
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.1
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0.0001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.05
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0.00001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.0001
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.001
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            End If
        ElseIf WS >= 1000 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0.5
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.75
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.9
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0.1
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.5
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.7
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0.01
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.1
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.25
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.01
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.1
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(4, 5) = 0.0001
                ElseIf MC3 = "3" Then
                tblA.Range(4, 5) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(4, 5) = 0.01
                Else
                tblA.Range(4, 5) = "Select main category for processing"
                End If
            End If
        End If
ElseIf tblA.Range(2, 5).Value = "TableA4.5" Then
    If Worksheets("PEC Calculator").Range("D4").Value = "Water-based" Then
        If Use = "Colouring agents" Or Use = "Corrosion inhibitors" Or Use = "Fillers" Then
        tblA.Range(4, 5) = 0.005
        ElseIf Use = "Surface-active agents" Then
             If WS < 10 Then
            tblA.Range(4, 5) = 0.005
            ElseIf WS >= 10 And WS < 100 Then
            tblA.Range(4, 5) = 0.01
            ElseIf WS >= 100 Then
            tblA.Range(4, 5) = 0.05
            End If
        ElseIf Use = "Softeners" Or Use = "Viscosity adjustors" Or Use = "Other" Then
            If WS < 10 Then
            tblA.Range(4, 5) = 0.005
            ElseIf WS >= 10 And WS < 100 Then
            tblA.Range(4, 5) = 0.01
            ElseIf WS >= 100 Then
            tblA.Range(4, 5) = 0.05
            End If
        ElseIf Use = "Solvents" Then
        tblA.Range(4, 5) = 0.15
        Else
        tblA.Range(4, 5) = "Select appropriate use"
        End If
    ElseIf Worksheets("PEC Calculator").Range("D4").Value = "Solvent-based" Then
        If Use = "Aerosol Propellants" Then
        tblA.Range(4, 5) = 0
        ElseIf Use = "Colouring agents" Or Use = "Corrosion inhibitors" Or Use = "Fillers" Then
        tblA.Range(4, 5) = 0.001
        ElseIf Use = "Softeners" Or Use = "Viscosity adjustors" Or Use = "Other" Then
            If WS < 10 Then
            tblA.Range(4, 5) = 0.001
            ElseIf WS >= 10 And WS < 100 Then
            tblA.Range(4, 5) = 0.005
            ElseIf WS >= 100 Then
            tblA.Range(4, 5) = 0.01
            End If
        ElseIf Use = "Solvents" Then
        tblA.Range(4, 5) = 0.04
        Else
        tblA.Range(4, 5) = "Select appropriate use"
        End If
    Else
    tblA.Range(4, 5) = "Specify water or solvent-based"
    End If
ElseIf tblA.Range(2, 5).Value = "TableA4.2" Then
tblA.Range(4, 5) = 0.0005
ElseIf tblA.Range(2, 5).Value = "TableA4.3" Then
tblA.Range(4, 5) = 0.4
Else
tblA.Range(4, 5) = ""
End If

'Soil ------------------------------------------------------
If tblA.Range(2, 5).Value = "TableA4.5" Then
    If Worksheets("PEC Calculator").Range("D4").Value = "Water-based" Then
        If Use = "Colouring agents" Or Use = "Corrosion inhibitors" Or Use = "Fillers" Or _
        Use = "Surface-active agents" Or Use = "Softeners" Or Use = "Viscosity adjustors" Or Use = "Other" Then
        tblA.Range(5, 5) = 0.005
        ElseIf Use = "Solvents" Then
        tblA.Range(5, 5) = 0.01
        Else
        tblA.Range(5, 5) = "Select appropriate use"
        End If
    ElseIf Worksheets("PEC Calculator").Range("D4").Value = "Solvent-based" Then
        If Use = "Aerosol Propellants" Then
        tblA.Range(5, 5) = 0
        ElseIf Use = "Colouring agents" Or Use = "Corrosion inhibitors" Or Use = "Fillers" Or _
        Use = "Softeners" Or Use = "Viscosity adjustors" Or Use = "Other" Then
        tblA.Range(5, 5) = 0.005
        ElseIf Use = "Solvents" Then
        tblA.Range(5, 5) = 0.01
        Else
        tblA.Range(5, 5) = "Select appropriate use"
        End If
    Else
    tblA.Range(5, 5) = "Specify water or solvent-based"
    End If
ElseIf tblA.Range(2, 5).Value = "TableA3.16" Then
        If WS < 100 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0.005
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0.01
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0.05
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0.005
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0.01
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0.0005
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0.005
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0.0005
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0.001
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0.0005
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            End If
        ElseIf WS >= 100 And WS < 1000 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0.001
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0.005
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0.01
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0.0005
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0.005
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0.0005
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0.001
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0.0005
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0.0001
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            End If
        ElseIf WS >= 1000 Then
            If VP < 10 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0.0005
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0.001
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0.005
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10 And VP < 100 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0.0005
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0.001
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 100 And VP < 1000 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0.0005
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 1000 And VP < 10000 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0.0001
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            ElseIf VP >= 10000 Then
                If MC3 = "2" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "3" Then
                tblA.Range(5, 5) = 0
                ElseIf MC3 = "4" Then
                tblA.Range(5, 5) = 0
                Else
                tblA.Range(5, 5) = "Select main category for processing"
                End If
            End If
        End If
ElseIf tblA.Range(2, 5).Value = "TableA4.2" Then
tblA.Range(5, 5) = 0.0001
ElseIf tblA.Range(2, 5).Value = "TableA4.3" Then
tblA.Range(5, 5) = 0
Else
tblA.Range(5, 5) = ""
End If

' WASTE WATER //////////////////////////////////////////////////////////////
' Air ---------------------------------------------------------------------
If tblA.Range(2, 6).Value = "TableA5.1" Then
    If VP < 1 Then
    tblA.Range(3, 6) = 0.000005
    ElseIf VP >= 1 And VP < 10 Then
    tblA.Range(3, 6) = 0.000025
    ElseIf VP >= 10 And VP < 100 Then
    tblA.Range(3, 6) = 0.00075
    ElseIf VP >= 100 And VP < 1000 Then
    tblA.Range(3, 6) = 0.0025
    ElseIf VP >= 1000 Then
    tblA.Range(3, 6) = 0.01
    End If
Else
tblA.Range(3, 6) = ""
End If
    
' Wastewater ---------------------------------------------------------------------
If tblA.Range(2, 6).Value = "TableA5.1" Then
tblA.Range(4, 6) = 0.2
Else
tblA.Range(4, 6) = ""
End If


' Soil ---------------------------------------------------------------------
If tblA.Range(2, 6).Value = "TableA5.1" Then
tblA.Range(5, 6) = 0
Else
tblA.Range(5, 6) = ""
End If


Worksheets("PEC Calculator").Protect "eau"


End Sub

