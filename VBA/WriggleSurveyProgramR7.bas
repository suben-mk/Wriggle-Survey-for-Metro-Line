Attribute VB_Name = "WriggleSurveyProgramR7"
' Topic; Wriggle Survey Program
' Created By; Suben Mukem (SBM) as Survey Engineer.
' Updated; 02/02/2023
'
'Option Base 1
Const Pi As Single = 3.141592654
 
'----------------General Private Function----------------'

'Convert Degrees to Radian.
Private Function DegtoRad(d)

    DegtoRad = d * (Pi / 180)

End Function

'Convert Radian to Degrees.
Private Function RadtoDeg(r)

    RadtoDeg = r * (180 / Pi)

End Function

'Compute Northing and Easting by Local Coordinate (Y, X) , Coordinate of Center and Azimuth.
Private Function CoorYXtoNE(ECL, NCL, AZCL, Y, X, EN)

    Ei = ECL + Y * Sin(DegtoRad(AZCL)) + X * Sin(DegtoRad(90 + AZCL))
    Ni = NCL + Y * Cos(DegtoRad(AZCL)) + X * Cos(DegtoRad(90 + AZCL))
    
    Select Case UCase$(EN)
     Case "E"
             CoorYXtoNE = Ei
     Case "N"
             CoorYXtoNE = Ni
  End Select
End Function 'Coordinate Y,X to N, E

'Compute Local Coordinate (Y, X, L) by Northing and Easting and Azimuth.

Private Function CoorNEtoYXL(ECL, NCL, AZCL, EA, NA, YXL)

    dE = EA - ECL: dN = NA - NCL
    Linear = Sqr(dE ^ 2 + dN ^ 2)
        
    If dN <> 0 Then Q = RadtoDeg(Atn(dE / dN))
      If dN = 0 Then
        If dE > 0 Then
          AZLinear = 90
        ElseIf dE < 0 Then
          AZLinear = 270
        Else
          AZLinear = False
      End If
      
    ElseIf dN > 0 Then
      If dE > 0 Then
          AZLinear = Q
      ElseIf dE < 0 Then
          AZLinear = 360 + Q
      End If
      
    ElseIf dN < 0 Then
          AZLinear = 180 + Q
    End If
    
        Delta = DegtoRad(AZLinear - AZCL)
        Y = Linear * Cos(Delta)
        X = Linear * Sin(Delta)
        
        Select Case UCase$(YXL)
            Case "Y"
                CoorNEtoYXL = Y
            Case "X"
                CoorNEtoYXL = X
            Case "L"
                CoorNEtoYXL = Linear
        End Select
    
End Function 'CoorNEtoYXL

'Compute Distance and Azimuth from 2 Points.
Private Function DirecDistAz(EStart, NStart, EEnd, NEnd, DA)

    dE = EEnd - EStart: dN = NEnd - NStart
    Distance = Sqr(dE ^ 2 + dN ^ 2)
    
    If dN <> 0 Then Q = RadtoDeg(Atn(dE / dN))
      If dN = 0 Then
        If dE > 0 Then
          Azi = 90
        ElseIf dE < 0 Then
          Azi = 270
        Else
          Azi = False
      End If
      
    ElseIf dN > 0 Then
      If dE > 0 Then
          Azi = Q
      ElseIf dE < 0 Then
          Azi = 360 + Q
      End If
      
    ElseIf dN < 0 Then
          Azi = 180 + Q
    End If
    
    Select Case UCase$(DA)
      Case "D"
          DirecDistAz = Distance
      Case "A"
          DirecDistAz = Azi
    End Select

End Function 'DirecDistAz

'Compute pitching
Private Function Pitching(ChStart, ZStrat, ChEnd, ZEnd)

    Pitching = (ZEnd - ZStrat) / (ChEnd - ChStart)

End Function

'Compute Vertical Deviation
Private Function DeviateVt(ChD, ZD, Pitching, ChA, ZA)

    ZFind = ZD + Pitching * (ChA - ChD)
    DeviateVt = ZA - ZFind

End Function
'-----------------End Private Function-----------------'


'-----------------Wriggle Survey Computation (Best-Fit Circle 3D)-----------------'

Sub WriggleSurvey()

    Dim totalWriggle As Long
    Dim totalRing As Long
    totalWrigglePnt = ThisWorkbook.Sheets("Import Wriggle Data").Cells(Rows.Count, 1).End(xlUp).Row - 3
    totalRing = ThisWorkbook.Sheets("Import Wriggle Data").Range("G:G").Cells.SpecialCells(xlCellTypeConstants).Count - 1
    MsgBox "TOTAL WRIGGLE DATA =" & " " & totalWrigglePnt & " & " & "TOTAL RING =" & " " & totalRing
    
    WriggleName = "Wriggle Comp."
    Sheets.Add(After:=Sheets("Import Tunnel Axis (DTA)")).Name = WriggleName
    
    Sheets.Add(After:=Sheets("Wriggle Comp.")).Name = "Wriggle Backup"
    
    Sheets("Import Tunnel Axis (DTA)").Select
    Dim totalDTA As Long
    totalDTA = ThisWorkbook.Sheets("Import Tunnel Axis (DTA)").Cells(Rows.Count, 1).End(xlUp).Row - 3
    
    DTAName = Range("C1")
    Direction = Range("C2")
    If Direction = "DIRECT" Then
        Excavate_Direc = 1
    ElseIf Direction = "REVERSE" Then
        Excavate_Direc = -1
    Else
        Excavate_Direc = 1 'Incase forget to input excavation direction.
    End If

'--------------------------Format Table (Best Fit Circle) -----------------------'
    
    Sheets(WriggleName).Select
    
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.Select
    Selection.RowHeight = 20
    Columns("B:B").Select
    Selection.ColumnWidth = 10
    Columns("C:F").Select
    Selection.ColumnWidth = 15
    Columns("G:H").Select
    Selection.ColumnWidth = 10
    Columns("I:J").Select
    Selection.ColumnWidth = 13
    Range("B2:C2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Font.Bold = True
    Range("D2:E2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    Range("B3:C3").Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("D3:E3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    Range("B5:B6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C5:E5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("F5:F6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("G5:H5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("I5:J5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("B5:J6").Select
    Selection.Font.Bold = True
    Rows("4:4").Select
    Selection.RowHeight = 30
    Range("B4:J4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("A1").Select
'--------------------------Head Table (Best Fit Circle) -----------------------'

    Range("B2").Select
    ActiveCell.FormulaR1C1 = "TUNNEL ALIGNMENT NAME :"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = DTAName
    
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "EXCAVATION DIRECTION  :"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = Direction
  
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "BEST-FIT CIRCLE 3D RESULT"
   
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "RING NO."
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "TUNNEL CENTER"
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "EASTING (M.)"
    Range("D6").Select
    ActiveCell.FormulaR1C1 = "NORTHING (M.)"
    Range("E6").Select
    ActiveCell.FormulaR1C1 = "ELEVATION (M.)"
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "CHAINAGE (M.)"
    Range("G5").Select
    ActiveCell.FormulaR1C1 = "DEVIATION"
    Range("G6").Select
    ActiveCell.FormulaR1C1 = "HOR. (M.)"
    Range("H6").Select
    ActiveCell.FormulaR1C1 = "VER. (M.)"
    Range("I5").Select
    ActiveCell.FormulaR1C1 = "AVERAGE"
    Range("I6").Select
    ActiveCell.FormulaR1C1 = "RADIUS (M.)"
    Range("J6").Select
    ActiveCell.FormulaR1C1 = "DIAMETER (M.)"
    
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "i, j"
    
    Range("B2").Select
  
  
'--------------------------Format Table (Backup) -----------------------'
    Sheets("Wriggle Backup").Select
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 7
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.ColumnWidth = 5
    Selection.RowHeight = 20
    Cells.Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    'TUNNEL CENTER
    Range("C2:E2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    'DEVIATION
    Range("G2:H2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    'AVERAGE
    Range("I2:J2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    'COORDINATE
    Range("K2:BF2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    'DESIGN CENTER
    ActiveWindow.SmallScroll ToRight:=26
    Range("BG2:BK2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    'LOCAL COORDINATE
    Range("BL2:CS2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    'RADIUS
    ActiveWindow.SmallScroll ToRight:=20
    Range("CT2:DI2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    'RDBC
    Range("DJ2:DY2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    'ANGLE
    Range("DZ2:EO2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    'P
    Range("EP2:EQ2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    'Q
    Range("ER2:ES2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
    Rows("2:3").Select
    Selection.Font.Bold = True
    Range("A1").Select
    
'--------------------------Head Table (Backup) -----------------------'
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "TUNNEL CENTER"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "DEVIATION"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "AVERAGE"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "COORDINATE"
    Range("BG2").Select
    ActiveCell.FormulaR1C1 = "DESIGN CENTER"
    Range("BL2").Select
    ActiveCell.FormulaR1C1 = "LOCAL COORDINATE"
    Range("CT2").Select
    ActiveCell.FormulaR1C1 = "RADIUS"
    Range("DJ2").Select
    ActiveCell.FormulaR1C1 = "RDBC"
    Range("DZ2").Select
    ActiveCell.FormulaR1C1 = "ANGLE"
    Range("EP2").Select
    ActiveCell.FormulaR1C1 = "P"
    Range("ER2").Select
    ActiveCell.FormulaR1C1 = "Q"
    
    Range("A3").Select
    Dim Header() As Variant
    Header = Array("INDEX", "RING NO.", "E", "N", "Z", "CH", "DH", "DV", "R", "DIA.", _
                    "E_P1", "N_P1", "Z_P1", "E_P2", "N_P2", "Z_P2", "E_P3", "N_P3", "Z_P3", "E_P4", "N_P4", "Z_P4", _
                    "E_P5", "N_P5", "Z_P5", "E_P6", "N_P6", "Z_P6", "E_P7", "N_P7", "Z_P7", "E_P8", "N_P8", "Z_P8", _
                    "E_P9", "N_P9", "Z_P9", "E_P10", "N_P10", "Z_P10", "E_P11", "N_P11", "Z_P11", "E_P12", "N_P12", "Z_P12", _
                    "E_P13", "N_P13", "Z_P13", "E_P14", "N_P14", "Z_P14", "E_P15", "N_P15", "Z_P15", "E_P16", "N_P16", "Z_P16", _
                    "E", "N", "Z", "R", "DIA.", _
                    "X_C", "Y_C", "X_P1", "Y_P1", "X_P2", "Y_P2", "X_P3", "Y_P3", "X_P4", "Y_P4", _
                    "X_P5", "Y_P5", "X_P6", "Y_P6", "X_P7", "Y_P7", "X_P8", "Y_P8", _
                    "X_P9", "Y_P9", "X_P10", "Y_P10", "X_P11", "Y_P11", "X_P12", "Y_P12", _
                    "X_P13", "Y_P13", "X_P14", "Y_P14", "X_P15", "Y_P15", "X_P16", "Y_P16", _
                    "R_P1", "R_P2", "R_P3", "R_P4", "R_P5", "R_P6", "R_P7", "R_P8", _
                    "R_P9", "R_P10", "R_P11", "R_P12", "R_P13", "R_P14", "R_P15", "R_P16", _
                    "DR_P1", "DR_P2", "DR_P3", "DR_P4", "DR_P5", "DR_P6", "DR_P7", "DR_P8", _
                    "DR_P9", "DR_P10", "DR_P11", "DR_P12", "DR_P13", "DR_P14", "DR_P15", "DR_P16", _
                    "ANG_P1", "ANG_P2", "ANG_P3", "ANG_P4", "ANG_P5", "ANG_P6", "ANG_P7", "ANG_P8", _
                    "ANG_P9", "ANG_P10", "ANG_P11", "ANG_P12", "ANG_P13", "ANG_P14", "ANG_P15", "ANG_P16", _
                    "E", "N", "E", "N", "OFFSET", "NUM.PNT")
    
    For t = LBound(Header) To UBound(Header)
        ActiveCell.Offset(0, t).Value = Header(t)
    Next
    
    Range("A1").Select
  
'--------------------------Compute Wriggle Survey-----------------------'
  
    Sheets("Import Wriggle Data").Select
    
    'numPnt = Range("C1")
    DiaDesign = Range("C2")
    'totalRing = totalWrigglePnt / numPnt
    
    u = 3
    j = 0
    w = 0
    For i = 0 To totalRing - 1
        Sheets("Import Wriggle Data").Select
        Range("A" & u).Select
        
        numPnt = Range("G" & u + 1)
        Debug.Print numPnt
        
        Dim Rngi() As Variant
        Dim Pi() As Variant
        Dim Ei() As Variant
        Dim Ni() As Variant
        Dim Zi() As Variant
        Dim OSi() As Variant
        ReDim Rngi(numPnt)
        ReDim Pi(numPnt)
        ReDim Ei(numPnt)
        ReDim Ni(numPnt)
        ReDim Zi(numPnt)
        ReDim OSi(numPnt)
        
        For k = 1 To numPnt

            Rngi(k) = ActiveCell.Offset(k, 0)
            Pi(k) = ActiveCell.Offset(k, 1)
            Ei(k) = ActiveCell.Offset(k, 2)
            Ni(k) = ActiveCell.Offset(k, 3)
            Zi(k) = ActiveCell.Offset(k, 4)
            OSi(k) = ActiveCell.Offset(k, 5)
            'Debug.Print Rngi(k), Pi(k), Ei(k), Ni(k), Zi(k), OSi(k)
        Next
        
        'Average Prism Offset
        For k = LBound(OSi) To UBound(OSi)
          avgOSi = WorksheetFunction.Average(OSi(k))
        Next
        'Debug.Print avgOSi
        
        'Linear Regression by Least Square
        
        sumRng = 0
        sumE = 0
        sumN = 0
        sumEN = 0
        sumE2 = 0
        For k = 1 To numPnt
        
            sumRng = RngP + Rngi(k)
            sumE = sumE + Ei(k)
            sumN = sumN + Ni(k)
            sumEN = sumEN + Ei(k) * Ni(k)
            sumE2 = sumE2 + Ei(k) ^ 2
            
        Next
        

        m = (sumEN - (sumE * sumN) / numPnt) / (sumE2 - (sumE ^ 2 / numPnt)) 'Slope
        b = (sumN / numPnt) - m * (sumE / numPnt) 'Intercept
        'Debug.Print m, b
        
        'P and Q point on Linear Regression
        minE = Application.Min(Ei) * 0.999999
        maxE = Application.max(Ei) * 1.0000005
        
        EP = minE
        NP = m * minE + b
        EQ = maxE
        NQ = m * maxE + b
        
        'Coordinates on PQ Line and Local coordinates Xi, Yi
        Dim Xi() As Variant 'Offset from P
        Dim Yi() As Variant 'Elevation
        Dim di() As Variant 'Offset from PQ Line
        ReDim Xi(numPnt)
        ReDim Yi(numPnt)
        ReDim di(numPnt)
        
        For k = 1 To numPnt
        
            AzPQ = DirecDistAz(EP, NP, EQ, NQ, "A")
            Xi(k) = CoorNEtoYXL(EP, NP, AzPQ, Ei(k), Ni(k), "Y")
            Yi(k) = Zi(k) + 100 'In case elevation < 0 m.
            di(k) = CoorNEtoYXL(EP, NP, AzPQ, Ei(k), Ni(k), "X")
            'Debug.Print Xi(k), Yi(k), di(k), EP, NP, EQ, NQ
        
        Next
        
        'Best-fit Circle 2D by Kasa Method (least square)
        sumX = 0
        sumY = 0
        sumX2 = 0
        sumY2 = 0
        sumXY = 0
        sumXY2 = 0
        sumX3 = 0
        sumYX2 = 0
        sumY3 = 0
        
        For k = 1 To numPnt
            
            sumX = sumX + Xi(k)
            sumY = sumY + Yi(k)
            sumX2 = sumX2 + Xi(k) ^ 2
            sumY2 = sumY2 + Yi(k) ^ 2
            sumXY = sumXY + Xi(k) * Yi(k)
            sumXY2 = sumXY2 + Xi(k) * Yi(k) ^ 2
            sumX3 = sumX3 + Xi(k) ^ 3
            sumYX2 = sumYX2 + Yi(k) * Xi(k) ^ 2
            sumY3 = sumY3 + Yi(k) ^ 3
            
        Next
        'Debug.Print sumX, sumY, sumX2, sumY2, sumXY, sumXY2, sumX3, sumYX2, sumY3
        
        KM1 = 2 * ((sumX ^ 2) - numPnt * sumX2)
        KM2 = 2 * (sumX * sumY - numPnt * sumXY)
        KM3 = 2 * ((sumY ^ 2) - numPnt * sumY2)
        KM4 = sumX2 * sumX - numPnt * sumX3 + sumX * sumY2 - numPnt * sumXY2
        KM5 = sumX2 * sumY - numPnt * sumY3 + sumY * sumY2 - numPnt * sumYX2
        
        Xc = (KM4 * KM3 - KM5 * KM2) / (KM1 * KM3 - (KM2 ^ 2))
        Yc = (KM1 * KM5 - KM2 * KM4) / (KM1 * KM3 - (KM2 ^ 2))
        Radius = Sqr((Xc ^ 2) + (Yc ^ 2) + (sumX2 - 2 * Xc * sumX + sumY2 - 2 * Yc * sumY) / numPnt)
        'Debug.Print Xc, Yc, Radius
        
        'Coordinates on PQ Line and Local coordinates Xi, Yi
        Dim Ri() As Variant 'Radius of each point
        Dim RDBCi() As Variant 'Deviation Radius of each point
        Dim ANGi() As Variant 'Angle of each point
        ReDim Ri(numPnt)
        ReDim RDBCi(numPnt)
        ReDim ANGi(numPnt)
        
        For k = 1 To numPnt
        
            Ri(k) = Sqr((Xi(k) - Xc) ^ 2 + (Yi(k) - Yc) ^ 2)
            RDBCi(k) = Ri(k) - Radius
            ANGi(k) = DirecDistAz(Xc, Yc, Xi(k), Yi(k), "A")
            'Debug.Print Ri(k), RDBCi(k), ANGi(k)
        
        Next
        
        'Transform center coordinates Xc,Yc to Ec, Nc, Zc
        
        Ec = CoorYXtoNE(EP, NP, AzPQ, Xc, 0, "E")
        Nc = CoorYXtoNE(EP, NP, AzPQ, Xc, 0, "N")
        Zc = Yc - 100
        'Debug.Print Ec, Nc, Zc
        
        'Compute extention data
        Dim extRi() As Variant
        Dim extXi() As Variant
        Dim extYi() As Variant
        Dim extEi() As Variant
        Dim extNi() As Variant
        Dim extZi() As Variant
        ReDim extRi(numPnt)
        ReDim extXi(numPnt)
        ReDim extYi(numPnt)
        ReDim extEi(numPnt)
        ReDim extNi(numPnt)
        ReDim extZi(numPnt)
        
        For k = 1 To numPnt
            'Local coordinate Xi, Yi
            extRi(k) = Ri(k) + OSi(k)
            extXi(k) = Xc + extRi(k) * Sin(DegtoRad(ANGi(k)))
            extYi(k) = Yc + extRi(k) * Cos(DegtoRad(ANGi(k)))
            
            'Grid coordinate Ei, Ni
            extEi(k) = CoorYXtoNE(Ec, Nc, AzPQ, extXi(k) - Xc, di(k), "E")
            extNi(k) = CoorYXtoNE(Ec, Nc, AzPQ, extXi(k) - Xc, di(k), "N")
            extZi(k) = extYi(k) - 100
            'Debug.Print extRi(k), extXi(k), extYi(k), extEi(k), extNi(k), extZi(k)
            
        Next

        'Deviation of Tunnel Center and Chainage
    
        Sheets("Import Tunnel Axis (DTA)").Select
        Range("A4").Select
        
        Dim PntDTA() As Variant
        Dim ChDTA() As Variant
        Dim EDTA() As Variant
        Dim NDTA() As Variant
        Dim ZDTA() As Variant
        Dim Linear() As Variant
        ReDim PntDTA(totalDTA - 1)
        ReDim ChDTA(totalDTA - 1)
        ReDim EDTA(totalDTA - 1)
        ReDim NDTA(totalDTA - 1)
        ReDim ZDTA(totalDTA - 1)
        ReDim Linear(totalDTA - 1)
        
        For d = 0 To totalDTA - 1
    
            PntDTA(d) = ActiveCell.Offset(d, 0)
            ChDTA(d) = ActiveCell.Offset(d, 1)
            EDTA(d) = ActiveCell.Offset(d, 2)
            NDTA(d) = ActiveCell.Offset(d, 3)
            ZDTA(d) = ActiveCell.Offset(d, 4)
            Linear(d) = Sqr((EDTA(d) - Ec) ^ 2 + (NDTA(d) - Nc) ^ 2)
    
        Next
        
        'Find minimum linear from tunnel center to tunnel axis
        minLinear = Application.Min(Linear)
        minIndex = Application.Match(minLinear, Linear, 0) - 1
        'Debug.Print minIndex, minLinear
        'Debug.Print PntDTA(minIndex), ChDTA(minIndex), EDTA(minIndex), NDTA(minIndex), ZDTA(minIndex), Linear(minIndex)
        
        'Point.B is back point, Point.M is middle point (nearly tunnel point), Point.H is ahead point. B------>M------>H
        
        'Point.B ; Point no., Chainage, Easting, Northing, Elevation
        PntB = PntDTA(minIndex - 1)
        ChB = ChDTA(minIndex - 1)
        EB = EDTA(minIndex - 1)
        NB = NDTA(minIndex - 1)
        ZB = ZDTA(minIndex - 1)
        
        'Point.M ; Point no., Chainage, Easting, Northing, Elevation
        PntM = PntDTA(minIndex)
        ChM = ChDTA(minIndex)
        EM = EDTA(minIndex)
        NM = NDTA(minIndex)
        ZM = ZDTA(minIndex)
        
        'Point.H ; Point no., Chainage, Easting, Northing, Elevation
        PntH = PntDTA(minIndex + 1)
        ChH = ChDTA(minIndex + 1)
        EH = EDTA(minIndex + 1)
        NH = NDTA(minIndex + 1)
        ZH = ZDTA(minIndex + 1)
        
        
        DistAC = DirecDistAz(EB, NB, Ec, Nc, "D")
        DistHC = DirecDistAz(EH, NH, Ec, Nc, "D")
        
        DistBM = DirecDistAz(EB, NB, EM, NM, "D")
        AzBM = DirecDistAz(EB, NB, EM, NM, "A")
        PitchBM = Pitching(ChB, ZB, ChM, ZM)
    
        DistMH = DirecDistAz(EM, NM, EH, NH, "D")
        AzMH = DirecDistAz(EM, NM, EH, NH, "A")
        PitchMH = Pitching(ChM, ZM, ChH, ZH)
    
        If DistAC < DistHC Then
    
            ChC = ChM + CoorNEtoYXL(EM, NM, AzBM, Ec, Nc, "Y") 'Chainage of tunnel center
            OsC = CoorNEtoYXL(EM, NM, AzBM, Ec, Nc, "X") 'Horizontal deviation of tunnel center
            VtC = DeviateVt(ChM, ZM, PitchBM, ChC, Zc) 'Vertical deviation of tunnel center
            
            Ed = CoorYXtoNE(EM, NM, AzBM, ChC - ChM, 0, "E") 'Design Easting
            Nd = CoorYXtoNE(EM, NM, AzBM, ChC - ChM, 0, "N") 'Design Northing
            ZD = ZM + PitchBM * (ChC - ChM) 'Design Elevation
            
        Else
    
            ChC = ChM + CoorNEtoYXL(EM, NM, AzMH, Ec, Nc, "Y") 'Chainage of tunnel center
            OsC = CoorNEtoYXL(EM, NM, AzMH, Ec, Nc, "X") 'Horizontal deviation of tunnel center
            VtC = DeviateVt(ChM, ZM, PitchMH, ChC, Zc) 'Vertical deviation of tunnel center
    
            Ed = CoorYXtoNE(EM, NM, AzMH, ChC - ChM, 0, "E") 'Design Easting
            Nd = CoorYXtoNE(EM, NM, AzMH, ChC - ChM, 0, "N") 'Design Northing
            ZD = ZM + PitchMH * (ChC - ChM) 'Design Elevation
    
        End If
        'Debug.Print "R" & sumRng, Ec, Nc, Zc, ChC, OsC, VtC
        
        'Print Result (Best Fit Circle)
        Sheets(WriggleName).Select
        Range("B7").Select

        Dim WRSValue() As Variant
        Dim WRSFormat() As Variant
        WRSValue = Array("R" & sumRng, Ec, Nc, Zc, ChC, OsC * Excavate_Direc, VtC, Radius + avgOSi, (Radius + avgOSi) * 2)
        WRSFormat = Array("@", "0.000", "0.000", "0.000", "0+000.000", "0.000", "0.000", "0.000", "0.000")
        For t = LBound(WRSValue) To UBound(WRSValue)
            ActiveCell.Offset(j, t).Value = WRSValue(t)
            ActiveCell.Offset(j, t).NumberFormat = WRSFormat(t)
        Next
        ActiveCell.Offset(j, -1).Value = i & "," & j
        
        'Print Result(Backup)
        Sheets("Wriggle Backup").Select
        Range("A4").Select
        
        'INDEX
        ActiveCell.Offset(w, 0).Value = w
        'RING NO.
        ActiveCell.Offset(w, 1).Value = "R" & sumRng
        'TUNNEL CENTER
        ActiveCell.Offset(w, 2).Value = Ec
        ActiveCell.Offset(w, 3).Value = Nc
        ActiveCell.Offset(w, 4).Value = Zc
        'CHAINAGE
        ActiveCell.Offset(w, 5).Value = ChC
        'DEVIATION
        ActiveCell.Offset(w, 6).Value = OsC * Excavate_Direc
        ActiveCell.Offset(w, 7).Value = VtC
        'AVERAGE
        ActiveCell.Offset(w, 8).Value = Radius + avgOSi
        ActiveCell.Offset(w, 9).Value = (Radius + avgOSi) * 2
        'COORDINATE
        c = 0
        For t = 1 To numPnt
            ActiveCell.Offset(w, 10 + c).Value = extEi(t)
            ActiveCell.Offset(w, 11 + c).Value = extNi(t)
            ActiveCell.Offset(w, 12 + c).Value = extZi(t)
            c = c + 3
        Next
        'DESIGN CENTER
        ActiveCell.Offset(w, 58).Value = Ed
        ActiveCell.Offset(w, 59).Value = Nd
        ActiveCell.Offset(w, 60).Value = ZD
        ActiveCell.Offset(w, 61).Value = DiaDesign / 2
        ActiveCell.Offset(w, 62).Value = DiaDesign
        'LOCAL COORDINATE
        ActiveCell.Offset(w, 63).Value = Xc
        ActiveCell.Offset(w, 64).Value = Yc
        c = 0
        For t = 1 To numPnt
            ActiveCell.Offset(w, 65 + c).Value = extXi(t)
            ActiveCell.Offset(w, 66 + c).Value = extYi(t)
            c = c + 2
        Next
        'RADIUS
        c = 0
        For t = 1 To numPnt
            ActiveCell.Offset(w, 97 + c).Value = Ri(t) + OSi(t)
            c = c + 1
        Next
        'RDBC
        c = 0
        For t = 1 To numPnt
            ActiveCell.Offset(w, 113 + c).Value = RDBCi(t)
            c = c + 1
        Next
        'ANGLE
        c = 0
        For t = 1 To numPnt
            ActiveCell.Offset(w, 129 + c).Value = ANGi(t)
            c = c + 1
        Next
        'P
        ActiveCell.Offset(w, 145).Value = EP
        ActiveCell.Offset(w, 146).Value = NP
        'Q
        ActiveCell.Offset(w, 147).Value = EQ
        ActiveCell.Offset(w, 148).Value = NQ
        
        'OFFSET
        ActiveCell.Offset(w, 149).Value = avgOSi
        
        'NUM. POINT
        ActiveCell.Offset(w, 150).Value = numPnt
        
        j = j + 1
        w = w + 1
        u = u + numPnt
    Next
    
    Sheets("Wriggle Backup").Select
    Range("A4").Select
    ActiveWindow.Zoom = 100
    
    Sheets(WriggleName).Select
    Range("B7").Select
    ActiveWindow.Zoom = 90
    
    Sheets("Import Wriggle Data").Select
    Range("A4").Select
    MsgBox "Wriggle Computation Complete!"
  
End Sub


Sub SaveToPDFs_WR1()

    Dim F As Long 'Start number
    Dim E As Long 'End number
    Dim Filename As String
    Dim FilePath As String

    F = Range("O4").Value
    E = Range("O5").Value
    
    For i = F To E
  
        Range("O4").Value = i
        Filename = "/" & "Wriggle Report" & " " & Range("D3").Value
        FilePath = ActiveWorkbook.Path
        
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=FilePath & Filename, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
     Next
End Sub

Sub SaveToPDFs_WR2()

    Dim F As Long 'Start number
    Dim E As Long 'End number
    Dim Filename As String
    Dim FilePath As String

    F = Range("W7").Value
    E = Range("W8").Value
    
    For i = F To E Step 5
  
        Range("W7").Value = i
        Filename = "/" & "Wriggle Report" & " " & " Page " & Range("P7").Value & " of " & Range("R7").Value
        FilePath = ActiveWorkbook.Path
        
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=FilePath & Filename, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
     Next
End Sub

Sub ClearData1()
    
    Range("C1:C2").Select
    Selection.ClearContents
    
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Range("A4").Select
    
End Sub

Sub ClearData2()
    
    Range("C1:C2").Select
    Selection.ClearContents
    
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Range("A4").Select
    
End Sub
