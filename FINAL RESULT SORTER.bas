Attribute VB_Name = "Module1"
Sub P3()

Dim FugLA(1 To 58800) As Double       ' Array to store first column values
Dim FugUA(1 To 58800) As Double       ' Array to store second column values
Dim FugSoil(1 To 58800) As Double     ' Array to store third column values
Dim FugFO(1 To 58800) As Double       ' Array to store first column values
Dim FugStem(1 To 58800) As Double       ' Array to store second column values
Dim FugRoots(1 To 58800) As Double     ' Array to store third column values
Dim FugCellLA(1 To 10, 1 To 10) As Double
Dim FugCellUA(1 To 10, 1 To 10) As Double
Dim FugCellSoil(1 To 10, 1 To 10) As Double
Dim FugCellFO(1 To 10, 1 To 10) As Double
Dim FugCellStem(1 To 10, 1 To 10) As Double
Dim FugCellRoots(1 To 10, 1 To 10) As Double
Dim i As Integer      ' the value of 4th column
Dim j As Integer      ' the value of 5th column
Dim k As Integer
Dim ws As Worksheet
Dim p As Integer

p = 0

For t = 0 To 1100 Step 100 'SimMonths, original one 11900

   Worksheets("Sheet1").Select

  For k = 1 + t To 100 + t

    FugLA(k) = Cells(k, 1).Value
    FugUA(k) = Cells(k, 2).Value
    FugSoil(k) = Cells(k, 3).Value
    FugFO(k) = Cells(k, 8).Value
    'FugStem(k) = Cells(k, 5).Value
    'FugRoots(k) = Cells(k, 6).Value
    i = Cells(k, 9).Value
    j = Cells(k, 10).Value
    FugCellLA(i, j) = FugLA(k)            '2D arrays arranged the values
    FugCellUA(i, j) = FugUA(k)
    FugCellSoil(i, j) = FugSoil(k)
    FugCellFO(i, j) = FugFO(k)            '2D arrays arranged the values
    'FugCellStem(i, j) = FugStem(k)
    'FugCellRoots(i, j) = FugRoots(k)
    
  Next k

 
    Worksheets("ConcLA").Select

  For i = 1 To 10
  For j = 1 To 10

    Cells(i + p, j).Value = FugCellLA(i, j)
    Cells(p + 1, 11).Value = "Month: " & (t / 100) + 1

  Next j
  Next i
colourscales
     Worksheets("ConcUA").Select

  For i = 1 To 10
  For j = 1 To 10

    Cells(i + p, j).Value = FugCellUA(i, j)
    Cells(p + 1, 11).Value = "Month: " & (t / 100) + 1
    
  Next j
  Next i
colourscales
     Worksheets("ConcSoil").Select

  For i = 1 To 10
  For j = 1 To 10

    Cells(i + p, j).Value = FugCellSoil(i, j)
    Cells(p + 1, 11).Value = "Month: " & (t / 100) + 1
    
  Next j
  Next i

colourscales
    Worksheets("ConcFO").Select

  For i = 1 To 10
  For j = 1 To 10

    Cells(i + p, j).Value = FugCellFO(i, j)
    Cells(p + 1, 11).Value = "Month: " & (t / 100) + 1

  Next j
  Next i
colourscales

 '   Worksheets("ConcStem").Select

  'For i = 1 To 10
  'For j = 1 To 10

   ' Cells(i + p, j).Value = FugCellStem(i, j)
   ' Cells(p + 1, 11).Value = "Month: " & (t / 100) + 1

  'Next j
  'Next i
'colourscales

 '   Worksheets("ConcRoots").Select

  'For i = 1 To 10
  'For j = 1 To 10

   ' Cells(i + p, j).Value = FugCellRoots(i, j)
   ' Cells(p + 1, 11).Value = "Month: " & (t / 100) + 1

  'Next j
  'Next i
'colourscales

 p = p + 10
 
Next t


End Sub

Sub colourscales()

Dim avg As Double
Dim rg As Range
Dim cs As ColorScale
avg = Application.WorksheetFunction.Average(Range("A1:T20"))

Set rg = Range("A1: T20", Range("A1: T20").End(xlDown))
rg.FormatConditions.Delete
'colour scale will have three colours
Set cs = rg.FormatConditions.AddColorScale(ColorScaleType:=3)
With cs
    'the first colour is blue
    With .ColorScaleCriteria(1)
        .FormatColor.Color = RGB(102, 153, 255)
        .Type = xlConditionValueLowestValue
    End With
    'the second colour is yellow
    With .ColorScaleCriteria(2)
        .FormatColor.Color = RGB(255, 230, 153)
        .Type = xlConditionValueNumber
        .Value = avg
    End With
    'the third colour is red
    With .ColorScaleCriteria(3)
        .FormatColor.Color = RGB(255, 51, 0)
        .Type = xlConditionValueHighestValue
    End With

End With

End Sub


