Function Slope(coords As Variant) As Double()
   Dim X1, X2, Y1, Y2, N As Integer
   Dim SlopeList() As Double
   coords(1, 1) = X1
   coords(2, 1) = X2
   coords(1, 2) = Y1
   coords(2, 2) = Y2

   N = 0
   While N <= (UBound(coords, 1) - 1)
      X1 = coords(1, 2 * (N + 1) - 1)
      X2 = coords(1, 2 * (N + 1))
      Y1 = coords(2, 2 * (N + 1) - 1)
      Y2 = coords(2, 2 * (N + 1))

      'BUG: "Slope" is the return value of the function which is an array of Double, not a number
      ' Just femoved this line & assigned to SlopeList(k)
      'Slope = (Y2 - Y1) / (X2 - X1)
      '
      ReDim Preserve SlopeList(k)
      SlopeList(k) = (Y2 - Y1) / (X2 - X1)
      k = k + 1
   Wend
   
   Slope = SlopeList()
End Function

' BUG: as with slope you need parens around the end to show you're returning an array
' Function Midpoint(coords()) As Double
Function Midpoint(coords()) As Double()
   Dim MidX, MidY As Double
   Dim X1, X2, Y1, Y2, N As Integer
   Dim MidList() As Double
   N = 0
   While N <= (UBound(coords, 1) - 1)
      X1 = coords(1, 2 * (N + 1) - 1)
      X2 = coords(1, 2 * (N + 1))
      Y1 = coords(2, 2 * (N + 1) - 1)
      Y2 = coords(2, 2 * (N + 1))
      MidX = (X1 + X2) / 2
      MidY = (Y1 + Y2) / 2
      ReDim Preserve MidList(2, N + 1)
      MidList(1, N) = MidX
      MidList(2, N) = MidY
      N = N + 1
   Wend
   Midpoint = MidList()

End Function

' BUG 1: as with slope you need parens around the end to show you're returning an array
' BUG 2: Changed parameter to type Variant - not sure why needed https://stackoverflow.com/a/30941227
'   Function Center(slopecoords()) As Double
Function Center(slopecoords As Variant) As Double()
   Dim X1, X2, Y1, Y2, N As Integer
   Dim Xcircle, Ycircle, M1, M2 As Double
   Dim CenterList() As Double
   N = 0
   While N <= (UBound(slopecoords, 1) - 1)
      X1 = slopecoords(1, 2 * (N + 1) - 1)
      X2 = slopecoords(1, 2 * (N + 1))
      Y1 = slopecoords(2, 2 * (N + 1) - 1)
      Y2 = slopecoords(2, 2 * (N + 1))
      M1 = slopecoords(3, 2 * (N + 1) - 1)
      M2 = slopecoords(3, 2 * (N + 1))

      Xcircle = (M2 * X2 - Y2 - M1 * X1 + Y1) / (M2 - M1)
      Ycircle = (-M1 * X1 + Y1) + M1 * X1

      ReDim Preserve CenterList(N)
      CenterList(1, N) = Xcircle
      CenterList(2, N) = Ycircle
      N = N + 1
   Wend
   Center = CenterList()

End Function

Sub Circles()
   Dim N As Integer
   Dim Row, PerpSlope As Double
   Dim coords() As Variant
   Dim SlopeList() As Double
   Dim PerpSlopeList() As Double
   Dim MidList() As Double
   Dim FinList() As Double
   Dim CenterList() As Double


   Debug.Print ("Starting Circles...")
   
   Row = 0
   While ActiveSheet.Cells(Row + 3, 3) <> ""
      ReDim Preserve coords(2, Row)
      coords(1, Row) = ActiveSheet.Cells(Row + 3, 3).Value
      coords(2, Row) = ActiveSheet.Cells(Row + 3, 4).Value
      Debug.Print ("Circles.coords: " & coords(1, Row) & ", " coords(2, Row))
      
      Row = Row + 1
   Wend

   SlopeList() = Slope(coords())
   N = 0
   While N < UBound(SlopeList, 1)
      PerpSlope = -(1 / SlopeList(k))
      ReDim Preserve PerpSlopeList(k)
      PerpSlopeList(N) = PerpSlope
      k = k + 1
   Wend

   ' BUG: No need to Dim because "Midpoint" is a funcion not a variable.
   ' Dim Midpoint(coords()) As Variant
   '
   ' Just call it when assigning to MidList...
   MidList() = Midpoint(coords())
   
   N = 0
   While N <= UBound(MidList())
      ReDim Preserve FinList(3, N)

      ' BUG: Variable MidList has the output of MidPoint, use it here...
      ' FinList(1, N) = Midpoint(1, N)
      ' FinList(2, N) = Midpoint(2, N)
      ' FinList(3, N) = PerpSlopeList(N)
      FinList(1, N) = MidList(1, N)
      FinList(2, N) = MidList(2, N)
      FinList(3, N) = PerpSlopeList(N)
      N = N + 1
   Wend

   ' BUG: just have this reversed.. variable = function()
   '
   ' Center(FinList()) = CenterList()

   'CenterList() = Center(FinList)
   CenterList() = Center(FinList)

   'Outputs
   ActiveSheet.Cells(3, 7) = CenterList(1, 1)
   ActiveSheet.Cells(4, 7) = CenterList(2, 1)
End Sub
