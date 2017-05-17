Sub Mortgage_Rates()


HousePrice = InputBox("Enter Mortgage Amount Here")
Rate_First_Five = InputBox("Enter the rate for years 1-5 here (in % terms)")
Rate_Six_To_Ten = InputBox("Enter the rate for years 6-10 here (in % terms)")
Rate_Eleven_To_Fifteen = InputBox("Enter the rate for years 10-15 here (in % terms)")
Rate_Sixteen_To_Twenty = InputBox("Enter the rate for years 15-20 here (in % terms)")
Rate_Twentyone_To_Twentyfive = InputBox("Enter the rate for years 21-25 here (in % terms)")
Rate_Twentyfive_To_Thirty = InputBox("Enter the rate for years 25-30 here (in % terms)")

Range("$I$5").value = HousePrice
Range("$C$5:$C$64").value = Rate_First_Five / 100
Range("$C$65:$C$124").value = Rate_Six_To_Ten / 100
Range("$C$125:$C$184").value = Rate_Eleven_To_Fifteen / 100
Range("$C$185:$C$244").value = Rate_Sixteen_To_Twenty / 100
Range("$C$245:$C$304").value = Rate_Twentyone_To_Twentyfive / 100
Range("$C$245:$C$364").value = Rate_Twentyfive_To_Thirty / 100

solverreset

solverok setcell:="$F$364", maxminval:=3, valueof:=0, bychange:="$I$6"

solversolve userfinish:=True

End Sub

Sub Coumpounding()

Dim i As Integer

For i = 15 To 24
Range("D" & i).Offset(rowoffset:=0, columnoffset:=1).Activate


End Sub
