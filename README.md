# ReadTxtFileVBA

Sub Acme_Sub_Runner()

processOneDataFile ThisWorkbook.path & “\WSI 2008x11.txt”

End sub

Sub processOneDataFile(path as String)

Dim oneLine As String
Dim mo_yr As String
Dim pos As Integer
Dim days As Variant
Dim x as Integer
Dim state as String
Dim city as String
Dim data as Variant
Dim r As Long

r = 2

Open path For Input as #1
  Line Input #1, oneLine
  Line Input #1, oneLine
  pos = instr(1, oneLine, “-“)
  mo_yr = mid(oneLine, pos, 7)


‘ Looking for the raw data
Do Until Right(onLIne, 4) = “NORM”
 Line Input #1, oneLine
 DoEvents
Loop

days = Trim(Mid(onLIne, 20, 40))
Do until instr(days, “  “) = 0
  days = replace(days, “  “, “ “)
Loop
days = split(days) ‘ splits delimiter by default is a space

Do
‘ Move down 3 lines to get state
For x = 1 to 3
  Line Input #1, oneLine
Next

state = mid(oneLine, 6, Len(oneLine)-8)

‘ Move down 2 more lines
 Line Input #1, oneLine

Do
   Line Input #1, oneLine
   If oneLine = “” Then Exit Do

  city = mid(oneLine, 2, 12)

  data = Trim(Mid(onLIne, 20, 40))
  Do until instr(days, “  “) = 0
    days = replace(days, “  “, “ “)
  Loop
  data = split(days) ‘ splits delimiter by default is a space

  For x = 0 to Ubound(data)
   Cells(r,1).Value = state
   Cells(r,2).Value = city
   Cells(r,3).Value = days(x) & mo_yr
   Cells(r,4).Value = data(x)
   r   = r +1
  Next
Loop

‘ Advance to next data set

Do Until Right(onLIne, 4) = “NORM”
 Line Input #1, oneLine
 If EOF(1) Then
  Close #1
  Exit Sub
 DoEvents
Loop

Loop
Close #1
End Sub
