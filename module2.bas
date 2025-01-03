Attribute VB_Name = "Module1"
Sub stocks()
Dim tick As String
Dim openp As Double
Dim closep As Double
Dim qchange As Double
Dim pchange As Double
Dim totalvol As Double
Dim totalvolbig As Double
Dim totalvoltick As String
Dim totalvolnext As Double
Dim currvol As Double
Dim lasttick As String
Dim current As Worksheet
Dim lastrow As Long
Dim big As Double
Dim bignext As Double
Dim bigtick As String
Dim small As Double
Dim smalltick As String
Dim i As Integer
Dim n As Long
Dim j As Long


wscount = ActiveWorkbook.Worksheets.Count
totalvol = 0

For Each current In Worksheets
    big = 0
    small = 0
    lastrow = current.Range("A" & Rows.Count).End(xlUp).Row
    current.Cells(1, "I") = "Ticker"
    current.Cells(1, "J") = "Quarterly Change"
    current.Cells(1, "K") = "Percentage Change"
    current.Cells(1, "L") = "Total Volume"
    lasttick = " "
    n = 2
    j = 1
    totalvol = 0
    big = current.Cells(j + 1, "J").Value
    bignext = 0
    small = 0
    smallnext = 0
    bigtick = " "
    smalltick = " "
    For n = 2 To lastrow
        tick = current.Cells(n, 1)
        If tick = lasttick Then
            closep = current.Cells(n, 6).Value
            currvol = current.Cells(n, 7).Value
            totalvol = currvol + totalvol
            current.Cells(j, "L").Value = totalvol
            current.Cells(j, "J").Value = closep - openp
            current.Cells(j, "K").Value = (closep - openp) / openp
        Else
            j = j + 1
            totalvol = 0
            openp = current.Cells(n, 3).Value
            current.Cells(j, "I").Value = tick
            lasttick = tick
            bignext = current.Cells(j + 1, "K").Value
            totalvolbig = current.Cells(j, "L").Value
            totalvolnext = current.Cells(j + 1, "L").Value
            If bignext > big Then
                big = bignext
                bigtick = current.Cells(j + 1, "I").Value
            End If
            If bignext < small Then
                small = smallnext
                smalltick = current.Cells(j + 1, "I").Value
            End If
            If totalvolnext > totalvol Then
                totalvolbig = totalvol
                totalvoltick = current.Cells(j + 1, "I").Value
            End If
        End If
        
    Next n
    current.Cells(2, "N") = "Greatest % Increase"
    current.Cells(3, "N") = "Greatest % Decrease"
    current.Cells(4, "N") = "Greatest Total Volume"
    current.Cells(1, "O") = "Ticker"
    current.Cells(1, "P") = "Value"
    current.Cells(2, "P") = big
    current.Cells(2, "O") = bigtick
    current.Cells(3, "P") = small
    current.Cells(3, "O") = smalltick
    current.Cells(4, "P") = totalvolbig
    current.Cells(4, "O") = totalvoltick

    
    
Next

End Sub

