Sub ununion()
    Dim rCell As Range, sValue$, sAddress$, i&
    Application.ScreenUpdating = False
      Set iSource = [A1:L600]
        For Each rCell In iSource   
    If rCell.MergeCells Then
        sAddress = rCell.MergeArea.Address: rCell.UnMerge
        Range(sAddress).Value = rCell.Value
    End If
    Next
    Application.ScreenUpdating = True
End Sub
Sub m()
Dim name As String
name = "result"
Dim oSheet As Excel.Worksheet
Set oSheet = Worksheets.Add() ' Nicaaai iiaue eeno
oSheet.name = name 'I?enaaeaaai aio eiy "Iiaue eeno"
    For i = 1 To Sheets.Count
        If Sheets(i).name <> name Then
           myR_Total = Sheets(name).Range("A" & Sheets(name).Rows.Count).End(xlUp).Row
           myR_i = Sheets(i).Range("A" & Sheets(i).Rows.Count).End(xlUp).Row
           Sheets(i).Rows("1:" & myR_i).Copy Destination:=Sheets(name).Range("A" & myR_Total + 1)
        End If
    Next
 Set ws = ActiveWorkbook.Sheets(name)
 ws.Activate
    ununion
'oaaeeou ionoua no?iee
    Dim r As Long, rng As Range
    For r = 1 To ActiveSheet.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.Count
        If Application.CountA(Rows(r)) = 0 Then
            If rng Is Nothing Then Set rng = Rows(r) Else Set rng = Union(rng, Rows(r))
        End If
    Next r
    If Not rng Is Nothing Then rng.Delete
'
ActiveSheet.Copy
Kill(rozklad.csv)
ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & "rozklad.csv", xlCSV, CreateBackup:=False, Local:=True
ActiveWorkbook.Close 0
End Sub