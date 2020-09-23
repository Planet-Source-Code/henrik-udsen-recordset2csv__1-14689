<div align="center">

## Recordset2CSV


</div>

### Description

Create a CSV file (ie. for Excel(???)) based on a parsed ADO-Recordset - very cute :)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Henrik Udsen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/henrik-udsen.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/henrik-udsen-recordset2csv__1-14689/archive/master.zip)





### Source Code

```
Sub dB_RsToCSVFile(Rs As ADODB.Recordset, FileName As String, Optional Delimiter As String = ",")
 Dim fh As Integer
 Dim FileIsOpen As Boolean, s As Variant
 Dim t As Integer
 Dim Buf As String, TempStr As String
 FileIsOpen = False
 On Error GoTo Err_Out
 fh = FreeFile()
 Open FileName For Output As fh
 FileIsOpen = True
 Buf = ""
 For t = 0 To Rs.Fields.Count - 1
  If Buf = "" Then
   Buf = """" & Rs.Fields(t).Name & """"
  Else
   Buf = Buf & Delimiter & """" & Rs.Fields(t).Name & """"
  End If
 Next t
 Print #fh, Buf
 Do While Not Rs.EOF
  Buf = ""
  For t = 0 To Rs.Fields.Count - 1
   If IsNull(Rs.Fields(t).Value) Then
    TempStr = ""
   Else
    TempStr = Rs.Fields(t).Value
   End If
   If Buf = "" Then
    Buf = """" & TempStr & """"
   Else
    Buf = Buf & Delimiter & """" & TempStr & """"
   End If
  Next t
  Print #fh, Buf
  Rs.MoveNext
 Loop
 Close fh
 Exit Sub
Err_Out:
 If FileIsOpen Then
  Close fh
 End If
 MsgBox "There was an error: " & Error, vbOKOnly, "The file was not created"
End Sub
```

