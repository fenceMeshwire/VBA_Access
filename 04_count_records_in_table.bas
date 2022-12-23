Option Compare Database
Option Explicit

' Activate the following modules:
' Microsoft ActiveX Data Objects 6.1 Library
' Microsoft ActiveX Data Objects Recordset 6.0 Library
' Create a table with the corresponding columns.

'_____________________________________________________________________________________________________________

Sub count_records_in_table()

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strTable$

strTable$ = "tbl_parts"

Set cn = CurrentProject.Connection
Set rs = New ADODB.Recordset

With rs
  .Open strTable$, cn, adOpenKeyset, adLockOptimistic
  Debug.Print .RecordCount & " records found in the table: " & strTable$
End With

Set rs = Nothing
Set cn = Nothing

End Sub
