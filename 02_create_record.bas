Option Compare Database
Option Explicit

' Activate the following modules:
' Microsoft ActiveX Data Objects 6.1 Library
' Microsoft ActiveX Data Objects Recordset 6.0 Library
' Create a table with the corresponding columns.

'_____________________________________________________________________________________________________________
Sub create_record()

Dim cn As New ADODB.connection
Dim rs As ADODB.recordset

Set cn = CurrentProject.connection
Set rs = New ADODB.recordset

rs.Open "tbl_parts", cn, adOpenKeyset, adLockOptimistic
rs.AddNew

rs!part_number = "A3482K134"
rs!drawing_number = "45JK341B2"
rs!part_description = "HEX Bolt M12"
rs!material = "Stainless Steel"
rs!standard = "DIN 933"

rs.Update
rs.Close

Set rs = Nothing
Set cn = Nothing

End Sub
