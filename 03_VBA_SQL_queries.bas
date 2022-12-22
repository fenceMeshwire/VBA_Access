Option Compare Database
Option Explicit

'____________________________________________________________________________________
Sub general_query()

Dim rs As New ADODB.Recordset

On Error GoTo err_msg

rs.Open "SELECT * FROM tbl_parts", CurrentProject.Connection

Debug.Print rs.GetString
' The complete record is displayed in the console.

rs.Close

Exit Sub

err_msg:
MsgBox "Unable to find the table."

End Sub

'____________________________________________________________________________________
Sub limited_query()

Dim rs As New ADODB.Recordset

On Error GoTo err_msg

rs.Open "SELECT part_number FROM tbl_parts", CurrentProject.Connection

Debug.Print rs.GetString
' The the value for the corresponding column of the record is displayed in the console.

rs.Close

Exit Sub

err_msg:
MsgBox "Unable to find the table."

End Sub
