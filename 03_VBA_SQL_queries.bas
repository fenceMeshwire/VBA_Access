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
MsgBox "Unable to find the table / wrong SQL query."

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
MsgBox "Unable to find the table / wrong SQL query."

End Sub
    
'____________________________________________________________________________________
Sub limited_query_order_ascending()

Dim rs As New ADODB.Recordset

On Error GoTo err_msg

rs.Open "SELECT part_number, part_description FROM tbl_parts ORDER BY part_description", CurrentProject.Connection

Debug.Print rs.GetString
' The the values of the corresponding columns of the record is displayed in the console.

rs.Close

Exit Sub

err_msg:
MsgBox "Unable to find the table / wrong SQL query."

End Sub
