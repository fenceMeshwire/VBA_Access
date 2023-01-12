Option Compare Database
Option Explicit

'____________________________________________________________________________________
Sub delete_records()

Dim strStatement$ ' String variable, ending "$"

strStatement$ = "DELETE * FROM tbl_parts WHERE (name = 'washer')"

DoCmd.RunSQL strStatement$

End Sub
