Option Compare Database
Option Explicit

'____________________________________________________________________________________
Sub delete_records()

Dim sqlStatement$ ' String variable, ending "$"

sqlStatement$ = "DELETE * FROM tbl_parts WHERE (name = 'washer')"

DoCmd.RunSQL sqlStatement$

End Sub
