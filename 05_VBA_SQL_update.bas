Option Compare Database
Option Explicit

'____________________________________________________________________________________
Sub update_records()

Dim sqlStatement$ ' String variable, ending "$"

sqlStatement$ = "UPDATE tbl_parts SET tbl_parts.partname = 'special washer' WHERE tbl_parts.partname = 'washer')"

DoCmd.RunSQL sqlStatement$
  
End Sub
