Option Compare Database
Option Explicit

'____________________________________________________________________________________
Sub update_records()

Dim strStatement$ ' String variable, ending "$"

strStatement$ = "UPDATE tbl_parts SET tbl_parts.partname = 'special washer' WHERE tbl_parts.partname = 'washer')"

DoCmd.RunSQL strStatement$
  
End Sub
