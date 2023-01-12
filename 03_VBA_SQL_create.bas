Option Compare Database
Option Explicit

'____________________________________________________________________________________
Sub create_record()

Dim sqlStatement$           ' String variable, ending "$"
Dim strTarget$, strValues$  ' Must match with the columns of the table

strTarget$ = "INTO tbl_parts(part_number, drawing_number, part_description, material, standard)"
strValues$ = "VALUES ('A3482K134', '45JK341B2', 'HEX Bolt M12', 'Stainless Steel', 'DIN 933')

' sqlStatement = "INSERT INTO tbl_parts(...) VALUES (...)
sqlStatement$ = "INSERT " & strTarget$ & " " & strValues

DoCmd.RunSQL sqlStatement$

End Sub
