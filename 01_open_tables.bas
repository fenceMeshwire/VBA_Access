Option Compare Database
Option Explicit

'_____________________________________________________________________________________________________________
Sub open_table()

On Error GoTo err_msg
DoCmd.OpenTable "tbl_parts", acViewNormal
'  Alternative table views:
'  DoCmd.OpenTable "tbl_part_to_vehicle_reference", acViewDesign
'  DoCmd.OpenTable "tbl_part_to_vehicle_reference", acViewPivotChart
'  DoCmd.OpenTable "tbl_part_to_vehicle_reference", acViewPivotTable
'  DoCmd.OpenTable "tbl_part_to_vehicle_reference", acViewPreview
  Exit Sub
  
err_msg:
  MsgBox "The table could not be found"
  
End Sub

'_____________________________________________________________________________________________________________
Sub open_table_find_record_field()

On Error GoTo err_msg

DoCmd.OpenTable "tbl_parts", acViewNormal
DoCmd.FindRecord "Tires", acEntire, True, acSearchAll, True, acAll
     'FindRecord(SearchFor, Compare, UpperLowerCase, Search, Formated, ActualField, StartFromBeginning)
Exit Sub

err_msg:
  MsgBox "The table could not be found"
End Sub

'_____________________________________________________________________________________________________________
Sub open_table_mark_row()

Dim row As Integer
row = 5

On Error GoTo err_msg
DoCmd.OpenTable "tbl_parts", acViewNormal
DoCmd.GoToRecord acDataTable, "tbl_part_to_vehicle_reference", acGoTo, row
     ' GoToRecord(ObjectType, ObjectName, Record, Offset)
Exit Sub

err_msg:
  MsgBox "The table could not be found"

End Sub
