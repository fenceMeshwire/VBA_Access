Option Compare Database
Option Explicit

' Activate the following modules:
' Microsoft ADO Ext 6.0 for DLL and Security

'____________________________________________________________________________________
Sub list_tables()

Dim catalogue As ADOX.Catalog
Dim tableInfo As ADOX.Table

Set catalogue = New ADOX.Catalog

catalogue.ActiveConnection = CurrentProject.Connection

For Each tableInfo In catalogue.Tables

  With tableInfo
    
    If tableInfo.Type = "TABLE" Then
      Debug.Print "Name of the table: " & .Name
      Debug.Print "Date of creation: " & .DateCreated
      Debug.Print "Date of last change: " & .DateModified
      Debug.Print vbLf
    End If
      
  End With

Next tableInfo

Set catalogue = Nothing

End Sub
  
'____________________________________________________________________________________
Sub list_table_schema_information()

Dim catalogue As New ADOX.catalog
Dim table As ADOX.table
Dim intCounter As Integer

catalogue.ActiveConnection = CurrentProject.Connection

Set table = catalogue.Tables("tbl_parts")

With table
  For intCounter = 0 To .Columns.Count - 1
    Debug.Print .Columns(intCounter).Name
    Debug.Print .Columns(intCounter).Properties("Description")
    Debug.Print .Columns(intCounter).DefinedSize
    Debug.Print .Columns(intCounter).Type
    Debug.Print .Columns(intCounter).NumericScale
    Debug.Print .Columns(intCounter).Precision
    Debug.Print .Columns(intCounter).Attributes & vbLf
  Next intCounter
End With

Set catalogue = Nothing

End Sub
