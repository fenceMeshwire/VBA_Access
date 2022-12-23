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
