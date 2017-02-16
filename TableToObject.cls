VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableToObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public tableRows As Collection
Public tableRow As Object
Private dataTable As ListObject
    
Private Sub Class_Initialize()



End Sub

Public Property Get Count() As Integer

Count = tableRows.Count

End Property

Public Property Get Table() As ListObject

    If dataTable <> Empty Then

        Set Table = dataTable
        
    End If

End Property
Public Property Set Table(Table As ListObject)

    Set dataTable = Table

End Property


Public Function LoadData(Optional sourceTable As ListObject) As Collection

Dim errors As New errorHandler

Dim listHeaders As Object
Dim cell As Range
Dim row As ListRow
Set tableRows = New Collection
Set listHeaders = CreateObject("Scripting.Dictionary")


    If sourceTable Is Nothing Then
    
        If dataTable Is Nothing Then
        
            errors.AddError "No table loaded and parameter not specified ,", StandardError
            Set LoadData = Nothing
            Exit Function
        
        End If
        
    Else
    
        Set dataTable = sourceTable
    
    End If
    
colIndex& = 1
For Each cell In dataTable.HeaderRowRange

    listHeaders.Add colIndex, cell.Value
    colIndex = colIndex + 1
    
Next cell


colIndex = 1
For Each row In dataTable.ListRows

    Set tableRow = CreateObject("Scripting.Dictionary")
    
    For Each element In row.Range.Value2
    
        tableRow.Add listHeaders(colIndex), element
        colIndex = colIndex + 1
        
    Next element
    tableRows.Add tableRow
    Set tableRow = Nothing
    colIndex = 1
    
Next row
        


End Function


