VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_QueryDesigner_Insert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private p_Columns As Collection
Private p_Values As Collection
Private p_TableName As String


''''''''''''
'CONSTRUCTOR
''''''''''''
Private Sub Class_Initialize()
    Set p_Columns = New Collection
    Set p_Values = New Collection
End Sub



'''''''''''
'PROPERTIES
'''''''''''

'Value: The name of the database table to insert the values into.
Public Property Let TableName(value As String)
    p_TableName = value
End Property



'PARAM (columnName): THE NAME OF THE COLUMN TO BE ADDED TO THE QUERY
'PARAM (columnValue): THE MATCHING VALUE TO BE INSERTED
Public Sub InsertColumn(columnName As String, columnValue As String)

        If IsNumeric(columnValue) Then
            p_Columns.Add "[" & columnName & "]"
            If Len(columnValue) >= 5 Then
                p_Values.Add Chr(34) & columnValue & Chr(34)
            Else
                p_Values.Add columnValue
            End If
            
        ElseIf InStr(columnValue, "#") <> 0 Then
            p_Columns.Add "[" & columnName & "]"
            p_Values.Add columnValue
        ElseIf InStr(columnValue, Chr(34)) <> 0 Then
            p_Columns.Add "[" & columnName & "]"
            p_Values.Add columnValue
        ElseIf InStr(columnValue, "Null") <> 0 Then
            p_Columns.Add "[" & columnName & "]"
            p_Values.Add columnValue
        Else
            p_Columns.Add "[" & columnName & "]"
            p_Values.Add Chr(34) & columnValue & Chr(34)
        
        End If
   
End Sub


'''''''''''''''
'INSERT COMMAND
'''''''''''''''
'RETURNS THE FORMATTED QUERY WITH ANY OPTIONAL CLAUSES ATTACHED
Public Function GetInsertCommandText() As String
    Dim insert As String
    insert = "INSERT INTO " & p_TableName & "(" & Join(CollectionToArray(p_Columns), ", ") & ")"
    insert = insert & " VALUES(" & Join(CollectionToArray(p_Values), ", ") & ");"
    Debug.Print insert
    GetInsertCommandText = CStr(insert)
End Function


'''''''''''''''''
'HELPER FUNCTIONS
'''''''''''''''''
'CONVERTS A COLLECTION INTO AN ARRAY
Private Function CollectionToArray(myCol As Collection) As Variant
 
    Dim result  As Variant
    Dim cnt     As Long
 
    ReDim result(myCol.Count - 1)
 
    For cnt = 0 To myCol.Count - 1
        result(cnt) = myCol(cnt + 1)
    Next cnt
 
    CollectionToArray = result
    
End Function

