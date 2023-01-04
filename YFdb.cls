VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "YFdb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private dbPath As String

Private Sub Class_Initialize()
    dbPath = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
                    "Data Source=" & ThisWorkbook.Path & "\YFdevTest.accdb;"
                    
End Sub




Public Function insertTracking(njiis_num As String, med_lic As String, _
                            current_stamp As String, tracking_num As String, _
                            stamp_date As Date, notes As String, status As String, _
                            Optional ByVal date_of_entry As Date, Optional ByVal approval_sent_date As Date, Optional ByVal approval_sent As Boolean)
    Dim conn As Object
    
    Set conn = CreateObject("ADODB.Connection")
    Dim qb As New c_QueryDesigner_Insert
    
    Dim approval_date_str As String
    Dim approval_status As String

    
    qb.TableName = "Tracking"
    
    qb.InsertColumn "njiis_num", Chr(34) & njiis_num & Chr(34)
    qb.InsertColumn "med_lic", med_lic
    qb.InsertColumn "stamp_number", current_stamp
    qb.InsertColumn "stamp_order_date", Format(stamp_date, "\#yyyy-mm-dd hh:nn:ss\#")
    qb.InsertColumn "tracking_number", tracking_num
    qb.InsertColumn "notes", notes
    qb.InsertColumn "status", status
    
    If IsNull(date_of_entry) Then
        date_of_entry = Date
    End If
    
    If IsNull(approval_sent_date) Then
        approval_date_str = "Null"
    Else
        approval_date_str = Format(approval_sent_date, "\#yyyy-mm-dd hh:nn:ss\#")
    End If
    
    If approval_sent Then
        approval_status = CStr(-1)
    Else
         approval_status = CStr(0)
    End If
    
    qb.InsertColumn "date_of_entry", Format(date_of_entry, "\#yyyy-mm-dd hh:nn:ss\#")
    
    qb.InsertColumn "approval_letter_sent_date", approval_date_str
    qb.InsertColumn "approval_letter_sent", approval_status
    
    With conn
        .Open dbPath
            .Execute qb.GetInsertCommandText
        .Close
    End With
    
    Set conn = Nothing
    
End Function
Public Function getTracking(status As String)
    Dim conn As Object, rs As Object
    Dim qrystr As String
    Set conn = CreateObject("ADODB.Connection")
    
    Dim qb As New c_QueryDesigner_Select
    Set rs = CreateObject("ADODB.Recordset")
    
    qb.fromTable = "Tracking"
    qb.SelectCol "njiis_num", "med_lic", "status", "stamp_number", "tracking_number", "notes", "stamp_order_date"
    qb.Where = "status = " & Chr(34) & status & Chr(34)
    
    With conn
        .Open dbPath
            qrystr = qb.GetQuery()
            Debug.Print qrystr
            Set rs = .Execute(qrystr)
            If Not rs.EOF Or Not rs.BOF Then
                    rs.MoveFirst
                    getTracking = rs.GetRows
            End If
        rs.Close
        .Close
    End With
    

    Set rs = Nothing
    Set conn = Nothing
    
     
End Function

Public Function getFacilities(med_lic As String) As Variant
    Dim conn As Object, rs As Object
    Dim qrystr As String
    
    
    Set conn = CreateObject("ADODB.Connection")
    
    Set rs = CreateObject("ADODB.Recordset")
    
    conn.Open dbPath
    Dim qb As New c_QueryDesigner_Select
    Dim last_row As Long, index As Long
    Dim colnames() As Variant
    
    With stampHolderCols
        last_row = .Cells(.Rows.Count, 2).End(xlUp).Row
        colnames = .Range("B1:B" & last_row)
        qb.fromTable = "Appren"
        qb.SelectByArr colnames
        qb.Where = "[Stamp Holder - Med Lic]=" & Chr(34) & med_lic & Chr(34)
        qrystr = qb.GetQuery
        Debug.Print qrystr
    End With
    
    Set rs = conn.Execute(qrystr)
    
    If Not rs.EOF Or Not rs.BOF Then
        rs.MoveFirst
        getFacilities = rs.GetRows
    End If
    
    
    rs.Close
    conn.Close

    Set rs = Nothing
    Set conn = Nothing

End Function


Public Function getStampHolder() As Variant
    Dim conn As Object, rs As Object
    Dim qrystr As String
    
    Set conn = CreateObject("ADODB.Connection")
    
    Set rs = CreateObject("ADODB.Recordset")
    
    conn.Open dbPath
    Dim qb As New c_QueryDesigner_Select
    Dim last_row As Long, index As Long
    Dim colnames() As Variant
    
    With stampHolderCols
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
        colnames = .Range("A1:A" & last_row)
        qb.fromTable = "StampHolder"
        qb.SelectByArr colnames
        qb.OrderBy "[Stamp Holder - Last Name]"
        qrystr = qb.GetQuery
    End With
    
    Set rs = conn.Execute(qrystr)
    
    If Not rs.EOF Or Not rs.BOF Then
        rs.MoveFirst
        getStampHolder = rs.GetRows
    End If
    
    
    rs.Close
    conn.Close

    Set rs = Nothing
    Set conn = Nothing


    
    
    
End Function

Public Function getMasterTrackingList() As Variant
    Dim conn As Object, rs As Object
    Dim qrystr As String
    
    Set conn = CreateObject("ADODB.Connection")
    
    Set rs = CreateObject("ADODB.Recordset")
    
    conn.Open dbPath
    Dim qb As New c_QueryDesigner_Select
    Dim last_row As Long, index As Long
    Dim colnames() As Variant
    
    qb.fromTable = "master_tracking_list"
    qb.SelectCol "*"
    qb.OrderBy "[Stamp Holder - Last Name]", "[Approval Letter Date]"
    qrystr = qb.GetQuery
    Debug.Print qrystr
    
    Set rs = conn.Execute(qrystr)
    
    If Not rs.EOF Or Not rs.BOF Then
        rs.MoveFirst
        getMasterTrackingList = rs.GetRows
    End If
    
    
    rs.Close
    conn.Close

    Set rs = Nothing
    Set conn = Nothing


    
    
    
End Function
