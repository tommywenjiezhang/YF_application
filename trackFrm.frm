VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} trackFrm 
   Caption         =   "Tracking Form"
   ClientHeight    =   26016
   ClientLeft      =   336
   ClientTop       =   2064
   ClientWidth     =   66888.01
   OleObjectBlob   =   "trackFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "trackFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private validate_flag As Boolean
    

Private Sub approval_letter_sent_chk_Click()
    With Me.approval_letter_sent_chk_no
        .value = False
    End With
End Sub

Private Sub approval_letter_sent_chk_no_Click()
    With Me.approval_letter_sent_chk
        .value = False
    End With
End Sub

Private Sub cancelBtn_Click()
     Unload Me
End Sub

Private Sub exportMasterbtn_Click()
    Call exportMasterList.export
End Sub


Private Sub facilityCbo_Change()
    Dim x As Integer
    Dim look_val As String, stamp_result As Variant, comment_result As Variant
    
       look_val = CStr(Me.facilityCbo.value)
       If Len(look_val) > 0 Then
        stamp_result = Application.WorksheetFunction.VLookup(look_val, facilityData.UsedRange, 18, False)
        comment_result = Application.WorksheetFunction.VLookup(look_val, facilityData.UsedRange, 16, False)
        If Not IsError(stamp_result) Then
            Me.stampNumtxt = CStr(stamp_result)
            Me.notes_txt = CStr(comment_result)
            Me.stampOrderDateTxt.value = Format(Date, "mm/dd/yyyy")
        End If
       End If
End Sub

Private Sub providerCbo_Change()
    Dim lookup_val As Variant, lookup_rng As Range
    
    With stampHolderData
        Set lookup_rng = .UsedRange
        lookup_val = Application.WorksheetFunction.VLookup(Me.providerCbo.value, lookup_rng, 6, False)
        If Not IsError(lookup_val) Then
            dataHelper.getFacilityData CStr(lookup_val)
        End If
    End With
    With facilityData
        Dim idx As Long
        Me.facilityCbo.Clear
        For idx = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row
            Me.facilityCbo.AddItem .Cells(idx, 1).value
        Next idx
    End With
End Sub


Private Sub quickSearchTxt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim x As Integer
    Dim look_val As String, stamp_result As Variant, comment_result As Variant
       If KeyCode = 13 And Len(Me.quickSearchTxt.value) > 1 Then
        Dim Item As Variant
        Dim n As Long
        look_val = Me.quickSearchTxt.value
        With Me.providerCbo
            For n = 0 To .ListCount - 1
                Item = .List(n)
                If Not IsNull(Item) And InStr(CStr(Item), Me.quickSearchTxt.value) > 0 Then
                    .value = CStr(Item)
                    Exit Sub
                End If
            
            Next n
            MsgBox ("Not Found")
        End With
    End If
End Sub

Private Sub status_cbo_Change()
    Dim status_str As String
    
    status_str = Me.status_cbo.value
    Select Case status_str
        Case "Temp Stamp Sent"
            Me.orderlbl.Caption = "Temp Sent date"
            Call dataHelper.getTrackingData(status_str)
            populate_status
        Case "Official Stamp Ordered"
            Me.orderlbl.Caption = "Official Stamp Order Date"
            Call dataHelper.getTrackingData(status_str)
            populate_status
        Case "Official Stamp Sent"
            Me.orderlbl.Caption = "Official Stamp Sent Date"
            Call dataHelper.getTrackingData(status_str)
            populate_status
    End Select
End Sub

Private Sub populate_status()
    Dim njiis_id As String, njiis_lookup As Variant
    Dim med_lic_lookup As Variant, med_lic  As String
    Dim order_date_lookup As Variant, order_date As String, look_up_tracking_key As String
    Dim track_num_lookup As Variant
    
    
    
    With stampHolderData
        med_lic_lookup = Application.WorksheetFunction.VLookup(Me.providerCbo.value, .UsedRange, 6)
        If Not IsError(med_lic_lookup) Then
             med_lic = CStr(med_lic_lookup)
        End If
    End With
    
    With facilityData
        look_val = CStr(Me.facilityCbo.value)
        If Not IsError(look_val) Then
            njiis_id = Application.WorksheetFunction.VLookup(CStr(look_val), .UsedRange, 2, False)
        End If
    End With
    
    If Application.WorksheetFunction.CountA(trackDatasht.UsedRange) = 0 Then
        Exit Sub
    Else
        look_up_tracking_key = Trim(njiis_id) & "-" & Trim(med_lic)
        order_date_lookup = Application.VLookup(look_up_tracking_key, trackDatasht.UsedRange, 8, False)
        track_num_lookup = Application.VLookup(look_up_tracking_key, trackDatasht.UsedRange, 6, False)
        If Not IsError(order_date_lookup) And Not IsError(track_num_lookup) Then
            Me.stampOrderDateTxt.value = Format(order_date_lookup, "mm/dd/yyyy")
            Me.trackingNum.value = track_num_lookup
        End If
    
    End If
    

End Sub

Private Sub submitBtn_Click()
    Dim njiis_id As String, njiis_lookup As Variant
    Dim med_lic_lookup As Variant, med_lic  As String
    Dim status_str As String, track_num_str As String, notes_str As String, stamp_order_date As Date, stamp_num As String, date_entry_date As Date, approval_sent_val As Boolean, approval_date  As Date
    Dim db As New YFdb
   
    validate_flag = True
   
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Then
            If ctrl.value = "" And Not ctrl.name = "quickSearchTxt" And Not ctrl.name = "notes_txt" Then
                validate_flag = False
                ctrl.BackColor = vbYellow
                Exit For
            End If
        End If
    Next ctrl
    
    If validate_flag Then
         With stampHolderData
            med_lic_lookup = Application.WorksheetFunction.VLookup(Me.providerCbo.value, .UsedRange, 6)
            If Not IsError(med_lic_lookup) Then
                 med_lic = CStr(med_lic_lookup)
            End If
        End With
        
        With facilityData
            look_val = CStr(Me.facilityCbo.value)
            If Not IsError(look_val) Then
                njiis_id = Application.WorksheetFunction.VLookup(CStr(look_val), .UsedRange, 2, False)
            End If
        End With
    
    
        status_str = Me.status_cbo.value
        track_num_str = Me.trackingNum.value
        notes_str = Me.notes_txt.value
        stamp_order_date = CDate(Me.stampOrderDateTxt.value)
        stamp_num = Me.stampNumtxt.value
        date_of_entry_date = CDate(Me.dateEntryTxt.value)
        approval_date = CDate(Me.approval_letter_sent_date_txt.value)
        
        If Me.approval_letter_sent_chk.value = True Then
            approval_sent_val = True
        End If
        
    
        db.insertTracking njiis_id, med_lic, stamp_num, track_num_str, stamp_order_date, notes_str, status_str, date_of_entry_date, approval_date, approval_sent_val
        MsgBox "Record Inserted"
    Else
        MsgBox "Please check your input"
    End If
    
    
End Sub



Private Sub UserForm_Initialize()

    Dim last_row As Long
    Call dataHelper.getStampHolderData
    
    Me.Height = 500
    Me.Width = 800
    
    With Me.status_cbo
        .AddItem "Temp Stamp Sent"
        .AddItem "Official Stamp Ordered"
        .AddItem "Official Stamp Sent"
        
    End With
    
    
    
    With stampHolderData
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
        Dim idx As Long
        For idx = 2 To last_row
            Me.providerCbo.AddItem .Cells(idx, 1).value
        Next idx
    End With
    
    Me.dateEntryTxt.value = Format(Date, "mm/dd/yyyy")
    
    
    
    
End Sub
