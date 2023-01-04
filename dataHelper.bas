Attribute VB_Name = "dataHelper"
Sub getStampHolderData()
    Dim db As New YFdb
    Dim data As Variant
    Dim result As Variant
    Dim last_col As Long, last_row As Long
    Dim data_wb As Workbook, data_sht As Worksheet, data_rng As Range
    
    
    data = db.getStampHolder()
    
    With stampHolderData
        last_col = UBound(data, 1)
        last_row = UBound(data, 2)
        .Cells.ClearContents
        Set data_wb = Workbooks.Add
        Set data_sht = data_wb.Sheets(1)
        data_sht.Range("B1").Resize(last_col + 1, last_row + 1).value = data
        Set data_rng = data_sht.UsedRange
        .Range(.Cells(1, 2), .Cells(1, 12)).value = _
        Application.WorksheetFunction.Transpose(stampHolderCols.Range("A1:A12"))
        .Range("B2").Resize(last_row + 1, last_col + 1).value = Application.WorksheetFunction.Transpose(data_rng)
        .Range("A1").value = "Key"
        .Range("A2:A" & last_row).FormulaR1C1 = "=CONCAT(RC[1],""-"",RC[5])"
        data_wb.Close savechanges:=False
        
    End With
    
End Sub



Public Sub getFacilityData(med_lic As String)
    Dim db As New YFdb
    Dim data As Variant
    Dim result As Variant
    Dim last_col As Long, last_row As Long, last_sht_row As Long
    Dim data_wb As Workbook, data_sht As Worksheet, data_rng As Range
    Dim col_last_row As Long

    
    
    data = db.getFacilities(med_lic)
    
    With facilityData
        last_col = UBound(data, 1)
        last_row = UBound(data, 2)
        .Range(.Cells(2, 1), .Cells(.Rows.Count, .columns.Count)).ClearContents
        
        Set data_wb = Workbooks.Add
        Set data_sht = data_wb.Sheets(1)
        data_sht.Range("A1").Resize(last_col + 1, last_row + 1).value = data
        Set data_rng = data_sht.UsedRange
        col_last_row = stampHolderCols.Cells(stampHolderCols.Rows.Count, 1).End(xlUp).Row
        .Range(.Cells(1, 2), .Cells(1, 1 + col_last_row)).value = _
        Application.WorksheetFunction.Transpose(stampHolderCols.Range("B1:B" & col_last_row))
        .Range("B2").Resize(last_row + 1, last_col + 1).value = Application.WorksheetFunction.Transpose(data_rng)
        .Range("A1").value = "Key"
        .Range("A2:A" & 2 + last_row).FormulaR1C1 = "=CONCAT(RC[1],""-"",RC[2])"
        data_wb.Close savechanges:=False
        
    End With
End Sub


Sub facility_test()
    getFacilityData "25MB08905200"

End Sub



Public Sub getTrackingData(status As String)
    Dim db As New YFdb
    Dim data As Variant
    Dim result As Variant
    Dim last_col As Long, last_row As Long, last_sht_row As Long
    Dim data_wb As Workbook, data_sht As Worksheet, data_rng As Range
    Dim col_last_row As Long

    
    
    data = db.getTracking(status)
    
    With trackDatasht
        last_col = UBound(data, 1)
        last_row = UBound(data, 2)
         .Range(.Cells(2, 1), .Cells(.Rows.Count, .columns.Count)).ClearContents
        Set data_wb = Workbooks.Add
        Set data_sht = data_wb.Sheets(1)
        data_sht.Range("A1").Resize(last_col + 1, last_row + 1).value = data
        Set data_rng = data_sht.UsedRange
        .Range("B2").Resize(last_row + 1, last_col + 1).value = Application.WorksheetFunction.Transpose(data_rng)
        .Range("A1").value = "Key"
        .Range("A2:A" & 2 + last_row).FormulaR1C1 = "=CONCAT(RC[1],""-"",RC[2])"
        .Range("H2:H" & 2 + last_row).NumberFormat = "mm/dd/yyyy"
        data_wb.Close savechanges:=False
        
    End With
End Sub



Public Sub getMasterListData()
    Dim db As New YFdb
    Dim data As Variant
    Dim result As Variant
    Dim last_col As Long, last_row As Long, last_sht_row As Long
    Dim data_wb As Workbook, data_sht As Worksheet, data_rng As Range, export_sht As Worksheet
    
    Dim col_last_row As Long

    
    
    data = db.getMasterTrackingList()
    Set data_wb = Workbooks.Add
    Set data_sht = data_wb.Sheets(1)
    Set export_sht = data_wb.Sheets.Add
    With data_sht
        last_col = UBound(data, 1)
        last_row = UBound(data, 2)
        .Range("A1").Resize(last_col + 1, last_row + 1).value = data
        Set data_rng = data_sht.UsedRange
    End With
    With export_sht
        col_last_row = stampHolderCols.Cells(stampHolderCols.Rows.Count, 3).End(xlUp).Row
        .Range(.Cells(1, 1), .Cells(1, 1 + col_last_row)).value = _
        Application.WorksheetFunction.Transpose(stampHolderCols.Range("C1:C" & col_last_row))
        .Range("A2").Resize(last_row + 1, last_col + 1).value = Application.WorksheetFunction.Transpose(data_rng)
        .name = "master_list_export"
    End With
    data_wb.Worksheets("Sheet1").Delete
    data_wb.SaveAs ThisWorkbook.Path & "\master_list_export.xlsx"
End Sub

