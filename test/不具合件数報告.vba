Sub Filter()
    Dim elements As String          'フィルターをかける文字
    Dim FieldColumn1 As Integer      'フィールドの列を指定
    Dim FieldColumn2 As Integer
    Dim FieldColumn3 As Integer
    Const FieldColumnName1 As String = "作成者"   '検索したいフィールドの名前
    Const FieldColumnName2 As String = "作成日"
    Const FieldColumnName3 As String = "外部Report*"
    Dim colum As Integer
    Dim row As Integer
    Dim sheetname As String
    Dim FailureNum As Integer
    Dim FailureNumToday As Integer
    '不具合件数
    
    
    
    row = 1 '行を入力
    colum = 1  '列を入力
    'sheetname = "0001"     'シート名を入力
    
       Sheets(1).Activate
    
    Do While True
        If (Cells(row, colum).Value = FieldColumnName1) Then
            FieldColumn1 = colum
        End If
        
        If (Cells(row, colum).Value = FieldColumnName2) Then
            FieldColumn2 = colum
        End If
        
          If (Cells(row, colum).Value = FieldColumnName3) Then
            FieldColumn3 = colum
        End If
        
        If (colum > Cells(1, Columns.Count).End(xlToLeft)) Then
            Exit Do
        End If
        
        colum = colum + 1
    Loop
    
    'step1全体の不具合件数
    FieldColumn1 = 2    'testのため
    Rows(1).AutoFilter Field:=FieldColumn1, Criteria1:="<>*QA*"
      
    If Cells(Rows.Count, 1).End(xlUp).row = 1 Then
        FailureNum = 0
    Else
        FailureNum = Range(Range("A2"), Cells(Rows.Count, 1).End(xlUp)) _
                .SpecialCells(xlCellTypeVisible).Count
    End If
    
    
    'step2今日日付のもの
    FieldColumn2 = 2    'testのため
    Rows(1).AutoFilter Field:=FieldColumn2, Criteria1:="*" & Month(Date) & "*" & Day(Date)
    
    
    If Cells(Rows.Count, 1).End(xlUp).row = 1 Then
        FailureNumToday = 0
    Else
        FailureNumToday = Range(Range("A2"), Cells(Rows.Count, 1).End(xlUp)) _
                .SpecialCells(xlCellTypeVisible).Count
    End If
    
    'step3
    FieldColumn3 = 2    'testのため
    Sheets("Sheet1").Select
    Sheets.Add after:=Worksheets(Worksheets.Count)
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Sheets(1).name & "!" & Cells(1, FieldColumn3).Address(ReferenceStyle:=xlR1C1) & ":" _
        & Cells(Rows.Count, FieldColumn3).End(xlUp).Address(ReferenceStyle:=xlR1C1), Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="Sheet5!R3C1", TableName:="ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ2", DefaultVersion _
        :=xlPivotTableVersion14
    Sheets("Sheet5").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ2").PivotFields("7")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ2").AddDataField ActiveSheet.PivotTables( _
        "ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ2").PivotFields("7"), "合計 / 7", xlSum
    Range("A4").Select
    
    
End Sub
