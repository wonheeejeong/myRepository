Attribute VB_Name = "Module1"
Sub start()

    Application.DisplayAlerts = False
    
    Dim FilePath As String
    Dim FileName As String
    
    today = Mid(Sheets("Sheet1").Cells(1, 4).Value, 1, 4) + Mid(Sheets("Sheet1").Cells(1, 4).Value, 6, 2) + _
                    Mid(Sheets("Sheet1").Cells(1, 4).Value, 9, 2)

    limit = Sheets("Sheet1").Cells(1, 7).Value
    
    FilePath = Application.ThisWorkbook.Path & "\데이터\"
    FileName = "전종목시세_" + today + ".xlsx"

    Call openFile(FilePath, FileName)

    ActiveSheet.UsedRange.AutoFilter

    Range("A2:N10000").Sort key1:=Range("G2"), order1:=xlDescending

    Set dataList = Range("A2", Cells(Rows.Count, "A").End(xlUp))

    curNum_KOSPI = 1
    curNum_KOSDAQ = 1
    curRow_KOSPI = 2
    curRow_KOSDAQ = 2

    For Each i In dataList
        If i.Offset(0, 2).Value = "KOSPI" And i.Offset(0, 12) >= 500000000000# And curNum_KOSPI <= limit Then
            ThisWorkbook.Sheets("코스피").Cells(curRow_KOSPI, 1) = i                                         '종목코드'
            ThisWorkbook.Sheets("코스피").Cells(curRow_KOSPI, 2) = i.Offset(0, 1).Value      '종목명'
            ThisWorkbook.Sheets("코스피").Cells(curRow_KOSPI, 3) = i.Offset(0, 4).Value      '종가'
            ThisWorkbook.Sheets("코스피").Cells(curRow_KOSPI, 4) = i.Offset(0, 5).Value      '대비'
            ThisWorkbook.Sheets("코스피").Cells(curRow_KOSPI, 5) = i.Offset(0, 6).Value      '등락률'
            ThisWorkbook.Sheets("코스피").Cells(curRow_KOSPI, 6) = i.Offset(0, 10).Value     '거래량'
            ThisWorkbook.Sheets("코스피").Cells(curRow_KOSPI, 7) = i.Offset(0, 11).Value     '거래대금'
            ThisWorkbook.Sheets("코스피").Cells(curRow_KOSPI, 8) = i.Offset(0, 12).Value     '시가총액'
            ThisWorkbook.Sheets("코스피").Cells(curRow_KOSPI, 9) = i.Offset(0, 13).Value     '주식수'
            curRow_KOSPI = curRow_KOSPI + 1
            curNum_KOSPI = curNum_KOSPI + 1
        ElseIf i.Offset(0, 2).Value = "KOSDAQ" And i.Offset(0, 12) >= 200000000000# And curNum_KOSDAQ <= limit Then
            ThisWorkbook.Sheets("코스닥").Cells(curRow_KOSDAQ, 1) = i                                         '종목코드'
            ThisWorkbook.Sheets("코스닥").Cells(curRow_KOSDAQ, 2) = i.Offset(0, 1).Value      '종목명'
            ThisWorkbook.Sheets("코스닥").Cells(curRow_KOSDAQ, 3) = i.Offset(0, 4).Value      '종가'
            ThisWorkbook.Sheets("코스닥").Cells(curRow_KOSDAQ, 4) = i.Offset(0, 5).Value      '대비'
            ThisWorkbook.Sheets("코스닥").Cells(curRow_KOSDAQ, 5) = i.Offset(0, 6).Value      '등락률'
            ThisWorkbook.Sheets("코스닥").Cells(curRow_KOSDAQ, 6) = i.Offset(0, 10).Value     '거래량'
            ThisWorkbook.Sheets("코스닥").Cells(curRow_KOSDAQ, 7) = i.Offset(0, 11).Value     '거래대금'
            ThisWorkbook.Sheets("코스닥").Cells(curRow_KOSDAQ, 8) = i.Offset(0, 12).Value     '시가총액'
            ThisWorkbook.Sheets("코스닥").Cells(curRow_KOSDAQ, 9) = i.Offset(0, 13).Value     '주식수'
            curRow_KOSDAQ = curRow_KOSDAQ + 1
            curNum_KOSDAQ = curNum_KOSDAQ + 1
        End If

    Next

    Call closeFile

    Sheets("코스피").Select
    Set dataList_KOSPI = Range("A2", Cells(Rows.Count, "A").End(xlUp))
    
    curRow = 3
    For Each a In dataList_KOSPI
        Sheets("코스피코스닥").Cells(curRow, 1).Value = a
        Sheets("코스피코스닥").Cells(curRow, 2).Value = a.Offset(0, 1)
        Sheets("코스피코스닥").Cells(curRow, 3).Value = a.Offset(0, 2)
        Sheets("코스피코스닥").Cells(curRow, 4).Value = a.Offset(0, 3)
        Sheets("코스피코스닥").Cells(curRow, 5).Value = a.Offset(0, 4)
        Sheets("코스피코스닥").Cells(curRow, 6).Value = a.Offset(0, 5)
        Sheets("코스피코스닥").Cells(curRow, 7).Value = a.Offset(0, 6)
        Sheets("코스피코스닥").Cells(curRow, 8).Value = a.Offset(0, 7)
        Sheets("코스피코스닥").Cells(curRow, 9).Value = a.Offset(0, 8)
        curRow = curRow + 1
    Next
    
    curRow = curRow + 2
    With Sheets("코스피코스닥").Cells(curRow, 1)
        .Value = "코스닥"
        .Font.Size = 16
        .RowHeight = 25.2
    End With
    
    Sheets("코스닥").Select
    Set dataList_KOSDAQ = Range("A2", Cells(Rows.Count, "A").End(xlUp))
    
    curRow = curRow + 1
    Sheets("코스피코스닥").Cells(curRow, 1).Value = Sheets("코스피코스닥").Cells(2, 1).Value
    Sheets("코스피코스닥").Cells(curRow, 2).Value = Sheets("코스피코스닥").Cells(2, 2).Value
    Sheets("코스피코스닥").Cells(curRow, 3).Value = Sheets("코스피코스닥").Cells(2, 3).Value
    Sheets("코스피코스닥").Cells(curRow, 4).Value = Sheets("코스피코스닥").Cells(2, 4).Value
    Sheets("코스피코스닥").Cells(curRow, 5).Value = Sheets("코스피코스닥").Cells(2, 5).Value
    Sheets("코스피코스닥").Cells(curRow, 6).Value = Sheets("코스피코스닥").Cells(2, 6).Value
    Sheets("코스피코스닥").Cells(curRow, 7).Value = Sheets("코스피코스닥").Cells(2, 7).Value
    Sheets("코스피코스닥").Cells(curRow, 8).Value = Sheets("코스피코스닥").Cells(2, 8).Value
    Sheets("코스피코스닥").Cells(curRow, 9).Value = Sheets("코스피코스닥").Cells(2, 9).Value
    curRow = curRow + 1
    
    For Each b In dataList_KOSDAQ
        Sheets("코스피코스닥").Cells(curRow, 1).Value = b
        Sheets("코스피코스닥").Cells(curRow, 2).Value = b.Offset(0, 1)
        Sheets("코스피코스닥").Cells(curRow, 3).Value = b.Offset(0, 2)
        Sheets("코스피코스닥").Cells(curRow, 4).Value = b.Offset(0, 3)
        Sheets("코스피코스닥").Cells(curRow, 5).Value = b.Offset(0, 4)
        Sheets("코스피코스닥").Cells(curRow, 6).Value = b.Offset(0, 5)
        Sheets("코스피코스닥").Cells(curRow, 7).Value = b.Offset(0, 6)
        Sheets("코스피코스닥").Cells(curRow, 8).Value = b.Offset(0, 7)
        Sheets("코스피코스닥").Cells(curRow, 9).Value = b.Offset(0, 8)
        curRow = curRow + 1
    Next
    
    Sheets("코스피코스닥").Select
    
    Range("A2:L1000").Select
    Selection.Style = "Comma [0]"
    
    
    Sheets("뉴스정보_코스피").Cells(2, 1) = Sheets("코스피").Cells(2, 2)
    Sheets("뉴스정보_코스피").Cells(12, 1) = Sheets("코스피").Cells(3, 2)
    Sheets("뉴스정보_코스피").Cells(22, 1) = Sheets("코스피").Cells(4, 2)
    Sheets("뉴스정보_코스피").Cells(32, 1) = Sheets("코스피").Cells(5, 2)
    Sheets("뉴스정보_코스피").Cells(42, 1) = Sheets("코스피").Cells(6, 2)
    Sheets("뉴스정보_코스피").Cells(52, 1) = Sheets("코스피").Cells(7, 2)
    Sheets("뉴스정보_코스피").Cells(62, 1) = Sheets("코스피").Cells(8, 2)
    Sheets("뉴스정보_코스피").Cells(72, 1) = Sheets("코스피").Cells(9, 2)
    Sheets("뉴스정보_코스피").Cells(82, 1) = Sheets("코스피").Cells(10, 2)
    Sheets("뉴스정보_코스피").Cells(92, 1) = Sheets("코스피").Cells(11, 2)

    Sheets("뉴스정보_코스닥").Cells(2, 1) = Sheets("코스닥").Cells(2, 2)
    Sheets("뉴스정보_코스닥").Cells(12, 1) = Sheets("코스닥").Cells(3, 2)
    Sheets("뉴스정보_코스닥").Cells(22, 1) = Sheets("코스닥").Cells(4, 2)
    Sheets("뉴스정보_코스닥").Cells(32, 1) = Sheets("코스닥").Cells(5, 2)
    Sheets("뉴스정보_코스닥").Cells(42, 1) = Sheets("코스닥").Cells(6, 2)
    Sheets("뉴스정보_코스닥").Cells(52, 1) = Sheets("코스닥").Cells(7, 2)
    Sheets("뉴스정보_코스닥").Cells(62, 1) = Sheets("코스닥").Cells(8, 2)
    Sheets("뉴스정보_코스닥").Cells(72, 1) = Sheets("코스닥").Cells(9, 2)
    Sheets("뉴스정보_코스닥").Cells(82, 1) = Sheets("코스닥").Cells(10, 2)
    Sheets("뉴스정보_코스닥").Cells(92, 1) = Sheets("코스닥").Cells(11, 2)
    
 
    
    

    savePath = Application.ThisWorkbook.Path & "\데이터\" + "결과값_" + today + ".xlsm"
    ActiveWorkbook.SaveAs FileName:=savePath

    

End Sub

Function openFile(FilePath As String, FileName As String)

    Application.Workbooks.Open FileName:=FilePath & FileName
    
End Function

Function closeFile()

    ActiveWorkbook.Close False

End Function

Function getHtml(url As String) As String
    Set httpObj = CreateObject("MSXML2.XMLHTTP")
    httpObj.Open "GET", url, False
    httpObj.send
    getHtml = httpObj.responseText
End Function
