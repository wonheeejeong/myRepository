Attribute VB_Name = "Module1"
Sub 일평균생성()

    Sheets("일평균거래대금").Select
    Set List = Range("B9", Range("B9").End(xlToRight))
    Set dateList = Range("A15", Cells(Rows.Count, "A").End(xlUp))

    For Each i In List
        Set tmpList = Worksheets("거래대금").Range("B9", "BQC9")
        Set target = tmpList.Find(what:=i.Value, lookat:=xlWhole)
        cnt = 0
        daegeum = 0
        For Each j In dateList
            If Worksheets("거래대금").Cells(j.Row, target.Column).Value > 0 Then
                cnt = cnt + 1
                daegeum = daegeum + Worksheets("거래대금").Cells(j.Row, target.Column).Value
                Worksheets("일평균거래대금").Cells(j.Row, i.Column).Value = daegeum / cnt
            End If
        Next
    Next
    
    
    Sheets("일평균시가총액").Select
    Set List = Range("B9", Range("B9").End(xlToRight))
    Set dateList = Range("A15", Cells(Rows.Count, "A").End(xlUp))
    
    For Each i In List
        Set tmpList = Worksheets("시가총액").Range("B9", "BQC9")
        Set target = tmpList.Find(what:=i.Value, lookat:=xlWhole)
        cnt = 0
        daegeum = 0
        For Each j In dateList
            If Worksheets("시가총액").Cells(j.Row, target.Column).Value > 0 Then
                cnt = cnt + 1
                daegeum = daegeum + Worksheets("시가총액").Cells(j.Row, target.Column).Value
                Worksheets("일평균시가총액").Cells(j.Row, i.Column).Value = daegeum / cnt
            End If
        Next
    Next
    
    MsgBox "end"

End Sub


Sub 선정과정()

    
    
    Dim sectorList(9, 1 To 100) As Integer
    Dim sector As String

    Sheets("일평균시가총액").Select
    Set dateList = Range("A15", Cells(Rows.Count, "A").End(xlUp))
    'Set dateList = Range("A15")
    Set companyList = Range("B9", Range("B9").End(xlToRight))
    Sheets("기존구성종목").Select
    Set beforeList = Range("B2", Cells(Rows.Count, "B").End(xlUp))

    
    Sheets("일평균시가총액").Select
   
    
    
        For Each d In dateList
             col = 0
            For i = 0 To 9
                Select Case i
                    Case 0
                        sector = "에너지"
                    Case 1
                        sector = "소재"
                    Case 2
                        sector = "산업재"
                    Case 3
                        sector = "자유소비재"
                    Case 4
                        sector = "필수소비재"
                    Case 5
                        sector = "헬스케어"
                    Case 6
                        sector = "금융 및 부동산"
                    Case 7
                        sector = "정보기술"
                    Case 8
                        sector = "커뮤니케이션서비스"
                    Case 9
                        sector = "유틸리티"
                End Select

                Total = 0
                cur = 0
                cnt = 0
                cnt_b = 0
    
                Dim arrV(200) As Double     '섹터별 일평균시가총액'
                Dim arrC(200) As String     '섹터별 기업명'
                
                Dim arrV_t(200) As Double   '섹터별 일평균거래대금'
                Dim arrC_t(200) As String   '섹터별 기업명'
                
                Dim arrV_b(300) As Double   '기존종목 일평균시가총액'
                Dim arrC_b(300) As String   '기존종목 일평균거래대금'
                
                For Each Item In beforeList
                    If Item.Offset(0, 5) = sector Then
                        Set target = companyList.Find(what:=Item.Value, lookat:=xlWhole)
                        arrV_b(cnt_b) = Worksheets("일평균시가총액").Cells(d.Row, target.Column).Value
                        arrC_b(cnt_b) = Item.Value
                        cnt_b = cnt_b + 1
                    End If
                Next
                
                For Each company In companyList
                    If company.Offset(1, 0) = sector Then
                        On Error Resume Next
                        'MsgBox company & Cells(d.Row, company.Column).Value
                        Total = Total + Cells(d.Row, company.Column).Value
                        arrV(cnt) = Cells(d.Row, company.Column).Value
                        arrC(cnt) = company
                        arrV_t(cnt) = Worksheets("일평균거래대금").Cells(d.Row, company.Column).Value
                        arrC_t(cnt) = company
                        cnt = cnt + 1
                    End If
                Next

                For p = 0 To cnt_b - 1
                    Debug.Print arrV_b(p) & ", " & arrC_b(p)
                    Debug.Print ""
                Next
                 
                Debug.Print "************************"
                 

                Dim SortTemp As Double
                Dim companyTemp As String
                 
                For n = 0 To cnt - 1
                   For m = 0 To cnt - n
                       If arrV(m) < arrV(m + 1) Then
                           SortTemp = arrV(m + 1)
                           companyTemp = arrC(m + 1)
                           arrV(m + 1) = arrV(m)
                           arrC(m + 1) = arrC(m)
                           arrV(m) = SortTemp
                           arrC(m) = companyTemp
                       End If
                   Next
               Next
                
                For n = 0 To cnt - 1
                    For m = 0 To cnt - n
                       If arrV_t(m) < arrV_t(m + 1) Then
                           SortTemp = arrV_t(m + 1)
                           companyTemp = arrC_t(m + 1)
                           arrV_t(m + 1) = arrV_t(m)
                           arrC_t(m + 1) = arrC_t(m)
                           arrV_t(m) = SortTemp
                           arrC_t(m) = companyTemp
                       End If
                    Next
                Next

                For p = 0 To cnt - 1
                    Debug.Print arrV(p) & ", " & arrC(p)
                    Debug.Print ""
                Next
                 
                deadline = WorksheetFunction.RoundDown(WorksheetFunction.VLookup(sector, Sheets("산업군(코스피)").Range("F2:G11"), 2, 0) * 0.85, 0)
                
                For p = 0 To cnt - 1
                    Dim flag As Boolean: flag = False
                    cur = cur + arrV(p)
                    Worksheets("1차선정").Cells(d.Row, col + 2) = arrC(p)
                    Worksheets("1차선정").Cells(d.Row, col + 2).Font.Color = RGB(255, 255, 255)
                    'Worksheets("1차선정").Cells(d.Row, col + 2) = arrV(p)
                    'Worksheets("1차선정").Cells(9, col + 2) = arrC(p)
                    'Worksheets("1차선정").Cells(9, col + 2).Font.Color = RGB(255, 255, 255)
                    'Worksheets("1차선정").Cells(10, col + 2) = sector
                    If Total * 0.85 > cur Then
                        Worksheets("1차선정").Cells(d.Row, col + 2).Interior.Color = RGB(255, 0, 0)
                        For q = 0 To deadline - 1
                            If arrC(p) = arrC_t(q) Then
                                flag = True
                            End If
                        Next
                        If flag = False Then
                            'MsgBox "in"
                            Worksheets("1차선정").Cells(d.Row, col + 2).Interior.Color = RGB(102, 0, 204)
                        End If
                    Else
                        Worksheets("1차선정").Cells(d.Row, col + 2).Interior.Color = RGB(0, 0, 255)
                    End If
                    col = col + 1
                Next
                
                Erase arrV: Erase arrC: Erase arrV_t: Erase arrC_t: Erase arrV_b: Erase arrC_b
                
            Next
        Next
        
        Worksheets("1차선정").Select
        
        
        
        MsgBox "end"
    



End Sub


Sub test()

    Sheets("1차선정").Select
    Set companyList = Range("B9", Range("B9").End(xlToRight))
    cnt = 0
    
    For Each company In companyList
        If company.Interior.Color = RGB(255, 0, 0) Then
            cnt = cnt + 1
        End If
    Next
    
    MsgBox cnt
    
    deadline = WorksheetFunction.VLookup("에너지", Sheets("산업군").Range("F2:G11"), 2, 0) * 0.85
    MsgBox deadline


End Sub


Public Function clear()
Dim x As Long
For x = 1 To 10
Debug.Print x
Next

Debug.Print Now
Application.SendKeys "^g ^a {DEL}"


End Function


Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
