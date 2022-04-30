Attribute VB_Name = "매크로"
Sub 실행()

    Dim FilePath As String
    Dim FileName As String

    
    Dim strAddr As String
    Dim firstAddr As String
    
    Application.ScreenUpdating = False
    
    curRowNum = 2
    curFileNum = 1
    
    FilePath = Application.ThisWorkbook.Path & "\월별데이터목록\"
    FileName = Dir(FilePath & "*.csv")
    
    Dim companyList
    companyList = Array("한양이엔지")
    

    While FileName <> ""

        Call openFile(FilePath, FileName)

        Set dataList = Range("B2", Cells(Rows.Count, "B").End(xlUp))
        Debug.Print "현재 번호 :  " & curFileNum
        curFileNum = curFileNum + 1

        For Each company In companyList
            Set target = dataList.Find(what:=company, lookat:=xlPart)

            If Not target Is Nothing Then
                strAddr = target.Address
                firstAddr = target.Address

                Do
                    ''MsgBox "자료생성년월 : " & target.Offset(0, -1) & "사업장명 : " & target & ", 가입자수 :  " & target.Offset(0, 17) & ", 읍면동코드 : " & target.Offset(0, 10)'
                    ThisWorkbook.Sheets("데이터모음").Cells(curRowNum, 1) = target.Offset(0, -1)
                    ThisWorkbook.Sheets("데이터모음").Cells(curRowNum, 2) = target.Offset(0, 0)
                    ThisWorkbook.Sheets("데이터모음").Cells(curRowNum, 3) = target.Offset(0, 17)
                    ThisWorkbook.Sheets("데이터모음").Cells(curRowNum, 4) = target.Offset(0, 10)
                    curRowNum = curRowNum + 1
                    Set target = dataList.FindNext(target)
                Loop While Not target Is Nothing And target.Address <> firstAddr
            End If
        Next

        Call closeFile

        FileName = Dir()

    Wend

    MsgBox "end"

    
End Sub



Function openFile(FilePath As String, FileName As String)

    Application.Workbooks.Open FileName:=FilePath & FileName
    
End Function

Function closeFile()

    ActiveWorkbook.Close False

End Function

Public Function clear()
    Dim x As Long
    For x = 1 To 10
    Debug.Print x
    Next
    
    Debug.Print Now
    Application.SendKeys "^g ^a {DEL}"

End Function

Sub test()
    
    Dim companyList
    companyList = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J")
    
    For Each company In companyList
        Debug.Print company
    Next
    
    Call clear

End Sub

