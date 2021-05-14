Sub CombineWSs()
    Dim wbDst       As Workbook
    Dim wbSrc       As Workbook
    Dim wsSrc       As Worksheet
    Dim MyPath      As String
    On Error Resume Next
    Dim strFilename As String
    
    Application.DisplayAlerts = FALSE
    Application.EnableEvents = FALSE
    Application.ScreenUpdating = FALSE
    
    MyPath = "D:\Testpress\Input"
    Set wbDst = ThisWorkbook
    strFilename = Dir(MyPath & "\*.csv", vbNormal)
    
    If Len(strFilename) = 0 Then Exit Sub
    Do Until strFilename = ""
        Set wbSrc = Workbooks.Open(Filename:=MyPath & "\" & strFilename)
        Set wsSrc = wbSrc.Worksheets(1)
        wsSrc.Copy After:=wbDst.Worksheets(wbDst.Worksheets.Count)
        wbSrc.Close FALSE
        strFilename = Dir()
        Call getdata
    Loop
    Application.DisplayAlerts = TRUE
    Application.EnableEvents = TRUE
    Application.ScreenUpdating = TRUE
    
    MsgBox "Work Done Shashi!"
    
End Sub

Private Sub getdata()
    
    Dim startnum    As Integer
    Dim endnum      As Integer
    lRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    endnum = lRow
    
    ActiveSheet.Cells(1, "Q") = "Question No."
    ActiveSheet.Cells(1, "R") = "Correct Answer"
    
    bb = 2
    
    colu = 2
    
    Questno = 1
    
    For startnum = 1 To endnum
        
        Bcell = ActiveSheet.Cells(startnum, "B")
        
        If Bcell = "Passage" Or Bcell = "MCQ" Then
            
            If Bcell = "Passage" Then
                passageref = ActiveSheet.Cells(startnum, "A").Value
                ref = 0
                For I = 1 To endnum
                    If ActiveSheet.Cells(I, "J") = passageref Then
                        ref = ref + 1
                    End If
                Next I
                
                ref = (Questno + ref) - 1
                ActiveSheet.Cells(bb, "P") = "Passage For Question No (" & Questno & "-" & ref & ") " & ActiveSheet.Cells(startnum, "G")
                link = ActiveSheet.Cells(bb, "P")
                link = Replace(link, "<br>", vbCrLf)
                Call removetag(link, bb)
                bb = bb + 1
                bb = bb + 1
            End If
            
            If Bcell = "MCQ" Then
                ActiveSheet.Cells(bb, "P") = "Q." & Questno & ")" & " " & Trim(ActiveSheet.Cells(startnum, "G").Value)
                Questno = Questno + 1
                link = ActiveSheet.Cells(bb, "P")
                link = Replace(link, "<br>", vbCrLf)
                Call removetag(link, bb)
                
                ActiveSheet.Cells(colu, "Q") = Questno - 1
                ActiveSheet.Cells(colu, "R") = LCase(ActiveSheet.Cells(startnum, "H"))
                ActiveSheet.Cells(colu, "S") = "Q." & Questno - 1 & ")" & " " & Trim(ActiveSheet.Cells(startnum, "I").Value)
                
                link = ActiveSheet.Cells(colu, "S")
                link = Replace(link, "<br>", vbCrLf)
                
                Call removetaginsolution(link, colu)
                
                colu = colu + 1
                
                bb = bb + 1
                
                ActiveSheet.Cells(bb, "P") = "a)" & " " & ActiveSheet.Cells(startnum, "K")
                link = ActiveSheet.Cells(bb, "P")
                link = Replace(link, "<br>", vbCrLf)
                Call removetag(link, bb)
                bb = bb + 1
                
                ActiveSheet.Cells(bb, "P") = "b)" & " " & ActiveSheet.Cells(startnum, "L")
                link = ActiveSheet.Cells(bb, "P")
                link = Replace(link, "<br>", vbCrLf)
                Call removetag(link, bb)
                
                bb = bb + 1
                
                ActiveSheet.Cells(bb, "P") = "c)" & " " & ActiveSheet.Cells(startnum, "M")
                link = ActiveSheet.Cells(bb, "P")
                link = Replace(link, "<br>", vbCrLf)
                Call removetag(link, bb)
                bb = bb + 1
                
                ActiveSheet.Cells(bb, "P") = "d)" & " " & ActiveSheet.Cells(startnum, "N")
                link = ActiveSheet.Cells(bb, "P")
                link = Replace(link, "<br>", vbCrLf)
                Call removetag(link, bb)
                bb = bb + 1
                bb = bb + 1
            End If
        End If
        
    Next startnum
    
    Dim nameofsheet As String
    nameofsheet = ActiveSheet.Name
    
    MyFilePath = "D:\Testpress\Output\"
    MyFileName = nameofsheet
    MyFileName = Replace(MyFileName, ".xlsm", " ")
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(MyFilePath & MyFileName & ".doc", True)
    nLastrow = ActiveSheet.Cells(Rows.Count, "P").End(xlUp).Row
    nFirstRow = 1
    
    For N = nFirstRow To nLastrow
        t = Replace(ActiveSheet.Cells(N, "P").Text, Chr(10), vbCrLf)
        a.WriteLine (t)
    Next
    a.Close
    
    MyFileName = nameofsheet & " Explanation"
    MyFileName = Replace(MyFileName, ".xlsm", " ")
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(MyFilePath & MyFileName & ".doc", True)
    nLastrow = ActiveSheet.Cells(Rows.Count, "S").End(xlUp).Row
    nFirstRow = 1
    
    For N = nFirstRow To nLastrow
        t = Replace(ActiveSheet.Cells(N, "S").Text, Chr(10), vbCrLf)
        a.WriteLine (t)
    Next
    a.Close
    
    MyFileName = nameofsheet & " answers"
    MyFileName = Replace(MyFileName, ".xlsm", " ")
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(MyFilePath & MyFileName & ".csv", True)
    nLastrow = ActiveSheet.Cells(Rows.Count, "S").End(xlUp).Row
    nFirstRow = 1
    
    For N = nFirstRow To nLastrow
        q = Replace(ActiveSheet.Cells(N, "Q").Text, Chr(10), vbCrLf)
        r = Replace(ActiveSheet.Cells(N, "R").Text, Chr(10), vbCrLf)
        t = q & "," & r
        a.WriteLine (t)
    Next
    a.Close
    
End Sub

Sub removetag(ByVal link As String, ByVal bb As Integer)
    
    bbb = 50
    For aaa = 1 To bbb
        Position = InStr(link, "<") Or InStr(link, "<")
        If Not Position = 0 Then
            startpos = InStr(link, "<")
            lastpos = InStr(link, ">")
            startpos = startpos
            lastpos = lastpos
            Length = lastpos - startpos
            Tag = Mid(link, startpos, Length + 1)
            link = Replace(link, Tag, "")
            ActiveSheet.Cells(bb, "P") = link
        End If
    Next aaa
    
    ActiveSheet.Cells(bb, "P").Value = link
    link = Replace(link, "……", "_________")
    ActiveSheet.Cells(bb, "P").Value = link
End Sub

Sub removetaginsolution(ByVal link As String, ByVal bb As Integer)
    
    bbb = 50
    For aaa = 1 To bbb
        Position = InStr(link, "<") Or InStr(link, "<")
        If Not Position = 0 Then
            startpos = InStr(link, "<")
            lastpos = InStr(link, ">")
            startpos = startpos
            lastpos = lastpos
            Length = lastpos - startpos
            Tag = Mid(link, startpos, Length + 1)
            link = Replace(link, Tag, "")
            ActiveSheet.Cells(bb, "S") = link
        End If
    Next aaa
    
    ActiveSheet.Cells(bb, "S").Value = link
    link = Replace(link, "……", "_________")
    ActiveSheet.Cells(bb, "S").Value = link
End Sub