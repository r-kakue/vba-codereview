VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub eigyousiryou()

myrro = 0
Num = Range("A3")  '�J��Ԃ��X�ܐ�
waitTime = Range("B3")  '�҂����Ԃ̎w��

'chrome���N������
Dim Driver As New ChromeDriver
    
Do While myrro < Num

    '�H�׃��O�����������Ă���
    Driver.Get Range("D6").Offset(0, myrro) '�J��URL�̎Q�ƌ�
    Application.Wait Now + waitTime
    
    '�X�܏��i�ڍׁj�𒲍�
    Set obj0 = Driver.FindElementById("rst-data-head").FindElementsByTag("tr")
        For Each obj In obj0
            If obj.FindElementsByTag("th")(1).Text = "�W������" Then
                Range("D9").Offset(0, myrro) = obj.FindElementsByTag("td")(1).Text
            ElseIf obj.FindElementsByTag("th")(1).Text = "�Ȑ�" Then
                Range("D10").Offset(0, myrro) = obj.FindElementsByTag("td")(1).FindElementsByTag("p")(1).Text
            ElseIf obj.FindElementsByTag("th")(1).Text = "�\�Z" Then
                Range("D11").Offset(0, myrro) = obj.FindElementsByTag("td")(1).FindElementsByTag("em")(1).Text
            ElseIf obj.FindElementsByTag("th")(1).Text = "��" Then
                Range("D12").Offset(0, myrro) = obj.FindElementsByTag("td")(1).FindElementsByTag("p")(1).Text
            ElseIf obj.FindElementsByTag("th")(1).Text = "�ݐ�" Then
                Range("D13").Offset(0, myrro) = obj.FindElementsByTag("td")(1).FindElementsByTag("p")(1).Text
            End If
            
            '�H�׃��O�̃v����
            TLplan = Driver.FindElementsByTag("html")(1).Attribute("innerHTML")
            prop15 = InStr(TLplan, "s.prop15")
            plan15 = Mid(TLplan, prop15 + 12, 3)
            Range("D19").Offset(0, myrro) = Replace(Replace(plan15, """", ""), ";", "")
        Next
    
    '�H�׃��O�̓_��
    If Driver.FindElementsByClass("rdheader-rating__score-val-dtl").Count > 0 Then
        Range("D21").Offset(0, myrro) = Driver.FindElementsByClass("rdheader-rating__score-val-dtl")(1).Text
    End If
    
    '�Z�����擾
    If Driver.FindElementsByClass("rstinfo-table__address").Count > 0 Then
        Range("D5").Offset(0, myrro) = Driver.FindElementsByClass("rstinfo-table__address")(1).Text
    End If


    '�z�b�g�y�b�p�[�����������Ă���
    Driver.Get Range("D7").Offset(0, myrro) '�J��URL�̎Q�ƌ�
    Application.Wait Now + waitTime
       
    '�z�b�g�y�b�p�[�̃v����
    hpplan0 = Driver.FindElementsByTag("html")(1).Attribute("innerHTML")
    If InStr(hpplan0, "storeDivision") > 0 Then
        hpplan1 = Mid(hpplan0, InStr(hpplan0, "storeDivision") + 16, 4)
        hpplan2 = Replace(Replace(hpplan1, """", ""), ";", "")
        Range("D16").Offset(0, myrro) = hpplan2
    End If
    
    '�z�b�g�y�b�p�[�̋ƃT�|�T��
    Driver.Get "https://www.google.com/search?q=" & Range("D3").Offset(0, myrro) & "owst.jp"  '�X����owst.jp�����Č�������
    Application.Wait Now + waitTime
    
    all_html = Driver.FindElementsByClass("g")(1).Attribute("innerHTML")
    mid_http = Mid(all_html, InStr(all_html, "http"))
    final_http = Left(mid_http, InStr(mid_http, """") - 1)  '�ŏ��������ʂ�URL������ė���
    sapourl = final_http
    searchurl = "owst.jp"
    
    If InStr(sapourl, searchurl) > 0 Then
        Range("D18").Offset(0, myrro).Value = sapourl
    Else
        Range("D18").Offset(0, myrro).Value = "�Y���Ȃ�"
    End If

    'GBP�̏�������ė���
    Driver.Get "https://www.google.com/search?tbs=lf:1,lf_ui:9&tbm=lcl&hl&q=" & Range("D3").Offset(0, myrro) & " " & Range("D5").Offset(0, myrro)
    Application.Wait Now + waitTime
    
    jmax = Driver.FindElementsByClass("uMdZh").Count
    
    For j = 1 To jmax
        If Driver.FindElementsByClass("uMdZh")(j).FindElementsByClass("pXf2tf").Count = 0 Then
            Driver.FindElementsByClass("uMdZh")(j).FindElementsByClass("OSrXXb")(1).Click
            Application.Wait Now + waitTime
        Exit For
        End If
    Next
    
        Range("D24").Offset(0, myrro) = Driver.FindElementsByClass("dbg0pd")(j).FindElementByTag("span").Attribute("innerText")  '�X�ܖ����擾
        Range("D25").Offset(0, myrro) = Driver.FindElementsByCss(".yi40Hd.YrbPuc")(j).Attribute("innerText")  '�_�����擾
        Range("D26").Offset(0, myrro) = Driver.FindElementsByCss(".RDApEe.YrbPuc")(j).Attribute("innerText") * -1  '���R�~�������擾
        
        '�V���̓��e�����邩�m�F
        If Driver.FindElementsByClass("o0AaCd").Count > 0 Then
            Range("D27").Offset(0, myrro) = "�Z"
        Else
            Range("D27").Offset(0, myrro) = "�~"
        End If
        
        '���R�~�ԐM�����邩�m�F
        Driver.FindElementsByCss(".F3Istb.sSWCId")(3).Click
        If Driver.FindElementsByClass("KmCjbd").Count > 0 Then
            Range("D28").Offset(0, myrro) = "�Z"
        Else
            Range("D28").Offset(0, myrro) = "�~"
        End If
        
        '���j���[���͂����邩�m�F
        Driver.FindElementsByCss(".F3Istb.sSWCId")(2).Click
        If Driver.FindElementsByClass("gq9CCd").Count > 0 Then
            Range("D29").Offset(0, myrro) = "�Z"
        Else
            Range("D29").Offset(0, myrro) = "�~"
        End If

    'Instagram�̏�������ė���
    Driver.Get "https://www.instagram.com/" & Range("D31").Offset(0, myrro)
    Application.Wait Now + waitTime
        
        '�t�H�����[�����擾
        For tabs = 1 To Driver.FindElementsByTag("button").Count
            If InStr(Driver.FindElementsByTag("button")(tabs).Attribute("innerText"), "�t�H�����[") Then
                Range("D32").Offset(0, myrro) = Driver.FindElementsByTag("button")(tabs).FindElementsByTag("span")(1).Attribute("innerText")
            End If
        Next
        
        '�t�H���[���̐����擾
        For tabs = 1 To Driver.FindElementsByTag("button").Count
            If InStr(Driver.FindElementsByTag("button")(tabs).Attribute("innerText"), "�t�H���[��") Then
                Range("D33").Offset(0, myrro) = Driver.FindElementsByTag("button")(tabs).FindElementsByTag("span")(1).Attribute("innerText")
            End If
        Next
    Application.Wait Now + waitTime
    Driver.Quit

myrro = myrro + 1
Loop

TexBox = "����"

End Sub










