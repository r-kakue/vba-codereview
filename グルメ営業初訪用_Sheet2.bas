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
Num = Range("A3")  '繰り返す店舗数
waitTime = Range("B3")  '待ち時間の指定

'chromeを起動する
Dim Driver As New ChromeDriver
    
Do While myrro < Num

    '食べログから情報を取ってくる
    Driver.Get Range("D6").Offset(0, myrro) '開くURLの参照元
    Application.Wait Now + waitTime
    
    '店舗情報（詳細）を調査
    Set obj0 = Driver.FindElementById("rst-data-head").FindElementsByTag("tr")
        For Each obj In obj0
            If obj.FindElementsByTag("th")(1).Text = "ジャンル" Then
                Range("D9").Offset(0, myrro) = obj.FindElementsByTag("td")(1).Text
            ElseIf obj.FindElementsByTag("th")(1).Text = "席数" Then
                Range("D10").Offset(0, myrro) = obj.FindElementsByTag("td")(1).FindElementsByTag("p")(1).Text
            ElseIf obj.FindElementsByTag("th")(1).Text = "予算" Then
                Range("D11").Offset(0, myrro) = obj.FindElementsByTag("td")(1).FindElementsByTag("em")(1).Text
            ElseIf obj.FindElementsByTag("th")(1).Text = "個室" Then
                Range("D12").Offset(0, myrro) = obj.FindElementsByTag("td")(1).FindElementsByTag("p")(1).Text
            ElseIf obj.FindElementsByTag("th")(1).Text = "貸切" Then
                Range("D13").Offset(0, myrro) = obj.FindElementsByTag("td")(1).FindElementsByTag("p")(1).Text
            End If
            
            '食べログのプラン
            TLplan = Driver.FindElementsByTag("html")(1).Attribute("innerHTML")
            prop15 = InStr(TLplan, "s.prop15")
            plan15 = Mid(TLplan, prop15 + 12, 3)
            Range("D19").Offset(0, myrro) = Replace(Replace(plan15, """", ""), ";", "")
        Next
    
    '食べログの点数
    If Driver.FindElementsByClass("rdheader-rating__score-val-dtl").Count > 0 Then
        Range("D21").Offset(0, myrro) = Driver.FindElementsByClass("rdheader-rating__score-val-dtl")(1).Text
    End If
    
    '住所を取得
    If Driver.FindElementsByClass("rstinfo-table__address").Count > 0 Then
        Range("D5").Offset(0, myrro) = Driver.FindElementsByClass("rstinfo-table__address")(1).Text
    End If


    'ホットペッパーから情報を取ってくる
    Driver.Get Range("D7").Offset(0, myrro) '開くURLの参照元
    Application.Wait Now + waitTime
       
    'ホットペッパーのプラン
    hpplan0 = Driver.FindElementsByTag("html")(1).Attribute("innerHTML")
    If InStr(hpplan0, "storeDivision") > 0 Then
        hpplan1 = Mid(hpplan0, InStr(hpplan0, "storeDivision") + 16, 4)
        hpplan2 = Replace(Replace(hpplan1, """", ""), ";", "")
        Range("D16").Offset(0, myrro) = hpplan2
    End If
    
    'ホットペッパーの業サポ探す
    Driver.Get "https://www.google.com/search?q=" & Range("D3").Offset(0, myrro) & "owst.jp"  '店名にowst.jpをつけて検索する
    Application.Wait Now + waitTime
    
    all_html = Driver.FindElementsByClass("g")(1).Attribute("innerHTML")
    mid_http = Mid(all_html, InStr(all_html, "http"))
    final_http = Left(mid_http, InStr(mid_http, """") - 1)  '最初検索結果のURLを取って来る
    sapourl = final_http
    searchurl = "owst.jp"
    
    If InStr(sapourl, searchurl) > 0 Then
        Range("D18").Offset(0, myrro).Value = sapourl
    Else
        Range("D18").Offset(0, myrro).Value = "該当なし"
    End If

    'GBPの情報を取って来る
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
    
        Range("D24").Offset(0, myrro) = Driver.FindElementsByClass("dbg0pd")(j).FindElementByTag("span").Attribute("innerText")  '店舗名を取得
        Range("D25").Offset(0, myrro) = Driver.FindElementsByCss(".yi40Hd.YrbPuc")(j).Attribute("innerText")  '点数を取得
        Range("D26").Offset(0, myrro) = Driver.FindElementsByCss(".RDApEe.YrbPuc")(j).Attribute("innerText") * -1  '口コミ件数を取得
        
        '新着の投稿があるか確認
        If Driver.FindElementsByClass("o0AaCd").Count > 0 Then
            Range("D27").Offset(0, myrro) = "〇"
        Else
            Range("D27").Offset(0, myrro) = "×"
        End If
        
        '口コミ返信があるか確認
        Driver.FindElementsByCss(".F3Istb.sSWCId")(3).Click
        If Driver.FindElementsByClass("KmCjbd").Count > 0 Then
            Range("D28").Offset(0, myrro) = "〇"
        Else
            Range("D28").Offset(0, myrro) = "×"
        End If
        
        'メニュー入力があるか確認
        Driver.FindElementsByCss(".F3Istb.sSWCId")(2).Click
        If Driver.FindElementsByClass("gq9CCd").Count > 0 Then
            Range("D29").Offset(0, myrro) = "〇"
        Else
            Range("D29").Offset(0, myrro) = "×"
        End If

    'Instagramの情報を取って来る
    Driver.Get "https://www.instagram.com/" & Range("D31").Offset(0, myrro)
    Application.Wait Now + waitTime
        
        'フォロワー数を取得
        For tabs = 1 To Driver.FindElementsByTag("button").Count
            If InStr(Driver.FindElementsByTag("button")(tabs).Attribute("innerText"), "フォロワー") Then
                Range("D32").Offset(0, myrro) = Driver.FindElementsByTag("button")(tabs).FindElementsByTag("span")(1).Attribute("innerText")
            End If
        Next
        
        'フォロー中の数を取得
        For tabs = 1 To Driver.FindElementsByTag("button").Count
            If InStr(Driver.FindElementsByTag("button")(tabs).Attribute("innerText"), "フォロー中") Then
                Range("D33").Offset(0, myrro) = Driver.FindElementsByTag("button")(tabs).FindElementsByTag("span")(1).Attribute("innerText")
            End If
        Next
    Application.Wait Now + waitTime
    Driver.Quit

myrro = myrro + 1
Loop

TexBox = "完了"

End Sub










