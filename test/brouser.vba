Sub BrouserStart()

    Const vbHide = 0             'ウィンドウを非表示
    Const vbNormalFocus = 1      '通常のウィンドウ、かつ最前面のウィンドウ
    Const vbMinimizedFocus = 2   '最小化、かつ最前面のウィンドウ
    Const vbMaximizedFocus = 3   '最大化、かつ最前面のウィンドウ
    Const vbNormalNoFocus = 4    '通常のウィンドウ、ただし、最前面にはならない
    Const vbMinimizedNoFocus = 6 '最小化、ただし、最前面にはならない

    Dim obj As Object
    Dim objWSshell
    Dim ApplicationFilePath As String
    Dim urlString As String
    Dim doubleQuote As String
    Dim sPath As String
    
    'IE起動
    'set obj = createobject("InternetExplorer.application")
    
    '任意のapplicationの起動
    Set objWShell = CreateObject("WScript.Shell")
    
    
    '任意のURLを入力
    urlString = "http://www.madoka-magica.com/"
    
    'applicationの".exe"fileへのpathを指定
    ApplicationFilePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" '注意"""のように「"」は3個いる模様
    
    doubleQuote = """"
    
    ApplicationFilePath = doubleQuote & ApplicationFilePath & doubleQuote & doubleQuote & urlString & doubleQuote
    
    'MsgBox ApplicationFilePath 'コメント確認
    objWShell.Run ApplicationFilePath, vbNormalFocus
    
    
    
End Sub










Sub GoogleSearch()
    Dim objIE As Object
   
    'IE起動
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
   
    'Googleに接続
    objIE.navigate "https://my.redmine.jp/demo/projects/demo/issues"
   
    'IEを待機
    waitNavigation objIE
   
    '検索窓に「VBA」と入力
    'objIE.document.getElementById("lst-id").Value = "VBA"
    'objIE.document.getelementbyclass ("btn btn-sm")
    
    '5秒停止
    'Call WaitFor(5)
   
    '検索ボタンを押す
    Call IEButtonClick(objIE) ', "Download ZIP")
 
    '5秒停止
    'Call WaitFor(5)
   
    'IE終了
    'objIE.Quit
   
    Set objIE = Nothing
End Sub
 
'ボタンを押す関数
Public Function IEButtonClick(ByRef objIE As Object) ', buttonValue As String)
    Dim objInput As Object
   
   
    
    For i = 0 To objIE.document.all.Length - 1
    If objIE.document.all(i).tagName = "OPTION" Then
           If objIE.document.all(i).Value = "estimated_hours" Then
               objIE.document.all(i).Selected = True
                objIE.document.all(i).Click
               Exit For
        End If
    End If
    Next i
    
    For i = 0 To objIE.document.all.Length - 1
        If objIE.document.all(i).ID = "add_filter_select" Then
            If objIE.document.all(i).tagName = "SELECT" Then
            objIE.document.all(i).selectedIndex = "1"
            objIE.document.all(i).Click
            End If
        End If
    Next
    
    For i = 0 To objIE.document.all.Length - 1
       If objIE.document.all(i).ID = "query_form" Then
       objIE.document.all(i).submit
       End If
    Next
    
    
End Function
 

 Sub waitNavigation(ie As Object)
    Do While ie.Busy Or ie.ReadyState < 4
        DoEvents
    Loop
End Sub
