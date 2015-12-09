Set objIE = Wscript.CreateObject("InternetExplorer.Application")
objIE.Navigate2 "https://my.redmine.jp/demo/projects/demo/issues"
objIE.Visible = TRUE

    Do While objIE.Busy Or objIE.ReadyState < 4
        WScript.Sleep 300
    Loop

For i = 0 To objIE.document.all.Length - 1
       If objIE.document.all(i).className ="csv" 	Then
       	objIE.document.all(i).click
       End If
    Next

WScript.Sleep 300

For i = 0 To objIE.document.all.Length - 1
       If objIE.document.all(i).id ="columns_all" 	Then
       	objIE.document.all(i).checked	= true
       End If
    Next
WScript.Sleep 300

objIE.document.forms("csv-export-form").submit

WScript.Sleep 10000


Set wsh = WScript.CreateObject("WScript.Shell")







Sub FileDownLoad_Proc()
Dim strCaption As String
Dim PWnd As IntPtr
Dim cWnd As IntPtr

' 親ウィンドウ取得
strCaption = "○○○○ - Windows Internet Explorer"
While PWnd = 0
PWnd = FindWindowEx(0, 0, "IEFrame", strCaption)
System.Threading.Thread.Sleep(50)
End While

' 通知バーのハンドル
While cWnd = 0
cWnd = FindWindowEx(PWnd, 0&, "Frame Notification Bar", vbNullString)
System.Threading.Thread.Sleep(50)
End While

' 通知バーボタン群のハンドル
Dim hChild As IntPtr = FindWindowEx(cWnd, 0&, "DirectUIHWND", vbNullString)
Dim objAcc As IAccessible = Nothing

AccessibleObjectFromWindow(hChild, OBJID_CLIENT, IID_IAccessible, objAcc)

If Not IsNothing(objAcc) Then
ClickPreserve(objAcc)
While cWnd = 0
cWnd = FindWindowEx(PWnd, 0&, "Frame Notification Bar", vbNullString)
System.Threading.Thread.Sleep(50)
End While
SendMessage(cWnd, WM_QUIT, 0, 0&)

End If

End Sub
Private Sub ClickPreserve(ByVal acc As IAccessible)

Dim i As Long
Dim count = acc.accChildCount
Dim lst(count - 1) As Object

If count > 0 Then
AccessibleChildren(acc, 0, count, lst, 0)
If Not IsNothing(lst) Then
For i = LBound(lst) To UBound(lst)
With lst(i)
'On Error Resume Next
'Debug.Print("ChildCount: " & .accChildCount)
'Debug.Print("Value: " & .accValue(CHILDID_SELF))
'Debug.Print("Name: " & .accName(CHILDID_SELF))
'Debug.Print("Description: " & .accDescription(CHILDID_SELF))
'On Error GoTo 0
'保存ボタンを見つけたらクリック（デフォルトアクション）する
If .accName(CHILDID_SELF) = "保存" Then

System.Threading.Thread.Sleep(500)
.accDoDefaultAction(CHILDID_SELF)
System.Threading.Thread.Sleep(500)
End If
End With
ClickPreserve(lst(i)) '再帰
Next
End If
End If
End Sub
