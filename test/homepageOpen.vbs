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
objIE.document.forms("csv-export-form").submit
call "text"




'Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hwnd As Long) As Long

Private Const WM_COMMAND As Long = &H111
Private Const CLASSNAME_ÀÞ²±Û¸Þ As String = "#32770"
Private Const CLASSNAME_ÎÞÀÝ   As String = "Button"

Private Sub Test()
  Const STEP1_TITLE  As String = "Ì§²ÙÉÀÞ³ÝÛ°ÄÞ"
  Const STEP2_TITLE  As String = "–¼‘O¦•t¹Ã•Û‘¶"

  Dim l_lngWnd_Window_Step1  As Long
  Dim l_lngWnd_Window_Step2  As Long
  
  'step1
  If Not FindDialog(STEP1_TITLE, l_lngWnd_Window_Step1) Then
    Exit Sub
  End If
  Call PushSaveBtn(l_lngWnd_Window_Step1)
  
  
  'step2
  If Not FindDialog(STEP2_TITLE, l_lngWnd_Window_Step2) Then
    Exit Sub
  End If
  Call PushSaveBtn(l_lngWnd_Window_Step2)
End Sub

'ÀÞ²±Û¸Þ¦’T½
Private Function FindDialog(ByVal p_strCaption As String, ByRef p_lngFindWnd As Long) As Boolean
  p_lngFindWnd = 0
  Do
    DoEvents
'    If Not IE¶ÞBUSY Then
'      Exit Do
'    End If
    p_lngFindWnd = FindWindow(CLASSNAME_ÀÞ²±Û¸Þ, p_strCaption)
  Loop While p_lngFindWnd = 0
  
  FindDialog = p_lngFindWnd <> 0
End Function

'ÎÞÀÝ¦‰Ÿ½
Private Sub PushSaveBtn(ByVal p_lngWindowWnd As Long, Optional p_strBtnCaption As String = "•Û‘¶(&S)")
  Dim l_lngWnd_Save  As Long
  l_lngWnd_Save = FindWindowEx(p_lngWindowWnd, 0, CLASSNAME_ÎÞÀÝ, p_strBtnCaption)
  Call SendMessage(p_lngWindowWnd, WM_COMMAND, GetDlgCtrlID(l_lngWnd_Save), ByVal l_lngWnd_Save)
End Sub
