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
