Set objIE = Wscript.CreateObject("InternetExplorer.Application")
objIE.Navigate2 "http://www.goo.ne.jp/"
objIE.Visible = TRUE
Set objIE =Nothing