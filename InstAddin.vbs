Set App = CreateObject("Excel.Application")

Rem http://blogs.msdn.com/b/accelerating_things/archive/2010/09/16/loading-excel-add-ins-at-runtime.aspx

fnf = True
For Each myAddIn In App.AddIns
	If UCase(myAddIn.Name) = UCase("エクセルファイルを比較.xlam") Then
		fnf = False
		If myAddIn.Installed Then
			WScript.Echo "活性化済みです。"
		Else
			WScript.Echo "活性化しました。"
			myAddIn.Installed = True
		End If
	End If
Next

If fnf Then
	WScript.Echo "見つかりません。"
End If

App.Quit