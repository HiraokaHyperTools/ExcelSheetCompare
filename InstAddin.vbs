Set App = CreateObject("Excel.Application")

Rem http://blogs.msdn.com/b/accelerating_things/archive/2010/09/16/loading-excel-add-ins-at-runtime.aspx

fnf = True
For Each myAddIn In App.AddIns
	If UCase(myAddIn.Name) = UCase("�G�N�Z���t�@�C�����r.xlam") Then
		fnf = False
		If myAddIn.Installed Then
			WScript.Echo "�������ς݂ł��B"
		Else
			WScript.Echo "���������܂����B"
			myAddIn.Installed = True
		End If
	End If
Next

If fnf Then
	WScript.Echo "������܂���B"
End If

App.Quit