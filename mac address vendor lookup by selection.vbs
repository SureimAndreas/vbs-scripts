# $language = "VBScript"
# $interface = "1.0"

' @ SureimAndreas
' Search for mac addres vendor by using macvendors.co api by marking the mac address and push hotkey.

Sub Main()

	strSelection = Trim(crt.Screen.Selection)

	Set g_shell = CreateObject("WScript.Shell")
	
	Dim url, req, json
	url = "https://macvendors.co/api/vendorname/" & strSelection

	Set req = CreateObject("Msxml2.XMLHttp.6.0")
	Call req.Open("GET", url, False)
	Call req.Send()

	If req.Status = 200 Then
  	  json = req.responseText
	End If

	If strSelection = "" Then
		Exit Sub
	Else
		msgbox(json)
		End If
End Sub
