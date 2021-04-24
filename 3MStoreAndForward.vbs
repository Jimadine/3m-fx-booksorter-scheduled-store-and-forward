' Created:    21/04/2021
' Modified:   24/04/2021
' Version:    1.1
' Author:     Jim Adamson
' Code formatted with http://www.vbindent.com
Option Explicit
Dim blnTestMode,colNamedArguments,intNumItemsToProcess,objHTTP,objShell,strBaseUrl,strCookie,strCurrentMode,strForwardSleepTime,strInductionPcName,strIntelligentReturnSystemManagerPassword,strMode,strResponse,strWhen
Set colNamedArguments = WScript.Arguments.Named
Set objShell = WScript.CreateObject("WScript.Shell")

' Check the web admin interface password is set as a user environment variable
If Not objShell.Environment("USER").Item("IntelligentReturnSystemManagerPassword") = "" Then
  strIntelligentReturnSystemManagerPassword = objShell.Environment("USER").Item("IntelligentReturnSystemManagerPassword")
Else
  WScript.Echo "The password must be set as a user environment variable with name IntelligentReturnSystemManagerPassword"
  WScript.Quit 1
End If

' Check whether a "/inductionpcname:string" has been supplied. If not, default to localhost
If IsEmpty(colNamedArguments.Item("inductionpcname")) Then
  strInductionPcName = "localhost"
Else
  strInductionPcName = colNamedArguments.Item("inductionpcname")
End If

' Check whether "/forwardsleeptime:seconds" has been supplied. If not, default to 2 minutes
If IsEmpty(colNamedArguments.Item("forwardsleeptime")) Then
  strForwardSleepTime = 120000
Else
  strForwardSleepTime = colNamedArguments.Item("forwardsleeptime")*1000
End If

' Check whether "/testmode" has been supplied. If so, this forces the mode changes & forwarding process to happen
If colNamedArguments.Exists("testmode") Then
    blnTestMode = true
Else
    blnTestMode = false
End If

strBaseUrl = "http://" & strInductionPcName

' Authenticate
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.open "POST", strBaseUrl & "/IntelligentReturn/pages/Index.aspx", False
' Don't follow redirects
objHTTP.Option(6) = False
objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.send "password=" & strIntelligentReturnSystemManagerPassword
strCookie = objHTTP.getResponseHeader("Set-Cookie")
If Left(strCookie, 17) = "ASP.NET_SessionId" Then
  Wscript.Echo Now() & ": " & "Authenticating using session cookie"
Else
  Wscript.Echo Now() & ": " & "Problem logging in"
  WScript.Quit 1
End If
Set objHTTP = nothing

' Check number of items BEFORE forwarding
intNumItemsToProcess = countItems("Before")

If intNumItemsToProcess >= 1 Or blnTestMode = true Then
  ' Set the Operation Mode to OUT OF SERVICE, while the store/forwarding process is done
    strCurrentMode = changeMode("OUT_OF_SERVICE")
    If strCurrentMode = "OUT_OF_SERVICE" Then
  ' Forward items
      wscript.echo Now() & ": Proceeding with Store/Forward"
      Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
      objHTTP.open "GET", strBaseUrl & "/IntelligentReturn/pages/StoreAndForwardStart.aspx", False
      objHTTP.SetRequestHeader "Cookie", strCookie
      objHTTP.send
      Set objHTTP = nothing
  ' Allow time for the items to be processed
      WScript.Sleep strForwardSleepTime
  ' Check number of items AFTER forwarding
      intNumItemsToProcess = countItems("After")
  ' Set mode back to normal
      strCurrentMode = changeMode("NORMAL")
      If strCurrentMode = "NORMAL" Then
        wscript.echo Now() & ": Script finished normally - exiting"
      Else
        wscript.echo Now() & ": Problem with mode setting mode to NORMAL - exiting"
        Wscript.quit 1
      End If
    Else
      wscript.echo Now() & ": Problem with mode setting mode to OUT_OF_SERVICE - exiting"
      Wscript.quit 1
    End If
Else
  wscript.echo Now() & ": Exiting early as 0 items to process"
  WScript.Quit
End If

Function countItems(strWhen)
' Check number of items for processing
  Dim objNumItemsToProcessElement,strNumItemsToProcessElementInnertext,objHtmlFile
  Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
  Set objHtmlFile = CreateObject("htmlfile")
  objHTTP.open "GET", strBaseUrl & "/IntelligentReturn/pages/StoreAndForward.aspx", False
  objHTTP.SetRequestHeader "Cookie", strCookie
  objHTTP.send
  If objHTTP.Status = 200 Then
    objHtmlFile.Write objHTTP.ResponseText
    objHtmlFile.Close
    Set objNumItemsToProcessElement = objHtmlFile.getElementById("ctl00_MainContentPlaceHolder_lblItemCount")
    strNumItemsToProcessElementInnertext = objNumItemsToProcessElement.Innertext
    countItems = CInt(strNumItemsToProcessElementInnertext)
    wscript.echo Now() & ": (" & strWhen & ") " & strNumItemsToProcessElementInnertext & " item(s) to process"
    strNumItemsToProcessElementInnertext=empty
  End If
  Set objHTTP = nothing
  Set objHtmlFile = nothing
End Function

Function changeMode(strMode)
  'Change the mode of operation
  Dim objModeElement,objHtmlFile,strModeElementSelectedValue,strRedirectLocation
  Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
  Set objHtmlFile = CreateObject("htmlfile")
  objHTTP.open "POST", strBaseUrl & "/IntelligentReturn/pages/Support.aspx", False
  objHTTP.SetRequestHeader "Cookie", strCookie
  objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  objHTTP.send "mode=" & strMode & "&submit=Set+Mode"
  wscript.echo Now() & ": Attempting to set operation mode to " & strMode
' Vista WinHttp doesn't seem to follow 302 redirects, so handle the redirect manually
  If objHTTP.Status = 302 Then
    strRedirectLocation = objHTTP.getResponseHeader("Location")
    Set objHTTP = nothing
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    objHTTP.open "GET", strBaseUrl & strRedirectLocation, False
    objHTTP.SetRequestHeader "Cookie", strCookie
    objHTTP.send
  End If
  If objHTTP.Status = 200 Then
    objHtmlFile.Write objHTTP.ResponseText
    objHtmlFile.Close
    Set objModeElement = objHtmlFile.getElementById("mode")
    strModeElementSelectedValue = objModeElement.options(objModeElement.selectedIndex).Value
    wscript.echo Now() & ": Operation mode currently set to " & strModeElementSelectedValue
    changeMode = strModeElementSelectedValue
    Set objHTTP = nothing
    Set objHtmlFile = nothing
  Else
    wscript.echo Now() & ": Unexpected HTTP status code: " & objHTTP.Status
    Set objHTTP = nothing
    Set objHtmlFile = nothing
  End If
End Function
