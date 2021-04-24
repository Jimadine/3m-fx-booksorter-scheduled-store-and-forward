' Created:    21/04/2021
' Modified:   24/04/2021
' Version:    1.1
' Author:     Jim Adamson
' Code formatted with http://www.vbindent.com
Option Explicit
Dim baseUrl,colArgs,currentMode,forwardSleepTime,inductionPcName,intelligentReturnSystemManagerPassword,mode,numItemsToProcess,objHTTP,objShell,strCookie,strResponse,testMode,when
Set colArgs = WScript.Arguments.Named
Set objShell = WScript.CreateObject("WScript.Shell")

' Check the web admin interface password is set as a user environment variable
If Not objShell.Environment("USER").Item("IntelligentReturnSystemManagerPassword") = "" Then
  intelligentReturnSystemManagerPassword = objShell.Environment("USER").Item("IntelligentReturnSystemManagerPassword")
Else
  WScript.Echo "The password must be set as a user environment variable with name IntelligentReturnSystemManagerPassword"
  WScript.Quit 1
End If

' Check whether a "/inductionpcname:string" has been supplied. If not, default to localhost
If colargs.Exists("inductionpcname") And Not(IsEmpty(colArgs.Item("inductionpcname"))) Then
  inductionPcName = colArgs.Item("inductionpcname")
Else
  inductionPcName = "localhost"
End If

' Check whether "/forwardsleeptime:seconds" has been supplied. If not, default to 2 minutes
If colargs.Exists("forwardsleeptime") And Not(IsEmpty(colArgs.Item("forwardsleeptime"))) Then
  forwardSleepTime = colArgs.Item("forwardsleeptime")*1000
Else
  forwardSleepTime = 120000
End If

' Check whether "/testmode" has been supplied. If so, this forces the mode changes & forwarding process to happen
If colargs.Exists("testmode") Then
    testMode = true
Else
    testMode = false
End If

baseUrl = "http://" & inductionPcName

' Authenticate
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.open "POST", baseUrl & "/IntelligentReturn/pages/Index.aspx", False
' Don't follow redirects
objHTTP.Option(6) = False
objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.send "password=" & intelligentReturnSystemManagerPassword
strCookie = objHTTP.getResponseHeader("Set-Cookie")
If Left(strCookie, 17) = "ASP.NET_SessionId" Then
  Wscript.Echo Now() & ": " & "Authenticating using session cookie"
Else
  Wscript.Echo Now() & ": " & "Problem logging in"
  WScript.Quit 1
End If
Set objHTTP = nothing

' Check number of items BEFORE forwarding
numItemsToProcess = countItems("Before")

If numItemsToProcess >= 1 Or testMode = true Then
  ' Set the Operation Mode to OUT OF SERVICE, while the store/forwarding process is done
    currentMode = changeMode("OUT_OF_SERVICE")
    If currentMode = "OUT_OF_SERVICE" Then
  ' Forward items
      wscript.echo Now() & ": Proceeding with Store/Forward"
      Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
      objHTTP.open "GET", baseUrl & "/IntelligentReturn/pages/StoreAndForwardStart.aspx", False
      objHTTP.SetRequestHeader "Cookie", strCookie
      objHTTP.send
      Set objHTTP = nothing
  ' Allow time for the items to be processed
      WScript.Sleep forwardSleepTime
  ' Check number of items AFTER forwarding
      numItemsToProcess = countItems("After")
  ' Set mode back to normal
      currentMode = changeMode("NORMAL")
    Else
      wscript.echo Now() & ": Problem with mode set - exiting"
      Wscript.quit 1
    End If
Else
  wscript.echo Now() & ": Exiting early as 0 items to process"
  WScript.Quit
End If

Function countItems(when)
' Check number of items for processing
  Dim numItemsToProcessElement,numItemsToProcessInnertext,objHtmlFile
  Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
  Set objHtmlFile = CreateObject("htmlfile")
  objHTTP.open "GET", baseUrl & "/IntelligentReturn/pages/StoreAndForward.aspx", False
  objHTTP.SetRequestHeader "Cookie", strCookie
  objHTTP.send
  If objHTTP.Status = 200 Then
    objHtmlFile.Write objHTTP.ResponseText
    objHtmlFile.Close
    Set numItemsToProcessElement = objHtmlFile.getElementById("ctl00_MainContentPlaceHolder_lblItemCount")
    numItemsToProcessInnertext = numItemsToProcessElement.Innertext
    countItems = CInt(numItemsToProcessInnertext)
    wscript.echo Now() & ": (" & when & ") " & numItemsToProcessInnertext & " item(s) to process"
    numItemsToProcessInnertext=empty
  End If
  Set objHTTP = nothing
  Set objHtmlFile = nothing
End Function

Function changeMode(mode)
  'Change the mode of operation
  Dim modeElement,objHtmlFile,selectedValue,redirectLocation
  Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
  Set objHtmlFile = CreateObject("htmlfile")
  objHTTP.open "POST", baseUrl & "/IntelligentReturn/pages/Support.aspx", False
  objHTTP.SetRequestHeader "Cookie", strCookie
  objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  objHTTP.send "mode=" & mode & "&submit=Set+Mode"
  wscript.echo Now() & ": Attempting to set operation mode to " & mode
' Vista WinHttp doesn't seem to follow 302 redirects, so handle the redirect manually
  If objHTTP.Status = 302 Then
    redirectLocation = objHTTP.getResponseHeader("Location")
    Set objHTTP = nothing
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    objHTTP.open "GET", baseUrl & redirectLocation, False
    objHTTP.SetRequestHeader "Cookie", strCookie
    objHTTP.send
  End If
  If objHTTP.Status = 200 Then
    objHtmlFile.Write objHTTP.ResponseText
    objHtmlFile.Close
    Set modeElement = objHtmlFile.getElementById("mode")
    selectedValue = modeElement.options(modeElement.selectedIndex).Value
    wscript.echo Now() & ": Operation mode currently set to " & selectedValue
    changeMode = selectedValue
    Set objHTTP = nothing
    Set objHtmlFile = nothing
  Else
    wscript.echo Now() & ": Unexpected HTTP status code: " & objHTTP.Status
    Set objHTTP = nothing
    Set objHtmlFile = nothing
  End If
End Function
