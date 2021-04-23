' Created:    21/04/2021
' Modified:   21/04/2021
' Version:    1.0
' Author:     Jim Adamson
' Code formatted with http://www.vbindent.com
Option Explicit
Dim FORWARDSLEEPTIME,INDUCTION_PC_NAME,INTELLIGENTRETURNSYSTEMMANAGERPASSWORD,baseurl,strCookie,strResponse,colArgs,objHTTP,objShell,ohtmlFile,numItemsToProcess,numItemsToProcessInnertext,currentMode,minNumItemsToProcess,modeElement,mode,selectedValue,numItemsToProcessElement,redirectLocation,when
Set colArgs = WScript.Arguments.Named
Set objShell = WScript.CreateObject("WScript.Shell")

' Check whether a named argument has been supplied. If not default to localhost
If colargs.Exists("inductionpcname") Then
  INDUCTION_PC_NAME = colArgs.Item("inductionpcname")
Else
  INDUCTION_PC_NAME = "localhost"
End If

If colargs.Exists("forwardsleeptime") Then
  FORWARDSLEEPTIME = colArgs.Item("forwardsleeptime")*1000
Else
' 2 minutes
  FORWARDSLEEPTIME = 12000
End If

minNumItemsToProcess = 0
If colargs.Exists("testmode") And colArgs.Item("testmode") = "true" Then
    minNumItemsToProcess = 1
End If

' Check the environment variable is set
If Not objShell.Environment("USER").Item("IntelligentReturnSystemManagerPassword") = "" Then
  INTELLIGENTRETURNSYSTEMMANAGERPASSWORD = objShell.Environment("USER").Item("IntelligentReturnSystemManagerPassword")
Else
  WScript.Echo "The password must be set as a user environment variable with name IntelligentReturnSystemManagerPassword"
  WScript.Quit 1
End If

baseurl = "http://" & INDUCTION_PC_NAME

' Authenticate
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.open "POST", baseurl & "/IntelligentReturn/pages/Index.aspx", False
' Don't follow redirects
objHTTP.Option(6) = False
objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.send "password=" & INTELLIGENTRETURNSYSTEMMANAGERPASSWORD
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

If numItemsToProcess = minNumItemsToProcess Then
  wscript.echo Now() & ": Exiting early as 0 items to process"
  WScript.Quit
Else
' Set the Operation Mode to OUT OF SERVICE, while the store/forwarding process is done
  currentMode = changeMode("OUT_OF_SERVICE")
  If currentMode = "OUT_OF_SERVICE" Then
' Forward items
    wscript.echo Now() & ": Proceeding with Store/Forward"
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    objHTTP.open "GET", baseurl & "/IntelligentReturn/pages/StoreAndForwardStart.aspx", False
    objHTTP.SetRequestHeader "Cookie", strCookie
    objHTTP.send
    Set objHTTP = nothing
' Allow time for the items to be processed
    WScript.Sleep FORWARDSLEEPTIME
' Check number of items AFTER forwarding
    numItemsToProcess = countItems("After")
' Set mode back to normal
    currentMode = changeMode("NORMAL")
  Else
    wscript.echo Now() & ": Problem with mode set - exiting"
    Wscript.quit 1
  End If
End If

Function countItems(when)
' Check number of items for processing
  Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
  Set ohtmlFile = CreateObject("htmlfile")
  objHTTP.open "GET", baseurl & "/IntelligentReturn/pages/StoreAndForward.aspx", False
  objHTTP.SetRequestHeader "Cookie", strCookie
  objHTTP.send
  If objHTTP.Status = 200 Then
    ohtmlFile.Write objHTTP.ResponseText
    ohtmlFile.Close
    Set numItemsToProcessElement = ohtmlFile.getElementById("ctl00_MainContentPlaceHolder_lblItemCount")
    numItemsToProcessInnertext = numItemsToProcessElement.Innertext
    countItems = CInt(numItemsToProcessInnertext)
    wscript.echo Now() & ": (" & when & ") " & numItemsToProcessInnertext & " item(s) to process"
    numItemsToProcessInnertext=empty
  End If
  Set objHTTP = nothing
  Set ohtmlFile = nothing
End Function

Function changeMode(mode)
  Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
  Set ohtmlFile = CreateObject("htmlfile")
  objHTTP.open "POST", baseurl & "/IntelligentReturn/pages/Support.aspx", False
  objHTTP.SetRequestHeader "Cookie", strCookie
  objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  objHTTP.send "mode=" & mode & "&submit=Set+Mode"
  wscript.echo Now() & ": Attempting to set operation mode to " & mode
' Vista WinHttp doesn't seem to follow 302 redirects, so handle the redirect manually
  If objHTTP.Status = 302 Then
    redirectLocation = objHTTP.getResponseHeader("Location")
    Set objHTTP = nothing
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    objHTTP.open "GET", baseurl & redirectLocation, False
    objHTTP.SetRequestHeader "Cookie", strCookie
    objHTTP.send
  End If
  If objHTTP.Status = 200 Then
    ohtmlFile.Write objHTTP.ResponseText
    ohtmlFile.Close
    Set modeElement = ohtmlFile.getElementById("mode")
    selectedValue = modeElement.options(modeElement.selectedIndex).Value
    wscript.echo Now() & ": Operation mode currently set to " & selectedValue
    changeMode = selectedValue
    Set objHTTP = nothing
    Set ohtmlFile = nothing
  Else
    wscript.echo Now() & ": Unexpected HTTP status code: " & objHTTP.Status
    Set objHTTP = nothing
    Set ohtmlFile = nothing
  End If
End Function
