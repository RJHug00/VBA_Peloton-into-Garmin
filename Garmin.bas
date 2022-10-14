Attribute VB_Name = "Garmin"
Option Explicit

Private Const GARMIN_BASE_HOST As String = "connect.garmin.com"
Private Const GARMIN_BASE As String = "https://" & GARMIN_BASE_HOST
Private Const GARMIN_SSO_HOST As String = "sso.garmin.com"
Private Const GARMIN_SSO As String = "https://" & GARMIN_SSO_HOST
Private Const GARMIN_SSO_SIGNIN_URL As String = GARMIN_SSO & "/sso/signin"
Private GARMIN_BASE_ENCODED As String

Private Const USERAGENT As String = "Mozilla/5.0"
Private Const MULTIPART_BOUNDARY As String = "--NonsensicalDidacticJibberish"

Private cookie_jar As Object

Public Function UploadFITFile(fName As String) As Boolean

  Dim oHttpReq As Object, byteArr() As Byte, body() As Byte, fileInt As Integer
  Dim rspHeaders As String, qS As String, t As String, i As Integer

  Set cookie_jar = CreateObject("Scripting.Dictionary")

  Set oHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
  oHttpReq.SetTimeouts 30000, 30000, 30000, 30000

  GARMIN_BASE_ENCODED = EncodeURL(GARMIN_BASE)

  ' Acquire Cookies we need for the login
  ' Attempt to minimize the backend processing, but its mostly just guesswork
  qS = "?clientId=GarminConnect"
  qS = qS & "&embedWidget=false"
  qS = qS & "&mobile=false"
  qS = qS & "&consumeServiceTicket=false"
  qS = qS & "&generateExtraServiceTicket=false"
  qS = qS & "&generateTwoExtraServiceTickets=false"

  oHttpReq.Open "GET", GARMIN_SSO_SIGNIN_URL & qS & "&generateNoServiceTicket=true", False
  oHttpReq.setRequestHeader "User-Agent", USERAGENT
  oHttpReq.setRequestHeader "Origin", GARMIN_SSO
  oHttpReq.Send ""

  If oHttpReq.Status <> 200 Then
    MsgBox "Initial Garmin HTTP GET failed: " & oHttpReq.StatusText, vbOKOnly, "Error"
    UploadFITFile = False: Exit Function
  End If

  rspHeaders = oHttpReq.GetAllResponseHeaders()

  ' Augment the queryString with just essentials
  qS = qS & "&generateNoServiceTicket=false"
  qS = qS & "&service=" & GARMIN_BASE_ENCODED
  qS = qS & "&gauthHost=" & EncodeURL(GARMIN_SSO)
  qS = qS & "&mfaRequired=false"
  qS = qS & "&performMFACheck=false"
  qS = qS & "&id=gauth-widget"
 'qS = qS & "&redirectAfterAccountCreationUrl" & GARMIN_BASE_ENCODED
 'qS = qS & "&cssUrl=" & EncodeURL("https://static.garmincdn.com/com.garmin.connect/ui/css/gauth-custom-v1.2-min.css")
 'qS = qS & "&initialFocus=true"
 'qS = qS & "&rememberMeShown=false"
 'qS = qS & "&rememberMeChecked=false"
 'qS = qS & "&createAccountShown=false"
 'qS = qS & "&openCreateAccount=false"
 'qS = qS & "&displayNameShown=false"
 'qS = qS & "&globalOptInShown=false"
 'qS = qS & "&globalOptInChecked=false"
 'qS = qS & "&connectLegalTerms=true"
 'qS = qS & "&locationPromptShown=false"
 'qS = qS & "&showPassword=false"
 'qS = qS & "&useCustomHeader=false"
 'qS = qS & "&rememberMyBrowserShown=false"
 'qS = qS & "&rememberMyBrowserChecked=false"
 'qS = qS & "&socialEnabled=false"
 'qS = qS & "&showTermsOfUse=false"
 'qS = qS & "&showPrivacyPolicy=false"
 'qS = qS & "&showConnectLegalAge=false"
 'qS = qS & "&locale=en_US"
 'qS = qS & "&source=" & GARMIN_BASE_ENCODED
 'qS = qS & "&webhost=" & GARMIN_BASE_ENCODED
 'qS = qS & "&redirectAfterAccountLoginUrl=" & GARMIN_BASE_ENCODED

  ' Garmin login phase-1: Obtain service ticket
  oHttpReq.Open "POST", GARMIN_SSO_SIGNIN_URL & qS, False
  oHttpReq.setRequestHeader "User-Agent", USERAGENT
  oHttpReq.setRequestHeader "Origin", GARMIN_SSO
  oHttpReq.setRequestHeader "Host", GARMIN_SSO_HOST
  oHttpReq.setRequestHeader "Referer", GARMIN_SSO_SIGNIN_URL
  oHttpReq.setRequestHeader "Cookie", GrabCookies(rspHeaders)
  oHttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  oHttpReq.Send "embed=false" & "&username=" & Environ("GARMIN_USER") & "&password=" & Environ("GARMIN_KEY")

  If oHttpReq.Status <> 200 Then
    MsgBox "Garmin login phase-1 failed: " & oHttpReq.StatusText, vbOKOnly, "Error"
    UploadFITFile = False: Exit Function
  End If

  rspHeaders = oHttpReq.GetAllResponseHeaders()

  t = oHttpReq.responseText   ' extract the service ticket ID
  i = InStr(t, "?ticket="): t = Mid(t, i + 8, 200): t = Left(t, InStr(t, """") - 1)

  If i < 1 Then
    MsgBox "Garmin login phase-1 did not return a service ticket", vbOKOnly, "Error"
    UploadFITFile = False: Exit Function
  End If

  ' Login phase 2
  oHttpReq.Open "GET", GARMIN_BASE & "/modern/?ticket=" & t, False
  oHttpReq.setRequestHeader "Cookie", GrabCookies(rspHeaders)
  oHttpReq.Send ""

  If oHttpReq.Status <> 200 Then
    MsgBox "Garmin login phase-2 failed: " & oHttpReq.StatusText, vbOKOnly, "Error"
    UploadFITFile = False: Exit Function
  End If

  rspHeaders = oHttpReq.GetAllResponseHeaders()

  If InStr(oHttpReq.responseText, "class=""signed-in") < 1 Then
    MsgBox "Garmin phase-2 login didn't complete as expected", vbOKOnly, "Error"
    UploadFITFile = False: Exit Function
  End If

  fileInt = FreeFile
  Open fName For Binary Access Read As #fileInt
  ReDim byteArr(0 To LOF(fileInt) - 1)
  Get #fileInt, , byteArr
  Close #fileInt

  oHttpReq.Open "POST", GARMIN_BASE & "/modern/proxy/upload-service/upload/.fit", False
  oHttpReq.setRequestHeader "Cookie", GrabCookies(rspHeaders)
  oHttpReq.setRequestHeader "User-Agent", USERAGENT
  oHttpReq.setRequestHeader "Origin", GARMIN_BASE
  oHttpReq.setRequestHeader "Host", GARMIN_BASE_HOST
  oHttpReq.setRequestHeader "Referer", GARMIN_BASE & "/modern/import-data"
  oHttpReq.setRequestHeader "NK", "NT"
  oHttpReq.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & MULTIPART_BOUNDARY
  body = BuildMultipartBody( _
       "--" & MULTIPART_BOUNDARY & vbCrLf & _
       "Content-Disposition: form-data; name=""file""; filename=""x.FIT""" & vbCrLf & _
       "Content-Type: application/octet-stream" & vbCrLf & vbCrLf, _
       byteArr, _
       vbCrLf & "--" & MULTIPART_BOUNDARY & "--")
  oHttpReq.Send (body) ' Without parenthesis, this gets "The parameter is incorrect" error

  If oHttpReq.Status = "202" Then UploadFITFile = True   ' 202 == "Accepted" == Success

  If oHttpReq.Status = "403" Then
    MsgBox "Failure: HTTP Error 403 - Are you logged into Garmin elsewhere" & vbLf & _
                    " or has this ride been uploaded already?", vbOKOnly, "Error"
  End If

  ' the responseText can be JSON, containing "errorId" and "error"
  ' or it can be HTML <!DOCTYPE... containing "signed-out"

  UploadFITFile = False  ' failed
End Function

' Maintain a catalog of current cookies; return an HTTP "Cookie" header string from it
Private Function GrabCookies(ByRef s As String) As String
  Dim t As String, v As String, c As Variant, i As Integer, j As Integer
  i = InStr(s, "Set-Cookie:")
  While i > 0
    t = Mid(s, i + 12, 256): t = Left(t, InStr(t, ";") - 1)
    j = InStr(t, "="): v = Left(t, j - 1): t = Mid(t, j + 1)
    If cookie_jar.Exists(v) Then cookie_jar.Remove v  ' remove the older one
    cookie_jar.Add v, RTrim(t)
    i = InStr(i + 1, s, "Set-Cookie:")  ' find next one
  Wend
  For Each c In cookie_jar.Keys: GrabCookies = GrabCookies & "; " & c & "=" & cookie_jar(c): Next c
  GrabCookies = Mid(GrabCookies, 3) ' drop the leading delimiter
End Function

' Brute-force, but I didn't want this nonsense in-line in the main routine
Private Function BuildMultipartBody(ByRef a As String, ByRef b() As Byte, ByRef c As String) As Byte()
  Dim x() As Byte, j As Long, i As Long
  Dim n1 As Long: n1 = Len(a)
  Dim n2 As Long: n2 = UBound(b)
  Dim n3 As Long: n3 = Len(c)
  ReDim x(n1 + n2 + n3): i = 0
  For j = 1 To n1: x(i) = Asc(Mid(a, j, 1)): i = i + 1: Next j
  For j = 0 To n2: x(i) = b(j): i = i + 1: Next j
  For j = 1 To n3: x(i) = Asc(Mid(c, j, 1)): i = i + 1: Next j
  BuildMultipartBody = x
End Function

Private Function EncodeURL(s As String) As String
  EncodeURL = Replace(s, " ", "%20")
  EncodeURL = Replace(EncodeURL, ":", "%3A")
  EncodeURL = Replace(EncodeURL, "/", "%2F")
End Function
