Attribute VB_Name = "ParseJSON"
Option Explicit

' This module implements a quick-and-dirty parser for JSON returned by the Peloton API.
' The approach is a bit brute-force, and only curiously sophisticated, but it works fine.
' Were the calling application more intense, and expected to run more frequently,
' a more polished implementation would seem appropriate.  Again, for processing about
' one Peloton ride of data per day, it works, and is well with speed expectations.

Private p&, token, dic

Public Function JSONparser(JSON$) As Object
  p = 1
  token = Tokenize(JSON)
  Set dic = CreateObject("Scripting.Dictionary")
  ParseObj "obj"
  Set JSONparser = dic
End Function

Private Function ReducePath$(key$)
  If InStr(key, ".") Then ReducePath = Left(key, InStrRev(key, ".") - 1) Else ReducePath = key
End Function

Private Function ParseObj(key$)
  Do: p = p + 1
    Select Case token(p)
      Case "]"
      Case "[":  ParseArr key
      Case "{"
        If token(p + 1) = "}" Then
          p = p + 1
          dic.Add key, "null"
        Else
          ParseObj key
        End If
      Case "}":  key = ReducePath(key): Exit Do
      Case ":":  key = key & "." & token(p - 1)
      Case ",":  key = ReducePath(key)
      Case Else: If token(p + 1) <> ":" Then dic.Add key, token(p)
    End Select
  Loop
End Function

Private Function ParseArr(key$)
  Dim e&
  Do: p = p + 1
    Select Case token(p)
      Case "}"
      Case "{":  ParseObj key & "(" & e & ")"
      Case "[":  ParseArr key
      Case "]":  Exit Do
      Case ":":  key = key & "(" & e & ")"
      Case ",":  e = e + 1
      Case Else: dic.Add key & "(" & e & ")", token(p)
    End Select
  Loop
End Function

Private Function Tokenize(s$)
  Const Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
  Tokenize = RExtract(s, Pattern, True)
End Function

Private Function RExtract(s$, Pattern, Optional bGroup1Bias As Boolean, Optional bGlobal As Boolean = True)
  Dim c&, m, n, v
  With CreateObject("vbscript.regexp")
    .Global = bGlobal
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = Pattern
    If .TEST(s) Then
      Set m = .Execute(s)
      ReDim v(1 To m.Count)
      For Each n In m
        c = c + 1
        v(c) = n.Value
        If bGroup1Bias Then If Len(n.submatches(0)) Or n.Value = """""" Then v(c) = n.submatches(0)
      Next
    End If
  End With
  RExtract = v
End Function

Public Function GetFilteredIntegers(dic, match)
  Dim c As Long, i As Long, n As Long
  Dim v: v = dic.Keys: n = UBound(v)
  Dim w() As Integer: ReDim w(1 To dic.Count)
  For i = 0 To n
    If v(i) Like match Then
      c = c + 1
      w(c) = CInt(dic(v(i)))
    End If
  Next i
  If c > 0 Then
    ReDim Preserve w(1 To c)
  Else
    ReDim w(0)
  End If
  GetFilteredIntegers = w
End Function

Public Function GetFilteredDoubles(dic, match)
  Dim c As Long, i As Long, n As Long
  Dim v: v = dic.Keys: n = UBound(v)
  Dim w() As Double: ReDim w(1 To dic.Count)
  For i = 0 To n
    If v(i) Like match Then
      c = c + 1
      w(c) = CDbl(dic(v(i)))
    End If
  Next i
  If c > 0 Then
    ReDim Preserve w(1 To c)
  Else
    ReDim w(0)
  End If
  GetFilteredDoubles = w
End Function
