VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtwRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Property-----------------------------------
Public Matches As Collection
Private reg As Object

Property Let Global2(x As Boolean)
    reg.Global = x
End Property

Property Let IgnoreCase(x As Boolean)
    reg.IgnoreCase = x
End Property

Property Let Multiline(x As Boolean)
    reg.Multiline = x
End Property

Property Let Pattern(ptn As String)
    reg.Pattern = ptn
End Property

Property Get MatchFirstIndex(matchIndex As Long) As Long
    MatchFirstIndex = -1
    If Matches.Count > 0 Then
        MatchFirstIndex = Matches(matchIndex).FirstIndex
    End If
End Property

Property Get MatchLength(matchIndex As Long) As Long
    MatchLength = 0
    If Matches.Count > 0 Then
        MatchLength = Matches(matchIndex).Length
    End If
End Property

Property Get MatchValue(matchIndex As Long) As String
    MatchValue = ""
    If Matches.Count > 0 Then
        MatchValue = Matches(matchIndex).Value
    End If
End Property

' indexes start from 1
Property Get Submatch(matchIndex As Long, submatchIndex As Long) As String
    Submatch = ""
    If Matches.Count > 0 Then
        Submatch = Matches(matchIndex).SubMatches(submatchIndex - 1)
    End If
End Property

'Method--------------------------------------
Private Sub Class_Initialize()
    Set reg = CreateObject("VBScript.RegExp")
    'Debug.Print "initalized"
End Sub

Private Sub Class_Terminate()
    'Debug.Print "terminated"
End Sub

Sub Execute(sourceString As String)
    Set Matches = New Collection
    
    Dim result As Object
    Dim results As Variant
    Set results = reg.Execute(sourceString)

    For Each result In results
        Matches.Add result
    Next result
End Sub

Function Replace(sourceString As String, replaceString As String) As String
    Replace = reg.Replace(sourceString, replaceString)
End Function

Function Test(sourceString As String) As Boolean
    Test = reg.Test(sourceString)
End Function

