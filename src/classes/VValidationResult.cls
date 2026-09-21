VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VValidationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mInitialized As Boolean
Private mSuccess As Boolean
Private mValue As Variant
Private mIssues As Collection
Private mErrorText As String

Public Function InternalInitialize(ByVal IsSuccess As Boolean, ByVal InputValue As Variant, ByVal SourceIssues As Collection) As VValidationResult
    If mInitialized Then
        Err.Raise vbObjectError + 2102, "VValidationResult.InternalInitialize", "Validation result has already been initialized."
    End If
    If SourceIssues Is Nothing Then
        Err.Raise vbObjectError + 2299, "VValidationResult.InternalInitialize", "Validation issues collection is required."
    End If

    mSuccess = IsSuccess
    If IsSuccess Then
        If IsObject(InputValue) Then
            Set mValue = InputValue
        Else
            mValue = InputValue
        End If
    Else
        mValue = Empty
    End If

    Set mIssues = CloneIssues(SourceIssues)
    mErrorText = BuildErrorText(mIssues)
    mInitialized = True
    Set InternalInitialize = Me
End Function

Public Property Get Success() As Boolean
    EnsureInitialized "VValidationResult.Success"
    Success = mSuccess
End Property

Public Property Get Value() As Variant
    EnsureInitialized "VValidationResult.Value"
    If IsObject(mValue) Then
        Set Value = mValue
    Else
        Value = mValue
    End If
End Property

Public Property Get Issues() As Collection
    EnsureInitialized "VValidationResult.Issues"
    Set Issues = CloneIssues(mIssues)
End Property

Public Property Get ErrorText() As String
    EnsureInitialized "VValidationResult.ErrorText"
    ErrorText = mErrorText
End Property

Private Function CloneIssues(ByVal SourceIssues As Collection) As Collection
    If SourceIssues Is Nothing Then
        Err.Raise vbObjectError + 2299, "VValidationResult.CloneIssues", "Validation issues collection is required."
    End If

    Dim copiedIssues As Collection
    Set copiedIssues = New Collection

    Dim index As Long
    For index = 1 To SourceIssues.Count
        If Not IsObject(SourceIssues.Item(index)) Then
            Err.Raise vbObjectError + 2299, "VValidationResult.CloneIssues", "Issue entries must be Dictionary objects."
        End If

        Dim sourceIssue As Object
        Set sourceIssue = SourceIssues.Item(index)
        copiedIssues.Add CloneIssue(sourceIssue)
    Next index

    Set CloneIssues = copiedIssues
End Function

Private Function CloneIssue(ByVal SourceIssue As Object) As Object
    Dim copiedIssue As Object
    Set copiedIssue = NewIssueDictionary("VValidationResult.CloneIssue")

    copiedIssue.Add "path", ReadIssueString(SourceIssue, "path")
    copiedIssue.Add "code", ReadIssueString(SourceIssue, "code")
    copiedIssue.Add "message", ReadIssueString(SourceIssue, "message")
    copiedIssue.Add "expected", ReadIssueString(SourceIssue, "expected")
    copiedIssue.Add "received", ReadIssueString(SourceIssue, "received")
    Set CloneIssue = copiedIssue
End Function

Private Function ReadIssueString(ByVal SourceIssue As Object, ByVal Key As String) As String
    On Error GoTo InvalidIssue
    If SourceIssue Is Nothing Then
        GoTo InvalidIssue
    End If
    If Not SourceIssue.Exists(Key) Then
        GoTo InvalidIssue
    End If

    Dim value As Variant
    value = SourceIssue.Item(Key)
    If IsObject(value) Or IsArray(value) Or IsNull(value) Then
        GoTo InvalidIssue
    End If

    ReadIssueString = CStr(value)
    Exit Function

InvalidIssue:
    Err.Raise vbObjectError + 2299, "VValidationResult.ReadIssueString", "Issue dictionary is missing a valid '" & Key & "' value."
End Function

Private Function BuildErrorText(ByVal Issues As Collection) As String
    Dim lines As String
    Dim index As Long
    For index = 1 To Issues.Count
        Dim issue As Object
        Set issue = Issues.Item(index)

        Dim line As String
        line = CStr(issue.Item("path")) & ": " & CStr(issue.Item("message")) & " (expected=" & CStr(issue.Item("expected")) & ", received=" & CStr(issue.Item("received")) & ")"
        If Len(lines) > 0 Then
            lines = lines & vbCrLf
        End If
        lines = lines & line
    Next index
    BuildErrorText = lines
End Function

Private Function NewIssueDictionary(ByVal Source As String) As Object
    On Error GoTo DictionaryUnavailable

    Dim issue As Object
    Set issue = CreateObject("Scripting.Dictionary")
    issue.CompareMode = 0
    Set NewIssueDictionary = issue
    Exit Function

DictionaryUnavailable:
    Err.Raise vbObjectError + 2200, Source, "Scripting.Dictionary is required for validation issues."
End Function

Private Sub EnsureInitialized(ByVal ProcedureName As String)
    If Not mInitialized Then
        Err.Raise vbObjectError + 2102, ProcedureName, "Validation result has not been initialized."
    End If
End Sub
