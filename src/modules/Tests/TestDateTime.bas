Attribute VB_Name = "TestDateTime"
Option Explicit

Public Sub Test_DateTime_AcceptsOnlyDateVariant()
    Dim result As VValidationResult
    Set result = Schema.DateTime().SafeParse(CDate("2026-09-21 12:34:56"))
    XlflowAssert.AssertTrue result.Success

    Set result = Schema.DateTime().SafeParse("2026-09-21")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")
    XlflowAssert.AssertStrictEquals "String(""2026-09-21"")", IssueValue(result, "received")
End Sub

Private Function IssueValue(ByVal result As VValidationResult, ByVal Key As String) As String
    Dim issues As Collection
    Set issues = result.Issues

    Dim issue As Object
    Set issue = issues.Item(1)
    IssueValue = CStr(issue.Item(Key))
End Function
