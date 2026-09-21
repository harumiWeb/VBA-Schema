Attribute VB_Name = "TestBool"
Option Explicit

Public Sub Test_Bool_AcceptsOnlyBoolean()
    Dim result As VValidationResult
    Set result = Schema.Bool().SafeParse(True)
    XlflowAssert.AssertTrue result.Success
    XlflowAssert.AssertStrictEquals True, result.Value

    Set result = Schema.Bool().SafeParse(1)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")
    XlflowAssert.AssertStrictEquals "Boolean", IssueValue(result, "expected")
End Sub

Private Function IssueValue(ByVal result As VValidationResult, ByVal Key As String) As String
    Dim issues As Collection
    Set issues = result.Issues

    Dim issue As Object
    Set issue = issues.Item(1)
    IssueValue = CStr(issue.Item(Key))
End Function
