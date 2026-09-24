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

Public Sub Test_Bool_RejectsNullEmptyAndErrorVariant()
    Dim nullValue As Variant
    nullValue = Null
    Dim emptyValue As Variant
    emptyValue = Empty
    Dim errorValue As Variant
    errorValue = CVErr(2042)

    ' Special Variant states are classified before the vbBoolean check;
    ' none of them is implicitly converted to Boolean.
    Dim result As VValidationResult
    Set result = Schema.Bool().SafeParse(nullValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "required", IssueValue(result, "code")

    Set result = Schema.Bool().SafeParse(emptyValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Set result = Schema.Bool().SafeParse(errorValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")
    XlflowAssert.AssertStrictEquals "Error(2042)", IssueValue(result, "received")
End Sub

Private Function IssueValue(ByVal result As VValidationResult, ByVal Key As String) As String
    Dim issues As Collection
    Set issues = result.Issues

    Dim issue As Object
    Set issue = issues.Item(1)
    IssueValue = CStr(issue.Item(Key))
End Function
