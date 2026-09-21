Attribute VB_Name = "TestNullable"
Option Explicit

Public Sub Test_Nullable_AllowsNullOnlyWhenConfigured()
    Dim inputValue As Variant
    inputValue = Null

    Dim result As VValidationResult
    Set result = Schema.Text().SafeParse(inputValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "required", IssueValue(result, "code")

    Set result = Schema.Text().Nullable().SafeParse(inputValue)
    XlflowAssert.AssertTrue result.Success
    XlflowAssert.AssertNull result.Value
End Sub

Public Sub Test_OptionalField_DoesNotMakeRootEmptyValid()
    Dim inputValue As Variant
    inputValue = Empty

    Dim result As VValidationResult
    Set result = Schema.Text().OptionalField().SafeParse(inputValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")
End Sub

Private Function IssueValue(ByVal result As VValidationResult, ByVal Key As String) As String
    Dim issues As Collection
    Set issues = result.Issues

    Dim issue As Object
    Set issue = issues.Item(1)
    IssueValue = CStr(issue.Item(Key))
End Function
