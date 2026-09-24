Attribute VB_Name = "TestText"
Option Explicit

Public Sub Test_Text_AcceptsOnlyString()
    Dim result As VValidationResult
    Set result = Schema.Text().SafeParse("hello")
    XlflowAssert.AssertTrue result.Success
    XlflowAssert.AssertStrictEquals "hello", result.Value

    Set result = Schema.Text().SafeParse(10)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")
    XlflowAssert.AssertStrictEquals "$", IssueValue(result, "path")
    XlflowAssert.AssertStrictEquals "Number(10)", IssueValue(result, "received")
End Sub

Public Sub Test_Text_EnforcesLengthAndBoundsInOrder()
    Dim result As VValidationResult
    Set result = Schema.Text().Length(3).Min(2).Max(5).SafeParse("x")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_length", IssueValue(result, "code")

    Set result = Schema.Text().Min(2).Max(5).SafeParse("x")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "too_small", IssueValue(result, "code")

    Set result = Schema.Text().Min(2).Max(5).SafeParse("123456")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "too_big", IssueValue(result, "code")

    Set result = Schema.Text().Length(0).SafeParse("")
    XlflowAssert.AssertTrue result.Success, "Length(0) should accept an empty String"
End Sub

Public Sub Test_Text_UsesFixedErrorTextGrammar()
    Dim result As VValidationResult
    Set result = Schema.Text().Min(3).SafeParse("x")

    XlflowAssert.AssertStrictEquals "$: Text length is below the minimum. (expected=Text length >= 3, received=String(""x""))", result.ErrorText
End Sub

Public Sub Test_Text_FormatsErrorVariantReceivedDescriptor()
    Dim inputValue As Variant
    inputValue = CVErr(2042)

    Dim result As VValidationResult
    Set result = Schema.Text().SafeParse(inputValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "Error(2042)", IssueValue(result, "received")
End Sub

Public Sub Test_Text_ConstraintsRejectNullBeforeEvaluation()
    Dim nullValue As Variant
    nullValue = Null

    ' Null must be classified before Length, Min, Max, Pattern, or Email
    ' is evaluated; Len(CStr(Null)) would raise "Invalid use of Null".
    Dim result As VValidationResult
    Set result = Schema.Text().Length(3).SafeParse(nullValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "required", IssueValue(result, "code")

    Set result = Schema.Text().Min(1).SafeParse(nullValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "required", IssueValue(result, "code")

    Set result = Schema.Text().Max(5).SafeParse(nullValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "required", IssueValue(result, "code")

    Set result = Schema.Text().Pattern("x").SafeParse(nullValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "required", IssueValue(result, "code")

    Set result = Schema.Text().Email().SafeParse(nullValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "required", IssueValue(result, "code")
End Sub

Public Sub Test_Text_ConstraintsRejectEmptyBeforeEvaluation()
    Dim emptyValue As Variant
    emptyValue = Empty

    ' Empty must be classified before constraint evaluation; VBA would
    ' otherwise coerce it to an empty String and let it pass Min(0).
    Dim result As VValidationResult
    Set result = Schema.Text().Min(0).SafeParse(emptyValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")
    XlflowAssert.AssertStrictEquals "Empty", IssueValue(result, "received")

    Set result = Schema.Text().Length(0).SafeParse(emptyValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Set result = Schema.Text().Max(5).SafeParse(emptyValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Set result = Schema.Text().Pattern(".*").SafeParse(emptyValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Set result = Schema.Text().Email().SafeParse(emptyValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")
End Sub

Public Sub Test_Text_ConstraintsRejectErrorVariantBeforeEvaluation()
    Dim errorValue As Variant
    errorValue = CVErr(2042)

    ' An Error Variant must be classified before Length, Pattern, or
    ' Email is evaluated; otherwise CStr could turn it into "Error <n>"
    ' and let it pass through constraints as ordinary text.
    Dim result As VValidationResult
    Set result = Schema.Text().Length(3).SafeParse(errorValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Set result = Schema.Text().Min(1).SafeParse(errorValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Set result = Schema.Text().Max(5).SafeParse(errorValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Set result = Schema.Text().Pattern("x").SafeParse(errorValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Set result = Schema.Text().Email().SafeParse(errorValue)
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
