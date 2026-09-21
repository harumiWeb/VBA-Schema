Attribute VB_Name = "TestPattern"
Option Explicit

Public Sub Test_Pattern_UsesCaseSensitiveFullMatch()
    Dim patternSchema As VSchema
    Set patternSchema = Schema.Text().Pattern("[A-Z]{3}-[0-9]{4}")

    Dim result As VValidationResult
    Set result = patternSchema.SafeParse("ABC-1234")
    XlflowAssert.AssertTrue result.Success

    Set result = patternSchema.SafeParse("xABC-1234")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_pattern", result.Issues.Item(1).Item("code")

    Set result = patternSchema.SafeParse("abc-1234")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_pattern", result.Issues.Item(1).Item("code")
End Sub

Public Sub Test_Pattern_InvalidExpressionRaisesProgrammerError()
    Dim patternSchema As VSchema
    Set patternSchema = Schema.Text().Pattern("[")

    On Error Resume Next
    Err.Clear
    Call patternSchema.SafeParse("x")
    Dim errorNumber As Long
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0

    XlflowAssert.AssertEquals vbObjectError + 2100, errorNumber
End Sub

Public Sub Test_Email_UsesPracticalAsciiForm()
    Dim emailSchema As VSchema
    Set emailSchema = Schema.Text().Email()

    Dim result As VValidationResult
    Set result = emailSchema.SafeParse("alice@example.com")
    XlflowAssert.AssertTrue result.Success

    Set result = emailSchema.SafeParse("first.last+tag@example.co.jp")
    XlflowAssert.AssertTrue result.Success

    Set result = emailSchema.SafeParse("aliceexample.com")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_email", result.Issues.Item(1).Item("code")

    Set result = emailSchema.SafeParse("alice@@example.com")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_email", result.Issues.Item(1).Item("code")

    Set result = emailSchema.SafeParse("alice @example.com")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_email", result.Issues.Item(1).Item("code")

    Set result = emailSchema.SafeParse("")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_email", result.Issues.Item(1).Item("code")

    Set result = emailSchema.SafeParse("alice")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_email", result.Issues.Item(1).Item("code")
End Sub

Public Sub Test_Email_RejectsUnicodeAndExcessiveLength()
    Dim emailSchema As VSchema
    Set emailSchema = Schema.Text().Email()

    Dim result As VValidationResult
    Set result = emailSchema.SafeParse("alice@例.example")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_email", result.Issues.Item(1).Item("code")

    Dim longAddress As String
    longAddress = String$(250, "a") & "@x.com"
    Set result = emailSchema.SafeParse(longAddress)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_email", result.Issues.Item(1).Item("code")
End Sub

Public Sub Test_PatternAndEmail_UseFixedEvaluationOrder()
    Dim compositionSchema As VSchema
    Set compositionSchema = Schema.Text().Email().Pattern(".*")

    Dim result As VValidationResult
    Set result = compositionSchema.SafeParse("not-an-email")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_email", result.Issues.Item(1).Item("code")
End Sub
