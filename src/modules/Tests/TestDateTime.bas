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

Public Sub Test_DateTime_RejectsNullEmptyAndErrorVariant()
    Dim nullValue As Variant
    nullValue = Null
    Dim emptyValue As Variant
    emptyValue = Empty
    Dim errorValue As Variant
    errorValue = CVErr(2042)

    ' Special Variant states are classified before the vbDate check;
    ' none of them is implicitly converted to Date.
    Dim result As VValidationResult
    Set result = Schema.DateTime().SafeParse(nullValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "required", IssueValue(result, "code")

    Set result = Schema.DateTime().SafeParse(emptyValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Set result = Schema.DateTime().SafeParse(errorValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")
    XlflowAssert.AssertStrictEquals "Error(2042)", IssueValue(result, "received")
End Sub

Public Sub Test_DateTime_RejectsDateStringsRegardlessOfLocale()
    ' Strict schemas never parse localized date text, so these strings
    ' are invalid_type on every machine locale.
    Dim result As VValidationResult
    Set result = Schema.DateTime().SafeParse("2024-05-01")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Set result = Schema.DateTime().SafeParse("05/01/2024")
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
