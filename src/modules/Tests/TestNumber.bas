Attribute VB_Name = "TestNumber"
Option Explicit

Public Sub Test_Number_AcceptsSupportedNumericSubtypes()
    Dim result As VValidationResult

    Set result = Schema.Number().SafeParse(CByte(1))
    XlflowAssert.AssertTrue result.Success
    Set result = Schema.Number().SafeParse(CInt(2))
    XlflowAssert.AssertTrue result.Success
    Set result = Schema.Number().SafeParse(CLng(3))
    XlflowAssert.AssertTrue result.Success
    Set result = Schema.Number().SafeParse(CSng(4.5))
    XlflowAssert.AssertTrue result.Success
    Set result = Schema.Number().SafeParse(CDbl(5.5))
    XlflowAssert.AssertTrue result.Success
    Set result = Schema.Number().SafeParse(CCur(6.5))
    XlflowAssert.AssertTrue result.Success
End Sub

Public Sub Test_Number_RejectsImplicitConversions()
    Dim result As VValidationResult
    Set result = Schema.Number().SafeParse("10")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Set result = Schema.Number().SafeParse(True)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Dim decimalValue As Variant
    decimalValue = CDec(10)
    Set result = Schema.Number().SafeParse(decimalValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")
End Sub

Public Sub Test_Number_EnforcesBoundsAndWholeNumber()
    Dim result As VValidationResult
    Set result = Schema.Number().Min(3).Max(5).SafeParse(2)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "too_small", IssueValue(result, "code")

    Set result = Schema.Number().Min(3).Max(5).SafeParse(6)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "too_big", IssueValue(result, "code")

    Set result = Schema.Number().WholeNumber().SafeParse(3.25)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_integer", IssueValue(result, "code")

    Set result = Schema.Number().WholeNumber().SafeParse(3#)
    XlflowAssert.AssertTrue result.Success
End Sub

Public Sub Test_Number_AcceptsFiniteDoubleOutsideDecimalRange()
    Dim largeValue As Double
    largeValue = 1E+100

    Dim result As VValidationResult
    Set result = Schema.Number().SafeParse(largeValue)
    XlflowAssert.AssertTrue result.Success

    Set result = Schema.Number().Min(1E+99).Max(1E+101).WholeNumber().SafeParse(largeValue)
    XlflowAssert.AssertTrue result.Success

    Set result = Schema.Number().Max(1E+99).SafeParse(largeValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "too_big", IssueValue(result, "code")
End Sub

Public Sub Test_Number_ConstraintsRejectNullBeforeComparison()
    Dim nullValue As Variant
    nullValue = Null

    ' Null must be classified before any numeric boundary is reached;
    ' in VBA "Null > boundary" evaluates to Null, not False.
    Dim result As VValidationResult
    Set result = Schema.Number().Min(0).SafeParse(nullValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertStrictEquals "required", IssueValue(result, "code")
    XlflowAssert.AssertStrictEquals "Null", IssueValue(result, "received")

    Set result = Schema.Number().Max(10).SafeParse(nullValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertStrictEquals "required", IssueValue(result, "code")

    Set result = Schema.Number().WholeNumber().SafeParse(nullValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertStrictEquals "required", IssueValue(result, "code")
End Sub

Public Sub Test_Number_ConstraintsRejectEmptyBeforeComparison()
    Dim emptyValue As Variant
    emptyValue = Empty

    ' Empty must be classified before constraint evaluation; VBA would
    ' otherwise coerce it to 0 and let it pass a Min(0) boundary.
    Dim result As VValidationResult
    Set result = Schema.Number().Min(0).SafeParse(emptyValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")
    XlflowAssert.AssertStrictEquals "Empty", IssueValue(result, "received")

    Set result = Schema.Number().Max(10).SafeParse(emptyValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Set result = Schema.Number().WholeNumber().SafeParse(emptyValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")
End Sub

Public Sub Test_Number_ConstraintsRejectErrorVariantBeforeComparison()
    Dim errorValue As Variant
    errorValue = CVErr(2042)

    ' An Error Variant must be classified before any conversion or
    ' comparison; CDec(CVErr(...)) would raise a runtime error.
    Dim result As VValidationResult
    Set result = Schema.Number().Min(0).SafeParse(errorValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")
    XlflowAssert.AssertStrictEquals "Error(2042)", IssueValue(result, "received")

    Set result = Schema.Number().Max(10).SafeParse(errorValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Set result = Schema.Number().WholeNumber().SafeParse(errorValue)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")
End Sub

Public Sub Test_Number_RejectsNumericStringsRegardlessOfLocale()
    ' Strict schemas never parse localized numeric text, so these strings
    ' are invalid_type on every machine locale.
    Dim result As VValidationResult
    Set result = Schema.Number().SafeParse("1.5")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", IssueValue(result, "code")

    Set result = Schema.Number().SafeParse("1,5")
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
