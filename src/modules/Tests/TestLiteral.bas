Attribute VB_Name = "TestLiteral"
Option Explicit

Public Sub Test_Literal_UsesScalarCategoryComparison()
    Dim textSchema As VSchema
    Set textSchema = Schema.Literal("1")

    Dim result As VValidationResult
    Set result = textSchema.SafeParse("1")
    XlflowAssert.AssertTrue result.Success

    Set result = textSchema.SafeParse(1)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_literal", result.Issues.Item(1).Item("code")

    Dim numberSchema As VSchema
    Set numberSchema = Schema.Literal(1)
    Set result = numberSchema.SafeParse(CByte(1))
    XlflowAssert.AssertTrue result.Success

    Set result = numberSchema.SafeParse(True)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_literal", result.Issues.Item(1).Item("code")

    Dim binaryStringSchema As VSchema
    Set binaryStringSchema = Schema.Literal("A")
    Set result = binaryStringSchema.SafeParse("a")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_literal", result.Issues.Item(1).Item("code")
End Sub

Public Sub Test_Literal_SupportsNullEmptyAndDateCategories()
    Dim nullValue As Variant
    nullValue = Null
    Dim nullSchema As VSchema
    Set nullSchema = Schema.Literal(nullValue)
    Dim result As VValidationResult
    Set result = nullSchema.SafeParse(Null)
    XlflowAssert.AssertTrue result.Success

    Dim emptyValue As Variant
    emptyValue = Empty
    Dim emptySchema As VSchema
    Set emptySchema = Schema.Literal(emptyValue)
    Set result = emptySchema.SafeParse(Empty)
    XlflowAssert.AssertTrue result.Success

    Dim expectedDate As Date
    expectedDate = DateSerial(2026, 9, 21) + TimeSerial(12, 34, 56)
    Dim dateSchema As VSchema
    Set dateSchema = Schema.Literal(expectedDate)
    Set result = dateSchema.SafeParse(expectedDate)
    XlflowAssert.AssertTrue result.Success

    Set result = dateSchema.SafeParse(CDbl(expectedDate))
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_literal", result.Issues.Item(1).Item("code")

    Set result = nullSchema.SafeParse(Empty)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_literal", result.Issues.Item(1).Item("code")

    Set result = emptySchema.SafeParse(Null)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_literal", result.Issues.Item(1).Item("code")
End Sub

Public Sub Test_Enum_SnapshotsTypedCandidates()
    Dim candidates(0 To 2) As Variant
    candidates(0) = "pending"
    candidates(1) = 1
    candidates(2) = Null

    Dim enumSchema As VSchema
    Set enumSchema = Schema.EnumOf(candidates)

    candidates(0) = "changed"
    candidates(1) = 2

    Dim result As VValidationResult
    Set result = enumSchema.SafeParse("pending")
    XlflowAssert.AssertTrue result.Success

    Set result = enumSchema.SafeParse(1#)
    XlflowAssert.AssertTrue result.Success

    Set result = enumSchema.SafeParse(Null)
    XlflowAssert.AssertTrue result.Success

    Set result = enumSchema.SafeParse("changed")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_enum", result.Issues.Item(1).Item("code")
End Sub

Public Sub Test_Enum_AcceptsTypedScalarArray()
    Dim candidates(0 To 1) As String
    candidates(0) = "alpha"
    candidates(1) = "beta"

    Dim enumSchema As VSchema
    Set enumSchema = Schema.EnumOf(candidates)
    Dim result As VValidationResult
    Set result = enumSchema.SafeParse("beta")
    XlflowAssert.AssertTrue result.Success
End Sub
