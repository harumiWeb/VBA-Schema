Attribute VB_Name = "TestUnion"
Option Explicit

Public Sub Test_Union_FirstAndLaterBranchSuccess()
    Dim branches As Collection
    Set branches = New Collection
    branches.Add Schema.Text()
    branches.Add Schema.Number()

    Dim unionSchema As VSchema
    Set unionSchema = Schema.UnionOf(branches)

    Dim result As VValidationResult
    Set result = unionSchema.SafeParse("alpha")
    XlflowAssert.AssertTrue result.Success

    Set result = unionSchema.SafeParse(42)
    XlflowAssert.AssertTrue result.Success
End Sub

Public Sub Test_Union_AllBranchesFailWithSingleIssue()
    Dim branches As Collection
    Set branches = New Collection
    branches.Add Schema.Text()
    branches.Add Schema.Number()

    Dim unionSchema As VSchema
    Set unionSchema = Schema.UnionOf(branches)

    Dim result As VValidationResult
    Set result = unionSchema.SafeParse(True)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertStrictEquals "invalid_union", result.Issues.Item(1).Item("code")
    XlflowAssert.AssertStrictEquals "$", result.Issues.Item(1).Item("path")
End Sub

Public Sub Test_Union_SnapshotsCollectionAndSharesSchemaReferences()
    Dim branches As Collection
    Set branches = New Collection

    Dim textSchema As VSchema
    Set textSchema = Schema.Text()
    branches.Add textSchema

    Dim unionSchema As VSchema
    Set unionSchema = Schema.UnionOf(branches)

    branches.Add Schema.Bool()

    Dim result As VValidationResult
    Set result = unionSchema.SafeParse(True)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_union", result.Issues.Item(1).Item("code")

    Call textSchema.Length(2)
    Set result = unionSchema.SafeParse("ok")
    XlflowAssert.AssertTrue result.Success

    Set result = unionSchema.SafeParse("x")
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_union", result.Issues.Item(1).Item("code")
End Sub

Public Sub Test_Union_PreservesNestedComposition()
    Dim innerBranches As Collection
    Set innerBranches = New Collection
    innerBranches.Add Schema.Text()
    innerBranches.Add Schema.Number()

    Dim innerSchema As VSchema
    Set innerSchema = Schema.UnionOf(innerBranches)

    Dim outerBranches As Collection
    Set outerBranches = New Collection
    outerBranches.Add innerSchema
    outerBranches.Add Schema.Bool()

    Dim outerSchema As VSchema
    Set outerSchema = Schema.UnionOf(outerBranches)

    Dim result As VValidationResult
    Set result = outerSchema.SafeParse(100)
    XlflowAssert.AssertTrue result.Success

    Set result = outerSchema.SafeParse(True)
    XlflowAssert.AssertTrue result.Success

    Set result = outerSchema.SafeParse(DateSerial(2026, 9, 21))
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertStrictEquals "invalid_union", result.Issues.Item(1).Item("code")
End Sub

Public Sub Test_Union_PropagatesBranchProgrammerError()
    Dim branches As Collection
    Set branches = New Collection
    branches.Add Schema.Text().Pattern("[")
    branches.Add Schema.Number()

    Dim unionSchema As VSchema
    Set unionSchema = Schema.UnionOf(branches)

    On Error Resume Next
    Err.Clear
    Call unionSchema.SafeParse("x")
    Dim errorNumber As Long
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0

    XlflowAssert.AssertEquals vbObjectError + 2100, errorNumber
End Sub

Public Sub Test_Union_SupportsNullAndEmptyLiteralBranches()
    Dim nullValue As Variant
    nullValue = Null
    Dim emptyValue As Variant
    emptyValue = Empty

    Dim branches As Collection
    Set branches = New Collection
    branches.Add Schema.Literal(nullValue)
    branches.Add Schema.Literal(emptyValue)

    Dim unionSchema As VSchema
    Set unionSchema = Schema.UnionOf(branches)

    Dim result As VValidationResult
    Set result = unionSchema.SafeParse(Null)
    XlflowAssert.AssertTrue result.Success

    Set result = unionSchema.SafeParse(Empty)
    XlflowAssert.AssertTrue result.Success
End Sub

Public Sub Test_Union_ParticipatesInSchemaCyclePreflight()
    Dim objectSchema As VSchema
    Set objectSchema = Schema.ObjectSchema()

    Dim branches As Collection
    Set branches = New Collection
    branches.Add objectSchema

    Dim unionSchema As VSchema
    Set unionSchema = Schema.UnionOf(branches)
    Call objectSchema.Field("next", unionSchema)

    Dim inputObject As Object
    Set inputObject = CreateObject("Scripting.Dictionary")
    inputObject.Add "next", inputObject

    On Error Resume Next
    Err.Clear
    Call objectSchema.SafeParse(inputObject)
    Dim errorNumber As Long
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0

    XlflowAssert.AssertEquals vbObjectError + 2103, errorNumber
End Sub
