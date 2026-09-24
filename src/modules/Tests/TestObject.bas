Attribute VB_Name = "TestObject"
Option Explicit

Public Sub Test_Object_ValidatesRequiredAndOptionalFields()
    Dim objectSchema As VSchema
    Set objectSchema = Schema.ObjectSchema() _
        .Field("name", Schema.Text()) _
        .Field("age", Schema.Number().OptionalField())

    Dim inputObject As Object
    Set inputObject = CreateObject("Scripting.Dictionary")
    If inputObject Is Nothing Then
        Err.Raise vbObjectError + 2299, "TestObject", "Input dictionary is required."
    End If
    inputObject.Add "name", "Ada"

    Dim result As VValidationResult
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertTrue result.Success

    If inputObject Is Nothing Then
        Err.Raise vbObjectError + 2299, "TestObject", "Input dictionary is required."
    End If
    inputObject.Add "age", Empty
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    AssertIssue result.Issues.Item(1), "$.age", "invalid_type"

    If inputObject Is Nothing Then
        Err.Raise vbObjectError + 2299, "TestObject", "Input dictionary is required."
    End If
    inputObject.Remove "name"
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 2, result.Issues.Count
    AssertIssue result.Issues.Item(1), "$.name", "required"
    AssertIssue result.Issues.Item(2), "$.age", "invalid_type"
End Sub

Public Sub Test_Object_ValidatesNestedPathsAndPreservesInput()
    Dim addressSchema As VSchema
    Set addressSchema = Schema.ObjectSchema().Field("zip.code", Schema.Text().Length(5))

    Dim objectSchema As VSchema
    Set objectSchema = Schema.ObjectSchema().Field("user", addressSchema)

    Dim address As Object
    Set address = CreateObject("Scripting.Dictionary")
    address.Add "zip.code", "12"

    Dim inputObject As Object
    Set inputObject = CreateObject("Scripting.Dictionary")
    If inputObject Is Nothing Then
        Err.Raise vbObjectError + 2299, "TestObject", "Input dictionary is required."
    End If
    inputObject.Add "user", address
    inputObject.Add "unchanged", 7

    Dim result As VValidationResult
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    AssertIssue result.Issues.Item(1), "$.user[" & Chr$(34) & "zip.code" & Chr$(34) & "]", "invalid_length"
    If inputObject Is Nothing Then
        Err.Raise vbObjectError + 2299, "TestObject", "Input dictionary is required."
    End If
    XlflowAssert.AssertEquals 2, inputObject.Count
    XlflowAssert.AssertTrue inputObject.Exists("unchanged")
    XlflowAssert.AssertTrue address.Exists("zip.code")
End Sub

Public Sub Test_Object_UsesBinaryFieldMatchingAndStrictUnknownOrdering()
    Dim objectSchema As VSchema
    Set objectSchema = Schema.ObjectSchema().Field("Name", Schema.Text()).Strict()

    Dim inputObject As Object
    Set inputObject = CreateObject("Scripting.Dictionary")
    inputObject.CompareMode = vbTextCompare
    inputObject.Add "name", "lowercase"

    Dim result As VValidationResult
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    AssertIssue result.Issues.Item(1), "$.Name", "required"
    AssertIssue result.Issues.Item(2), "$.name", "unknown_field"

    Set inputObject = CreateObject("Scripting.Dictionary")
    inputObject.CompareMode = vbBinaryCompare
    inputObject.Add "z", 1
    inputObject.Add "A", 2
    inputObject.Add "a", 3

    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 4, result.Issues.Count

    Dim issues As Collection
    Set issues = result.Issues
    AssertIssue issues.Item(1), "$.Name", "required"
    AssertIssue issues.Item(2), "$.A", "unknown_field"
    AssertIssue issues.Item(3), "$.a", "unknown_field"
    AssertIssue issues.Item(4), "$.z", "unknown_field"
End Sub

Public Sub Test_Object_RejectsInvalidInputAndNonStringKeys()
    Dim objectSchema As VSchema
    Set objectSchema = Schema.ObjectSchema()

    Dim collectionValue As Collection
    Set collectionValue = New Collection
    collectionValue.Add "item"

    Dim result As VValidationResult
    Set result = objectSchema.SafeParse(collectionValue)
    XlflowAssert.AssertFalse result.Success
    AssertIssue result.Issues.Item(1), "$", "invalid_type"

    Dim inputObject As Object
    Set inputObject = CreateObject("Scripting.Dictionary")
    inputObject.Add 1, "numeric key"
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    AssertIssue result.Issues.Item(1), "$", "invalid_key"
    XlflowAssert.AssertContains "Number(1)", result.ErrorText
End Sub

Public Sub Test_Object_DistinguishesMissingNullAndEmpty()
    Dim objectSchema As VSchema
    Set objectSchema = Schema.ObjectSchema() _
        .Field("required", Schema.Text()) _
        .Field("nullable", Schema.Text().Nullable())

    Dim inputObject As Object
    Set inputObject = CreateObject("Scripting.Dictionary")
    inputObject.Add "required", "present"
    inputObject.Add "nullable", Null

    Dim result As VValidationResult
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertTrue result.Success

    If inputObject Is Nothing Then
        Err.Raise vbObjectError + 2299, "TestObject", "Input dictionary is required."
    End If
    inputObject.Remove "nullable"
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    AssertIssue result.Issues.Item(1), "$.nullable", "required"

    If inputObject Is Nothing Then
        Err.Raise vbObjectError + 2299, "TestObject", "Input dictionary is required."
    End If
    inputObject.Add "nullable", Empty
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    AssertIssue result.Issues.Item(1), "$.nullable", "invalid_type"
End Sub

Public Sub Test_Object_UsesDeclarationOrderAndPassthrough()
    Dim objectSchema As VSchema
    Set objectSchema = Schema.ObjectSchema() _
        .Field("first", Schema.Text()) _
        .Field("second", Schema.Text())

    Dim inputObject As Object
    Set inputObject = CreateObject("Scripting.Dictionary")
    inputObject.Add "second", 2
    inputObject.Add "first", 1
    inputObject.Add "extra", "passthrough"

    Dim result As VValidationResult
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 2, result.Issues.Count
    AssertIssue result.Issues.Item(1), "$.first", "invalid_type"
    AssertIssue result.Issues.Item(2), "$.second", "invalid_type"

    Set inputObject = CreateObject("Scripting.Dictionary")
    inputObject.Add "first", "one"
    inputObject.Add "extra", "passthrough"
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    AssertIssue result.Issues.Item(1), "$.second", "required"
End Sub

Public Sub Test_Object_RejectsNothingAndEscapesSpecialPath()
    Dim objectSchema As VSchema
    Set objectSchema = Schema.ObjectSchema()

    Dim nothingObject As Object
    Set nothingObject = Nothing
    Dim result As VValidationResult
    Set result = objectSchema.SafeParse(nothingObject)
    XlflowAssert.AssertFalse result.Success
    AssertIssue result.Issues.Item(1), "$", "invalid_type"

    Dim specialName As String
    specialName = "a.b[" & Chr$(34) & Chr$(92) & vbCr & vbLf & vbTab & "]"
    Set objectSchema = Schema.ObjectSchema().Field(specialName, Schema.Text())
    Dim inputObject As Object
    Set inputObject = CreateObject("Scripting.Dictionary")
    inputObject.Add specialName, 1
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success

    Dim expectedPath As String
    expectedPath = "$[" & Chr$(34) & "a.b[" & Chr$(92) & Chr$(34) & Chr$(92) & Chr$(92) & Chr$(92) & "r" & Chr$(92) & "n" & Chr$(92) & "t]" & Chr$(34) & "]"
    AssertIssue result.Issues.Item(1), expectedPath, "invalid_type"
End Sub

Public Sub Test_Object_RejectsPresentErrorVariantField()
    Dim objectSchema As VSchema
    Set objectSchema = Schema.ObjectSchema() _
        .Field("field", Schema.Text())

    Dim errorValue As Variant
    errorValue = CVErr(2042)

    Dim inputObject As Object
    Set inputObject = CreateObject("Scripting.Dictionary")
    inputObject.Add "field", errorValue

    ' A present Error Variant is a distinct state from a missing field
    ' and must reach the child schema as invalid_type, not required.
    Dim result As VValidationResult
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    AssertIssue result.Issues.Item(1), "$.field", "invalid_type"
    XlflowAssert.AssertStrictEquals "Error(2042)", result.Issues.Item(1).Item("received")
End Sub

Public Sub Test_Object_OptionalFieldRejectsPresentSpecialStates()
    Dim objectSchema As VSchema
    Set objectSchema = Schema.ObjectSchema() _
        .Field("field", Schema.Text().OptionalField())

    Dim inputObject As Object
    Set inputObject = CreateObject("Scripting.Dictionary")

    ' Field absence is the only state OptionalField allows.
    Dim result As VValidationResult
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertTrue result.Success

    Dim emptyValue As Variant
    emptyValue = Empty
    inputObject.Add "field", emptyValue
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    AssertIssue result.Issues.Item(1), "$.field", "invalid_type"

    inputObject.Remove "field"
    inputObject.Add "field", Null
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    AssertIssue result.Issues.Item(1), "$.field", "required"

    Dim errorValue As Variant
    errorValue = CVErr(2042)
    inputObject.Remove "field"
    inputObject.Add "field", errorValue
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    AssertIssue result.Issues.Item(1), "$.field", "invalid_type"
End Sub

Public Sub Test_Object_NullableFieldRejectsPresentEmptyAndError()
    Dim objectSchema As VSchema
    Set objectSchema = Schema.ObjectSchema() _
        .Field("field", Schema.Text().Nullable())

    Dim inputObject As Object
    Set inputObject = CreateObject("Scripting.Dictionary")
    inputObject.Add "field", Null

    ' Nullable allows a present Null only; it does not cover absence.
    Dim result As VValidationResult
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertTrue result.Success

    Dim emptyValue As Variant
    emptyValue = Empty
    inputObject.Remove "field"
    inputObject.Add "field", emptyValue
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    AssertIssue result.Issues.Item(1), "$.field", "invalid_type"

    Dim errorValue As Variant
    errorValue = CVErr(2042)
    inputObject.Remove "field"
    inputObject.Add "field", errorValue
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    AssertIssue result.Issues.Item(1), "$.field", "invalid_type"

    inputObject.Remove "field"
    Set result = objectSchema.SafeParse(inputObject)
    XlflowAssert.AssertFalse result.Success
    AssertIssue result.Issues.Item(1), "$.field", "required"
End Sub

Private Sub AssertIssue(ByVal issue As Object, ByVal expectedPath As String, ByVal expectedCode As String)
    If issue Is Nothing Then
        Err.Raise vbObjectError + 2299, "TestObject.AssertIssue", "Issue is required."
    End If
    XlflowAssert.AssertStrictEquals expectedPath, issue.Item("path")
    XlflowAssert.AssertStrictEquals expectedCode, issue.Item("code")
End Sub
