Attribute VB_Name = "TestErrors"
Option Explicit

Public Sub Test_Errors_RejectInvalidConstraintArguments()
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureTextMinError(-1), "Text.Min negative"
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureTextLengthError(1.5), "Text.Length fractional"
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureNumberMinError("3"), "Number.Min string"
End Sub

Public Sub Test_Errors_RejectInvalidConstraintComposition()
    XlflowAssert.AssertEquals vbObjectError + 2101, CaptureTextMaxAfterMinError(), "Max less than Min"
    XlflowAssert.AssertEquals vbObjectError + 2101, CaptureTextMinAfterLengthError(), "Min greater than Length"
    XlflowAssert.AssertEquals vbObjectError + 2101, CaptureDuplicateWholeNumberError(), "duplicate WholeNumber"
    XlflowAssert.AssertEquals vbObjectError + 2101, CaptureInvalidWholeNumberKindError(), "WholeNumber on Text"
End Sub

Public Sub Test_Errors_RejectUninitializedAndDirectInitialization()
    XlflowAssert.AssertEquals vbObjectError + 2102, CaptureUninitializedSafeParseError(), "uninitialized SafeParse"
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureDirectInitializationError(), "direct InternalInitialize"
    XlflowAssert.AssertEquals vbObjectError + 2102, CaptureReinitializationError(), "InternalInitialize twice"
    XlflowAssert.AssertEquals vbObjectError + 2102, CaptureUninitializedResultError(), "uninitialized result"
End Sub

Private Function CaptureTextMinError(ByVal Argument As Variant) As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.Text()
    On Error Resume Next
    Err.Clear
    Call targetSchema.Min(Argument)
    CaptureTextMinError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureTextLengthError(ByVal Argument As Variant) As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.Text()
    On Error Resume Next
    Err.Clear
    Call targetSchema.Length(Argument)
    CaptureTextLengthError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureNumberMinError(ByVal Argument As Variant) As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.Number()
    On Error Resume Next
    Err.Clear
    Call targetSchema.Min(Argument)
    CaptureNumberMinError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureTextMaxAfterMinError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.Text().Min(3)
    On Error Resume Next
    Err.Clear
    Call targetSchema.Max(2)
    CaptureTextMaxAfterMinError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureTextMinAfterLengthError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.Text().Length(3)
    On Error Resume Next
    Err.Clear
    Call targetSchema.Min(4)
    CaptureTextMinAfterLengthError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureDuplicateWholeNumberError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.Number().WholeNumber()
    On Error Resume Next
    Err.Clear
    Call targetSchema.WholeNumber()
    CaptureDuplicateWholeNumberError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureInvalidWholeNumberKindError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.Text()
    On Error Resume Next
    Err.Clear
    Call targetSchema.WholeNumber()
    CaptureInvalidWholeNumberKindError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureUninitializedSafeParseError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = New VSchema
    On Error Resume Next
    Err.Clear
    Call targetSchema.SafeParse("x")
    CaptureUninitializedSafeParseError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureDirectInitializationError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = New VSchema
    On Error Resume Next
    Err.Clear
    Call targetSchema.InternalInitialize(1)
    CaptureDirectInitializationError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureReinitializationError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = New VSchema
    Call targetSchema.InternalInitialize(27183)
    On Error Resume Next
    Err.Clear
    Call targetSchema.InternalInitialize(27183)
    CaptureReinitializationError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureUninitializedResultError() As Long
    Dim result As VValidationResult
    Set result = New VValidationResult
    Dim success As Boolean
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    success = result.Success
    CaptureUninitializedResultError = Err.Number
    Err.Clear
    On Error GoTo 0
    If success Then success = False
End Function

Public Sub Test_Errors_RejectInvalidObjectBuildersAndCycles()
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureObjectEmptyFieldNameError(), "empty field name"
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureObjectNothingFieldSchemaError(), "Nothing field schema"
    XlflowAssert.AssertEquals vbObjectError + 2101, CaptureObjectDuplicateFieldError(), "duplicate field"
    XlflowAssert.AssertEquals vbObjectError + 2101, CaptureObjectDuplicateStrictError(), "duplicate Strict"
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureObjectUninitializedFieldSchemaError(), "uninitialized field schema"
    XlflowAssert.AssertEquals vbObjectError + 2103, CaptureDirectObjectCycleError(), "direct object cycle"
    XlflowAssert.AssertEquals vbObjectError + 2103, CaptureIndirectObjectCycleError(), "indirect object cycle"
End Sub

Public Sub Test_Errors_RejectInvalidArrayBuilders()
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureArrayNothingItemError(), "ArrayOf Nothing"
    XlflowAssert.AssertEquals vbObjectError + 2101, CaptureArrayDuplicateLengthError(), "duplicate array Length"
    XlflowAssert.AssertEquals vbObjectError + 2101, CaptureArrayMinAfterLengthError(), "array Min greater than Length"
    XlflowAssert.AssertEquals vbObjectError + 2101, CaptureArrayMaxBeforeLengthError(), "array Max less than Length"
End Sub

Public Sub Test_Errors_RejectInvalidLiteralAndEnumDefinitions()
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureLiteralErrorValueError(), "Literal Error value"
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureLiteralObjectError(), "Literal Object value"
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureEnumNonArrayError(), "Enum non-array"
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureEnumUninitializedArrayError(), "Enum uninitialized array"
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureEnumEmptyArrayError(), "Enum empty array"
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureEnumMultidimensionalArrayError(), "Enum multidimensional array"
    XlflowAssert.AssertEquals vbObjectError + 2101, CaptureEnumDuplicateError(), "Enum duplicate candidate"
End Sub

Public Sub Test_Errors_RejectDuplicatePatternAndEmail()
    XlflowAssert.AssertEquals vbObjectError + 2101, CaptureDuplicatePatternError(), "duplicate Pattern"
    XlflowAssert.AssertEquals vbObjectError + 2101, CaptureDuplicateEmailError(), "duplicate Email"
    XlflowAssert.AssertEquals vbObjectError + 2101, CapturePatternOnNumberError(), "Pattern on Number"
    XlflowAssert.AssertEquals vbObjectError + 2101, CaptureEmailOnNumberError(), "Email on Number"
End Sub

Public Sub Test_Errors_RejectInvalidUnionDefinitions()
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureUnionEmptyError(), "empty Union branches"
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureUnionNothingCollectionError(), "Nothing Union collection"
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureUnionNonSchemaError(), "non-schema Union branch"
    XlflowAssert.AssertEquals vbObjectError + 2100, CaptureUnionUninitializedSchemaError(), "uninitialized Union branch"
End Sub

Private Function CaptureDuplicatePatternError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.Text().Pattern("x")
    On Error Resume Next
    Err.Clear
    Call targetSchema.Pattern("y")
    CaptureDuplicatePatternError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureDuplicateEmailError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.Text().Email()
    On Error Resume Next
    Err.Clear
    Call targetSchema.Email()
    CaptureDuplicateEmailError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CapturePatternOnNumberError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.Number()
    On Error Resume Next
    Err.Clear
    Call targetSchema.Pattern("x")
    CapturePatternOnNumberError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureEmailOnNumberError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.Number()
    On Error Resume Next
    Err.Clear
    Call targetSchema.Email()
    CaptureEmailOnNumberError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureUnionEmptyError() As Long
    Dim branches As Collection
    Set branches = New Collection
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Call Schema.UnionOf(branches)
    CaptureUnionEmptyError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureUnionNothingCollectionError() As Long
    Dim branches As Collection
    Set branches = Nothing
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Call Schema.UnionOf(branches)
    CaptureUnionNothingCollectionError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureUnionNonSchemaError() As Long
    Dim branches As Collection
    Set branches = New Collection
    branches.Add "not a schema"
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Call Schema.UnionOf(branches)
    CaptureUnionNonSchemaError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureUnionUninitializedSchemaError() As Long
    Dim branches As Collection
    Set branches = New Collection
    Dim childSchema As VSchema
    Set childSchema = New VSchema
    branches.Add childSchema
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Call Schema.UnionOf(branches)
    CaptureUnionUninitializedSchemaError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureLiteralErrorValueError() As Long
    Dim errorValue As Variant
    errorValue = CVErr(2042)
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Call Schema.Literal(errorValue)
    CaptureLiteralErrorValueError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureLiteralObjectError() As Long
    Dim objectValue As Object
    Set objectValue = CreateObject("Scripting.Dictionary")
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Call Schema.Literal(objectValue)
    CaptureLiteralObjectError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureEnumNonArrayError() As Long
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Call Schema.EnumOf("x")
    CaptureEnumNonArrayError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureEnumUninitializedArrayError() As Long
    Dim candidates() As String
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Call Schema.EnumOf(candidates)
    CaptureEnumUninitializedArrayError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureEnumMultidimensionalArrayError() As Long
    Dim candidates(0 To 0, 0 To 0) As String
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Call Schema.EnumOf(candidates)
    CaptureEnumMultidimensionalArrayError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureEnumEmptyArrayError() As Long
    Dim candidates As Variant
    candidates = Array()
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Call Schema.EnumOf(candidates)
    CaptureEnumEmptyArrayError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureEnumDuplicateError() As Long
    Dim candidates(0 To 1) As String
    candidates(0) = "same"
    candidates(1) = "same"
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Call Schema.EnumOf(candidates)
    CaptureEnumDuplicateError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureArrayNothingItemError() As Long
    Dim nothingSchema As VSchema
    Set nothingSchema = Nothing
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Call Schema.ArrayOf(nothingSchema)
    CaptureArrayNothingItemError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureArrayDuplicateLengthError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.ArrayOf(Schema.Text()).Length(1)
    On Error Resume Next
    Err.Clear
    Call targetSchema.Length(1)
    CaptureArrayDuplicateLengthError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureArrayMinAfterLengthError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.ArrayOf(Schema.Text()).Length(3)
    On Error Resume Next
    Err.Clear
    Call targetSchema.Min(4)
    CaptureArrayMinAfterLengthError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureArrayMaxBeforeLengthError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.ArrayOf(Schema.Text()).Max(2)
    On Error Resume Next
    Err.Clear
    Call targetSchema.Length(3)
    CaptureArrayMaxBeforeLengthError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureObjectEmptyFieldNameError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.ObjectSchema()
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Call targetSchema.Field("", Schema.Text())
    CaptureObjectEmptyFieldNameError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureObjectNothingFieldSchemaError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.ObjectSchema()
    Dim noSchema As VSchema
    Set noSchema = Nothing
    On Error Resume Next
    Err.Clear
    Call targetSchema.Field("name", noSchema)
    CaptureObjectNothingFieldSchemaError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureObjectDuplicateFieldError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.ObjectSchema().Field("name", Schema.Text())
    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Call targetSchema.Field("name", Schema.Text())
    CaptureObjectDuplicateFieldError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureObjectDuplicateStrictError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.ObjectSchema().Strict()
    On Error Resume Next
    Err.Clear
    Call targetSchema.Strict()
    CaptureObjectDuplicateStrictError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureObjectUninitializedFieldSchemaError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.ObjectSchema()
    Dim childSchema As VSchema
    Set childSchema = New VSchema
    On Error Resume Next
    Err.Clear
    Call targetSchema.Field("name", childSchema)
    CaptureObjectUninitializedFieldSchemaError = Err.Number
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureDirectObjectCycleError() As Long
    Dim targetSchema As VSchema
    Set targetSchema = Schema.ObjectSchema()
    Call targetSchema.Field("self", targetSchema)

    Dim inputObject As Object
    Set inputObject = CreateObject("Scripting.Dictionary")
    inputObject.Add "self", inputObject

    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Dim result As VValidationResult
    Set result = targetSchema.SafeParse(inputObject)
    If result Is Nothing Then
        CaptureDirectObjectCycleError = Err.Number
    Else
        CaptureDirectObjectCycleError = 0
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function CaptureIndirectObjectCycleError() As Long
    Dim firstSchema As VSchema
    Dim secondSchema As VSchema
    Set firstSchema = Schema.ObjectSchema()
    Set secondSchema = Schema.ObjectSchema()
    Call firstSchema.Field("second", secondSchema)
    Call secondSchema.Field("first", firstSchema)

    Dim inputObject As Object
    Set inputObject = CreateObject("Scripting.Dictionary")
    inputObject.Add "second", inputObject

    ' xlflow:disable-next-line VBA214
    On Error Resume Next
    Err.Clear
    Dim result As VValidationResult
    Set result = firstSchema.SafeParse(inputObject)
    If result Is Nothing Then
        CaptureIndirectObjectCycleError = Err.Number
    Else
        CaptureIndirectObjectCycleError = 0
    End If
    Err.Clear
    On Error GoTo 0
End Function
