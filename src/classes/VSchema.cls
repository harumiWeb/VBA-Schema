VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const internalKindKey As Long = 27183
Private Const firstSchemaKind As Long = 0
Private Const lastSchemaKind As Long = 9

Private Const schemaKindAny As Long = 0
Private Const schemaKindText As Long = 1
Private Const schemaKindNumber As Long = 2
Private Const schemaKindBool As Long = 3
Private Const schemaKindDateTime As Long = 4
Private Const schemaKindObject As Long = 5
Private Const schemaKindArray As Long = 6
Private Const schemaKindLiteral As Long = 7
Private Const schemaKindEnum As Long = 8
Private Const schemaKindUnion As Long = 9

Private Const arrayStateUninitialized As Long = 0
Private Const arrayStateOneDimensional As Long = 1
Private Const arrayStateMultiDimensional As Long = 2

Private Const scalarCategoryUnsupported As Long = 0
Private Const scalarCategoryEmpty As Long = 1
Private Const scalarCategoryNull As Long = 2
Private Const scalarCategoryString As Long = 3
Private Const scalarCategoryNumber As Long = 4
Private Const scalarCategoryBoolean As Long = 5
Private Const scalarCategoryDate As Long = 6

Private mInitialized As Boolean
Private mKind As Long
Private mOptional As Boolean
Private mNullable As Boolean
Private mHasMin As Boolean
Private mMinValue As Variant
Private mHasMax As Boolean
Private mMaxValue As Variant
Private mHasLength As Boolean
Private mLengthValue As Long
Private mIntegerOnly As Boolean
Private mHasPattern As Boolean
Private mPattern As String
Private mEmail As Boolean
Private mHasLiteral As Boolean
Private mLiteralValue As Variant
Private mEnumValues As Variant
Private mEnumCount As Long
Private mPatternRegExp As Object
Private mEmailRegExp As Object
Private mFields As Object
Private mFieldOrder As Collection
Private mStrict As Boolean
Private mItemSchema As VSchema
Private mUnionSchemas As Collection

Public Function InternalInitialize(ByVal KindCode As Long) As VSchema
    Dim decodedKind As Long

    If mInitialized Then
        Err.Raise vbObjectError + 2102, "VSchema.InternalInitialize", "Schema has already been initialized."
    End If

    decodedKind = KindCode - internalKindKey
    If decodedKind < firstSchemaKind Or decodedKind > lastSchemaKind Then
        Err.Raise vbObjectError + 2100, "VSchema.InternalInitialize", "Invalid internal schema kind."
    End If

    mKind = decodedKind
    If mKind = schemaKindObject Then
        Set mFields = NewDictionary("VSchema.InternalInitialize")
        Set mFieldOrder = New Collection
    End If
    mInitialized = True
    Set InternalInitialize = Me
End Function

' Internal* members are implementation hooks for composition and are not stable user-facing API.
Public Function InternalEnsureComposed(ByVal ProcedureName As String) As VSchema
    If Not mInitialized Then
        Err.Raise vbObjectError + 2100, ProcedureName, "Field schema must be initialized."
    End If
    Set InternalEnsureComposed = Me
End Function

Public Function InternalValidateAtPath(ByVal InputValue As Variant, ByVal Path As String, ByVal Issues As Collection) As VSchema
    EnsureInitialized "VSchema.InternalValidateAtPath"
    EnsureSupportedKind "VSchema.InternalValidateAtPath"
    ValidateValue InputValue, Path, Issues
    Set InternalValidateAtPath = Me
End Function

Public Function InternalValidateMissing(ByVal Path As String, ByVal Issues As Collection) As VSchema
    EnsureInitialized "VSchema.InternalValidateMissing"
    EnsureSupportedKind "VSchema.InternalValidateMissing"
    If Not mOptional Then
        AddIssue Issues, Path, "required", "A value is required.", ExpectedForKind(), "Missing"
    End If
    Set InternalValidateMissing = Me
End Function

Public Function InternalVisitSchemaGraph(ByVal ActiveSchemas As Collection, ByVal ProcedureName As String) As VSchema
    EnsureInitialized ProcedureName
    VisitSchemaGraph ActiveSchemas, ProcedureName
    Set InternalVisitSchemaGraph = Me
End Function

Public Function InternalSetItemSchema(ByVal ItemSchema As VSchema) As VSchema
    EnsureInitialized "VSchema.InternalSetItemSchema"
    If mKind <> schemaKindArray Then
        Err.Raise vbObjectError + 2101, "VSchema.InternalSetItemSchema", "Item schema is only valid for Array schemas."
    End If
    If ItemSchema Is Nothing Then
        Err.Raise vbObjectError + 2100, "VSchema.InternalSetItemSchema", "Item schema is required."
    End If
    Call ItemSchema.InternalEnsureComposed("VSchema.InternalSetItemSchema")
    If Not mItemSchema Is Nothing Then
        Err.Raise vbObjectError + 2101, "VSchema.InternalSetItemSchema", "Item schema has already been specified."
    End If
    Set mItemSchema = ItemSchema
    Set InternalSetItemSchema = Me
End Function

Public Function InternalSetLiteral(ByVal ExpectedValue As Variant) As VSchema
    EnsureInitialized "VSchema.InternalSetLiteral"
    If mKind <> schemaKindLiteral Then
        Err.Raise vbObjectError + 2101, "VSchema.InternalSetLiteral", "Literal value is only valid for Literal schemas."
    End If
    If mHasLiteral Then
        Err.Raise vbObjectError + 2101, "VSchema.InternalSetLiteral", "Literal value has already been specified."
    End If

    ValidateScalarDefinition ExpectedValue, "VSchema.InternalSetLiteral"
    mLiteralValue = ExpectedValue
    mHasLiteral = True
    Set InternalSetLiteral = Me
End Function

Public Function InternalSetEnum(ByVal Values As Variant) As VSchema
    EnsureInitialized "VSchema.InternalSetEnum"
    If mKind <> schemaKindEnum Then
        Err.Raise vbObjectError + 2101, "VSchema.InternalSetEnum", "Enum values are only valid for Enum schemas."
    End If
    If mEnumCount > 0 Then
        Err.Raise vbObjectError + 2101, "VSchema.InternalSetEnum", "Enum values have already been specified."
    End If
    If Not IsArray(Values) Then
        Err.Raise vbObjectError + 2100, "VSchema.InternalSetEnum", "Enum values must be a one-dimensional array."
    End If
    If VarType(Values) - vbArray = vbObject Then
        Err.Raise vbObjectError + 2100, "VSchema.InternalSetEnum", "Enum values must be scalar values."
    End If

    Dim lowerBound As Long
    Dim upperBound As Long
    Dim arrayState As Long
    arrayState = GetNativeArrayState(Values, lowerBound, upperBound)
    If arrayState <> arrayStateOneDimensional Then
        Err.Raise vbObjectError + 2100, "VSchema.InternalSetEnum", "Enum values must be a non-empty one-dimensional array."
    End If
    If upperBound < lowerBound Then
        Err.Raise vbObjectError + 2100, "VSchema.InternalSetEnum", "Enum values must not be empty."
    End If

    Dim candidateCount As Long
    candidateCount = upperBound - lowerBound + 1
    If candidateCount <= 0 Then
        Err.Raise vbObjectError + 2100, "VSchema.InternalSetEnum", "Enum values must not be empty."
    End If

    Dim normalizedValues() As Variant
    ReDim normalizedValues(0 To candidateCount - 1)

    Dim sourceIndex As Long
    Dim normalizedIndex As Long
    Dim candidate As Variant
    Dim previousIndex As Long
    For sourceIndex = lowerBound To upperBound
        candidate = Values(sourceIndex)
        ValidateScalarDefinition candidate, "VSchema.InternalSetEnum"
        normalizedIndex = sourceIndex - lowerBound
        For previousIndex = 0 To normalizedIndex - 1
            If ScalarValuesEqual(candidate, normalizedValues(previousIndex), "VSchema.InternalSetEnum") Then
                Err.Raise vbObjectError + 2101, "VSchema.InternalSetEnum", "Enum values must be unique."
            End If
        Next previousIndex
        normalizedValues(normalizedIndex) = candidate
    Next sourceIndex

    mEnumValues = normalizedValues
    mEnumCount = candidateCount
    Set InternalSetEnum = Me
End Function

Public Function InternalSetUnion(ByVal Schemas As Collection) As VSchema
    EnsureInitialized "VSchema.InternalSetUnion"
    If mKind <> schemaKindUnion Then
        Err.Raise vbObjectError + 2101, "VSchema.InternalSetUnion", "Union branches are only valid for Union schemas."
    End If
    If Not mUnionSchemas Is Nothing Then
        Err.Raise vbObjectError + 2101, "VSchema.InternalSetUnion", "Union branches have already been specified."
    End If
    If Schemas Is Nothing Then
        Err.Raise vbObjectError + 2100, "VSchema.InternalSetUnion", "Union branches are required."
    End If
    If Schemas.Count = 0 Then
        Err.Raise vbObjectError + 2100, "VSchema.InternalSetUnion", "Union branches must not be empty."
    End If

    Dim normalizedSchemas As Collection
    Set normalizedSchemas = New Collection

    Dim branchIndex As Long
    Dim branchTypeName As String
    Dim branch As VSchema
    For branchIndex = 1 To Schemas.Count
        branchTypeName = TypeName(Schemas.Item(branchIndex))
        If branchTypeName <> "VSchema" Then
            Err.Raise vbObjectError + 2100, "VSchema.InternalSetUnion", "Union branches must be VSchema instances."
        End If
        Set branch = Schemas.Item(branchIndex)
        If branch Is Nothing Then
            Err.Raise vbObjectError + 2100, "VSchema.InternalSetUnion", "Union branches must be initialized VSchema instances."
        End If
        Call branch.InternalEnsureComposed("VSchema.InternalSetUnion")
        normalizedSchemas.Add branch
    Next branchIndex

    Set mUnionSchemas = normalizedSchemas
    Set InternalSetUnion = Me
End Function

Public Function OptionalField() As VSchema
    EnsureInitialized "VSchema.OptionalField"
    If mOptional Then
        Err.Raise vbObjectError + 2101, "VSchema.OptionalField", "OptionalField has already been specified."
    End If
    mOptional = True
    Set OptionalField = Me
End Function

Public Function Nullable() As VSchema
    EnsureInitialized "VSchema.Nullable"
    If mNullable Then
        Err.Raise vbObjectError + 2101, "VSchema.Nullable", "Nullable has already been specified."
    End If
    mNullable = True
    Set Nullable = Me
End Function

Public Function Min(ByVal Boundary As Variant) As VSchema
    EnsureInitialized "VSchema.Min"

    Select Case mKind
    Case schemaKindText
        If mHasMin Then
            Err.Raise vbObjectError + 2101, "VSchema.Min", "Min has already been specified."
        End If

        Dim textMinimum As Long
        textMinimum = NormalizeLengthBoundary(Boundary, "VSchema.Min")
        If mHasLength Then
            If textMinimum > mLengthValue Then
                Err.Raise vbObjectError + 2101, "VSchema.Min", "Min is greater than Length."
            End If
        End If
        If mHasMax Then
            If textMinimum > CLng(mMaxValue) Then
                Err.Raise vbObjectError + 2101, "VSchema.Min", "Min is greater than Max."
            End If
        End If
        mMinValue = textMinimum
        mHasMin = True

    Case schemaKindNumber
        If mHasMin Then
            Err.Raise vbObjectError + 2101, "VSchema.Min", "Min has already been specified."
        End If

        ValidateNumericBoundary Boundary, "VSchema.Min"
        If mHasMax Then
            If CompareNumeric(Boundary, mMaxValue, "VSchema.Min") > 0 Then
                Err.Raise vbObjectError + 2101, "VSchema.Min", "Min is greater than Max."
            End If
        End If
        mMinValue = Boundary
        mHasMin = True

    Case schemaKindArray
        If mHasMin Then
            Err.Raise vbObjectError + 2101, "VSchema.Min", "Min has already been specified."
        End If

        Dim arrayMinimum As Long
        arrayMinimum = NormalizeLengthBoundary(Boundary, "VSchema.Min")
        If mHasLength Then
            If arrayMinimum > mLengthValue Then
                Err.Raise vbObjectError + 2101, "VSchema.Min", "Min is greater than Length."
            End If
        End If
        If mHasMax Then
            If arrayMinimum > CLng(mMaxValue) Then
                Err.Raise vbObjectError + 2101, "VSchema.Min", "Min is greater than Max."
            End If
        End If
        mMinValue = arrayMinimum
        mHasMin = True

    Case Else
        Err.Raise vbObjectError + 2101, "VSchema.Min", "Min is not valid for this schema kind."
    End Select

    Set Min = Me
End Function

Public Function Max(ByVal Boundary As Variant) As VSchema
    EnsureInitialized "VSchema.Max"

    Select Case mKind
    Case schemaKindText
        If mHasMax Then
            Err.Raise vbObjectError + 2101, "VSchema.Max", "Max has already been specified."
        End If

        Dim textMaximum As Long
        textMaximum = NormalizeLengthBoundary(Boundary, "VSchema.Max")
        If mHasLength Then
            If textMaximum < mLengthValue Then
                Err.Raise vbObjectError + 2101, "VSchema.Max", "Max is less than Length."
            End If
        End If
        If mHasMin Then
            If textMaximum < CLng(mMinValue) Then
                Err.Raise vbObjectError + 2101, "VSchema.Max", "Max is less than Min."
            End If
        End If
        mMaxValue = textMaximum
        mHasMax = True

    Case schemaKindNumber
        If mHasMax Then
            Err.Raise vbObjectError + 2101, "VSchema.Max", "Max has already been specified."
        End If

        ValidateNumericBoundary Boundary, "VSchema.Max"
        If mHasMin Then
            If CompareNumeric(mMinValue, Boundary, "VSchema.Max") > 0 Then
                Err.Raise vbObjectError + 2101, "VSchema.Max", "Max is less than Min."
            End If
        End If
        mMaxValue = Boundary
        mHasMax = True

    Case schemaKindArray
        If mHasMax Then
            Err.Raise vbObjectError + 2101, "VSchema.Max", "Max has already been specified."
        End If

        Dim arrayMaximum As Long
        arrayMaximum = NormalizeLengthBoundary(Boundary, "VSchema.Max")
        If mHasLength Then
            If arrayMaximum < mLengthValue Then
                Err.Raise vbObjectError + 2101, "VSchema.Max", "Max is less than Length."
            End If
        End If
        If mHasMin Then
            If arrayMaximum < CLng(mMinValue) Then
                Err.Raise vbObjectError + 2101, "VSchema.Max", "Max is less than Min."
            End If
        End If
        mMaxValue = arrayMaximum
        mHasMax = True

    Case Else
        Err.Raise vbObjectError + 2101, "VSchema.Max", "Max is not valid for this schema kind."
    End Select

    Set Max = Me
End Function

Public Function Length(ByVal RequiredLength As Variant) As VSchema
    EnsureInitialized "VSchema.Length"
    If mKind <> schemaKindText And mKind <> schemaKindArray Then
        Err.Raise vbObjectError + 2101, "VSchema.Length", "Length is not valid for this schema kind."
    End If
    If mHasLength Then
        Err.Raise vbObjectError + 2101, "VSchema.Length", "Length has already been specified."
    End If

    Dim normalizedLength As Long
    normalizedLength = NormalizeLengthBoundary(RequiredLength, "VSchema.Length")
    If mHasMin Then
        If CLng(mMinValue) > normalizedLength Then
            Err.Raise vbObjectError + 2101, "VSchema.Length", "Length is less than Min."
        End If
    End If
    If mHasMax Then
        If CLng(mMaxValue) < normalizedLength Then
            Err.Raise vbObjectError + 2101, "VSchema.Length", "Length is greater than Max."
        End If
    End If

    mLengthValue = normalizedLength
    mHasLength = True
    Set Length = Me
End Function

Public Function WholeNumber() As VSchema
    EnsureInitialized "VSchema.WholeNumber"
    If mKind <> schemaKindNumber Then
        Err.Raise vbObjectError + 2101, "VSchema.WholeNumber", "WholeNumber is only valid for Number schemas."
    End If
    If mIntegerOnly Then
        Err.Raise vbObjectError + 2101, "VSchema.WholeNumber", "WholeNumber has already been specified."
    End If
    mIntegerOnly = True
    Set WholeNumber = Me
End Function

Public Function Pattern(ByVal Expression As String) As VSchema
    EnsureInitialized "VSchema.Pattern"
    If mKind <> schemaKindText Then
        Err.Raise vbObjectError + 2101, "VSchema.Pattern", "Pattern is only valid for Text schemas."
    End If
    If mHasPattern Then
        Err.Raise vbObjectError + 2101, "VSchema.Pattern", "Pattern has already been specified."
    End If
    mPattern = Expression
    mHasPattern = True
    Set Pattern = Me
End Function

Public Function Email() As VSchema
    EnsureInitialized "VSchema.Email"
    If mKind <> schemaKindText Then
        Err.Raise vbObjectError + 2101, "VSchema.Email", "Email is only valid for Text schemas."
    End If
    If mEmail Then
        Err.Raise vbObjectError + 2101, "VSchema.Email", "Email has already been specified."
    End If
    mEmail = True
    Set Email = Me
End Function

Public Function Field(ByVal FieldName As String, ByVal FieldSchema As VSchema) As VSchema
    EnsureInitialized "VSchema.Field"
    If mKind <> schemaKindObject Then
        Err.Raise vbObjectError + 2101, "VSchema.Field", "Field is only valid for Object schemas."
    End If
    If Len(FieldName) = 0 Then
        Err.Raise vbObjectError + 2100, "VSchema.Field", "Field name is required."
    End If
    If FieldSchema Is Nothing Then
        Err.Raise vbObjectError + 2100, "VSchema.Field", "Field schema is required."
    End If
    Call FieldSchema.InternalEnsureComposed("VSchema.Field")
    If mFields Is Nothing Or mFieldOrder Is Nothing Then
        Err.Raise vbObjectError + 2299, "VSchema.Field", "Object schema state is not initialized."
    End If
    If mFields.Exists(FieldName) Then
        Err.Raise vbObjectError + 2101, "VSchema.Field", "The field has already been specified."
    End If

    mFields.Add FieldName, FieldSchema
    mFieldOrder.Add FieldName
    Set Field = Me
End Function

Public Function Strict() As VSchema
    EnsureInitialized "VSchema.Strict"
    If mKind <> schemaKindObject Then
        Err.Raise vbObjectError + 2101, "VSchema.Strict", "Strict is only valid for Object schemas."
    End If
    If mStrict Then
        Err.Raise vbObjectError + 2101, "VSchema.Strict", "Strict has already been specified."
    End If
    mStrict = True
    Set Strict = Me
End Function

Public Function SafeParse(ByVal InputValue As Variant) As VValidationResult
    EnsureInitialized "VSchema.SafeParse"
    EnsureSupportedKind "VSchema.SafeParse"
    EnsureAcyclicGraph "VSchema.SafeParse"

    Dim issues As Collection
    Set issues = New Collection
    ValidateValue InputValue, "$", issues

    Dim result As VValidationResult
    Set result = New VValidationResult

    Set SafeParse = result.InternalInitialize(issues.Count = 0, InputValue, issues)
End Function

Private Sub ValidateValue(ByVal InputValue As Variant, ByVal Path As String, ByVal Issues As Collection)
    If IsNull(InputValue) Then
        If mKind = schemaKindAny Or mNullable Then
            Exit Sub
        End If
        If mKind = schemaKindLiteral Or mKind = schemaKindEnum Or mKind = schemaKindUnion Then
            If mKind = schemaKindUnion Then
                ValidateUnion InputValue, Path, Issues
            Else
                ValidateLiteralOrEnum InputValue, Path, Issues
            End If
            Exit Sub
        End If
        AddIssue Issues, Path, "required", "A value is required.", ExpectedForKind(), DescribeValue(InputValue)
        Exit Sub
    End If

    If mKind = schemaKindAny Then
        Exit Sub
    End If

    If IsEmpty(InputValue) Then
        If mKind = schemaKindLiteral Or mKind = schemaKindEnum Or mKind = schemaKindUnion Then
            If mKind = schemaKindUnion Then
                ValidateUnion InputValue, Path, Issues
            Else
                ValidateLiteralOrEnum InputValue, Path, Issues
            End If
            Exit Sub
        End If
        AddIssue Issues, Path, "invalid_type", "The value has an invalid type.", ExpectedForKind(), DescribeValue(InputValue)
        Exit Sub
    End If

    If IsError(InputValue) Then
        If mKind = schemaKindLiteral Or mKind = schemaKindEnum Or mKind = schemaKindUnion Then
            If mKind = schemaKindUnion Then
                ValidateUnion InputValue, Path, Issues
            Else
                ValidateLiteralOrEnum InputValue, Path, Issues
            End If
            Exit Sub
        End If
        AddIssue Issues, Path, "invalid_type", "The value has an invalid type.", ExpectedForKind(), DescribeValue(InputValue)
        Exit Sub
    End If

    Select Case mKind
    Case schemaKindText
        ValidateText InputValue, Path, Issues
    Case schemaKindNumber
        ValidateNumber InputValue, Path, Issues
    Case schemaKindBool
        ValidateBool InputValue, Path, Issues
    Case schemaKindDateTime
        ValidateDateTime InputValue, Path, Issues
    Case schemaKindObject
        ValidateObject InputValue, Path, Issues
    Case schemaKindArray
        ValidateArray InputValue, Path, Issues
    Case schemaKindLiteral, schemaKindEnum
        ValidateLiteralOrEnum InputValue, Path, Issues
    Case schemaKindUnion
        ValidateUnion InputValue, Path, Issues
    Case Else
        RaiseNotImplemented "VSchema.ValidateValue"
    End Select
End Sub

Private Sub ValidateText(ByVal InputValue As Variant, ByVal Path As String, ByVal Issues As Collection)
    If VarType(InputValue) <> vbString Then
        AddIssue Issues, Path, "invalid_type", "The value has an invalid type.", "Text", DescribeValue(InputValue)
        Exit Sub
    End If

    Dim actualLength As Long
    actualLength = Len(CStr(InputValue))
    If mHasLength Then
        If actualLength <> mLengthValue Then
            AddIssue Issues, Path, "invalid_length", "Text length must equal " & CStr(mLengthValue) & ".", "Text length = " & CStr(mLengthValue), DescribeValue(InputValue)
            Exit Sub
        End If
    End If
    If mHasMin Then
        If actualLength < CLng(mMinValue) Then
            AddIssue Issues, Path, "too_small", "Text length is below the minimum.", "Text length >= " & CStr(CLng(mMinValue)), DescribeValue(InputValue)
            Exit Sub
        End If
    End If
    If mHasMax Then
        If actualLength > CLng(mMaxValue) Then
            AddIssue Issues, Path, "too_big", "Text length is above the maximum.", "Text length <= " & CStr(CLng(mMaxValue)), DescribeValue(InputValue)
            Exit Sub
        End If
    End If
    If mHasPattern Then
        If Not MatchesPattern(CStr(InputValue)) Then
            AddIssue Issues, Path, "invalid_pattern", "String does not match the pattern.", "Pattern", DescribeValue(InputValue)
            Exit Sub
        End If
    End If
    If mEmail Then
        If Not MatchesEmail(CStr(InputValue)) Then
            AddIssue Issues, Path, "invalid_email", "Invalid email address.", "email", DescribeValue(InputValue)
            Exit Sub
        End If
    End If
End Sub

Private Sub ValidateNumber(ByVal InputValue As Variant, ByVal Path As String, ByVal Issues As Collection)
    If Not IsSupportedNumber(InputValue) Then
        AddIssue Issues, Path, "invalid_type", "The value has an invalid type.", "Number", DescribeValue(InputValue)
        Exit Sub
    End If

    Dim decimalValue As Variant
    If Not TryDecimal(InputValue, decimalValue) Then
        AddIssue Issues, Path, "invalid_type", "The value is not a finite Number.", "Number", DescribeValue(InputValue)
        Exit Sub
    End If

    If mHasMin Then
        If CompareNumeric(InputValue, mMinValue, "VSchema.ValidateNumber") < 0 Then
            AddIssue Issues, Path, "too_small", "Number is below the minimum.", "Number >= " & NumberText(mMinValue), DescribeValue(InputValue)
            Exit Sub
        End If
    End If
    If mHasMax Then
        If CompareNumeric(InputValue, mMaxValue, "VSchema.ValidateNumber") > 0 Then
            AddIssue Issues, Path, "too_big", "Number is above the maximum.", "Number <= " & NumberText(mMaxValue), DescribeValue(InputValue)
            Exit Sub
        End If
    End If
    If mIntegerOnly Then
        If Not IsWholeNumberValue(InputValue) Then
            AddIssue Issues, Path, "invalid_integer", "Number must be a whole number.", "Whole Number", DescribeValue(InputValue)
            Exit Sub
        End If
    End If
End Sub

Private Sub ValidateBool(ByVal InputValue As Variant, ByVal Path As String, ByVal Issues As Collection)
    If VarType(InputValue) <> vbBoolean Then
        AddIssue Issues, Path, "invalid_type", "The value has an invalid type.", "Boolean", DescribeValue(InputValue)
    End If
End Sub

Private Sub ValidateDateTime(ByVal InputValue As Variant, ByVal Path As String, ByVal Issues As Collection)
    If VarType(InputValue) <> vbDate Then
        AddIssue Issues, Path, "invalid_type", "The value has an invalid type.", "Date", DescribeValue(InputValue)
    End If
End Sub

Private Sub ValidateLiteralOrEnum(ByVal InputValue As Variant, ByVal Path As String, ByVal Issues As Collection)
    If ScalarCategory(InputValue) = scalarCategoryUnsupported Then
        AddIssue Issues, Path, "invalid_type", "The value has an invalid type.", ExpectedForKind(), DescribeValue(InputValue)
        Exit Sub
    End If

    If mKind = schemaKindLiteral Then
        If ScalarValuesEqual(InputValue, mLiteralValue, "VSchema.ValidateLiteral") Then
            Exit Sub
        End If
        AddIssue Issues, Path, "invalid_literal", "The value does not match the literal.", "Literal(" & DescribeValue(mLiteralValue) & ")", DescribeValue(InputValue)
        Exit Sub
    End If

    Dim enumIndex As Long
    For enumIndex = 0 To mEnumCount - 1
        If ScalarValuesEqual(InputValue, mEnumValues(enumIndex), "VSchema.ValidateEnum") Then
            Exit Sub
        End If
    Next enumIndex
    AddIssue Issues, Path, "invalid_enum", "The value is not one of the allowed enum values.", "Enum", DescribeValue(InputValue)
End Sub

Private Sub ValidateUnion(ByVal InputValue As Variant, ByVal Path As String, ByVal Issues As Collection)
    If mUnionSchemas Is Nothing Then
        Err.Raise vbObjectError + 2299, "VSchema.ValidateUnion", "Union schema state is not initialized."
    End If
    If mUnionSchemas.Count = 0 Then
        Err.Raise vbObjectError + 2299, "VSchema.ValidateUnion", "Union schema state is not initialized."
    End If

    Dim branchIndex As Long
    Dim branch As VSchema
    Dim branchIssues As Collection
    For branchIndex = 1 To mUnionSchemas.Count
        Set branch = mUnionSchemas.Item(branchIndex)
        Set branchIssues = New Collection
        If branchIssues Is Nothing Then
            Err.Raise vbObjectError + 2299, "VSchema.ValidateUnion", "Union issue state is not initialized."
        End If
        Call branch.InternalValidateAtPath(InputValue, Path, branchIssues)
        If branchIssues Is Nothing Then
            Err.Raise vbObjectError + 2299, "VSchema.ValidateUnion", "Union issue state is not initialized."
        End If
        If branchIssues.Count = 0 Then
            Exit Sub
        End If
    Next branchIndex

    AddIssue Issues, Path, "invalid_union", "The value does not match any union branch.", "Union", DescribeValue(InputValue)
End Sub

Private Function MatchesPattern(ByVal InputText As String) As Boolean
    Dim regex As Object
    Set regex = GetPatternRegExp()
    MatchesPattern = FullRegExpMatch(regex, InputText, "VSchema.Pattern")
End Function

Private Function MatchesEmail(ByVal InputText As String) As Boolean
    If Len(InputText) > 254 Then
        Exit Function
    End If

    Dim regex As Object
    Set regex = GetEmailRegExp()
    MatchesEmail = FullRegExpMatch(regex, InputText, "VSchema.Email")
End Function

Private Function GetPatternRegExp() As Object
    If Not mPatternRegExp Is Nothing Then
        Set GetPatternRegExp = mPatternRegExp
        Exit Function
    End If

    Dim regex As Object
    Set regex = CreateRegExp("VSchema.Pattern")
    On Error GoTo InvalidPattern
    regex.IgnoreCase = False
    regex.Global = False
    regex.MultiLine = False
    regex.Pattern = "^(" & mPattern & ")$"
    Call regex.Test(vbNullString)
    Set mPatternRegExp = regex
    Set GetPatternRegExp = regex
    Exit Function

InvalidPattern:
    Err.Raise vbObjectError + 2100, "VSchema.Pattern", "Pattern expression is invalid."
End Function

Private Function GetEmailRegExp() As Object
    If Not mEmailRegExp Is Nothing Then
        Set GetEmailRegExp = mEmailRegExp
        Exit Function
    End If

    Dim regex As Object
    Set regex = CreateRegExp("VSchema.Email")
    On Error GoTo InvalidEmailPattern
    regex.IgnoreCase = False
    regex.Global = False
    regex.MultiLine = False
    regex.Pattern = "^([A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+(\.[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+)*)@([A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?)(\.([A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?))+$"
    Set mEmailRegExp = regex
    Set GetEmailRegExp = regex
    Exit Function

InvalidEmailPattern:
    Err.Raise vbObjectError + 2299, "VSchema.Email", "The built-in email expression is invalid."
End Function

Private Function CreateRegExp(ByVal Source As String) As Object
    On Error GoTo Failed
    Set CreateRegExp = CreateObject("VBScript.RegExp")
    Exit Function

Failed:
    Err.Raise vbObjectError + 2200, Source, "VBScript.RegExp runtime is unavailable."
End Function

Private Function FullRegExpMatch(ByVal RegExpObject As Object, ByVal InputText As String, ByVal Source As String) As Boolean
    On Error GoTo Failed

    Dim matches As Object
    Set matches = RegExpObject.Execute(InputText)
    If matches Is Nothing Then
        Exit Function
    End If
    If matches.Count = 0 Then
        Exit Function
    End If

    Dim match As Object
    Set match = matches.Item(0)
    If match Is Nothing Then
        Exit Function
    End If
    If CLng(match.FirstIndex) <> 0 Then
        Exit Function
    End If
    If CLng(match.Length) <> Len(InputText) Then
        Exit Function
    End If
    FullRegExpMatch = True
    Exit Function

Failed:
    Err.Raise vbObjectError + 2200, Source, "VBScript.RegExp runtime is unavailable."
End Function

Private Sub ValidateScalarDefinition(ByVal Candidate As Variant, ByVal Source As String)
    Dim category As Long
    category = ScalarCategory(Candidate)
    If category = scalarCategoryUnsupported Then
        Err.Raise vbObjectError + 2100, Source, "Literal and Enum values must be scalar values."
    End If
    If category = scalarCategoryNumber Then
        Dim decimalValue As Variant
        If Not TryDecimal(Candidate, decimalValue) Then
            Err.Raise vbObjectError + 2100, Source, "Numeric definition value cannot be represented deterministically."
        End If
    End If
End Sub

Private Function ScalarCategory(ByVal InputValue As Variant) As Long
    If IsNull(InputValue) Then
        ScalarCategory = scalarCategoryNull
        Exit Function
    End If
    If IsEmpty(InputValue) Then
        ScalarCategory = scalarCategoryEmpty
        Exit Function
    End If
    If IsError(InputValue) Or IsObject(InputValue) Or IsArray(InputValue) Then
        ScalarCategory = scalarCategoryUnsupported
        Exit Function
    End If

    Select Case VarType(InputValue)
    Case vbString
        ScalarCategory = scalarCategoryString
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency
        ScalarCategory = scalarCategoryNumber
    Case vbBoolean
        ScalarCategory = scalarCategoryBoolean
    Case vbDate
        ScalarCategory = scalarCategoryDate
    Case Else
        ScalarCategory = scalarCategoryUnsupported
    End Select
End Function

Private Function ScalarValuesEqual(ByVal LeftValue As Variant, ByVal RightValue As Variant, ByVal Source As String) As Boolean
    Dim leftCategory As Long
    Dim rightCategory As Long
    leftCategory = ScalarCategory(LeftValue)
    rightCategory = ScalarCategory(RightValue)
    If leftCategory = scalarCategoryUnsupported Or rightCategory = scalarCategoryUnsupported Then
        Err.Raise vbObjectError + 2100, Source, "Only scalar values can be compared."
    End If
    If leftCategory <> rightCategory Then
        Exit Function
    End If

    Select Case leftCategory
    Case scalarCategoryEmpty, scalarCategoryNull
        ScalarValuesEqual = True
    Case scalarCategoryString
        ScalarValuesEqual = (StrComp(CStr(LeftValue), CStr(RightValue), vbBinaryCompare) = 0)
    Case scalarCategoryNumber
        ScalarValuesEqual = (CompareNumeric(LeftValue, RightValue, Source) = 0)
    Case scalarCategoryBoolean
        ScalarValuesEqual = (CBool(LeftValue) = CBool(RightValue))
    Case scalarCategoryDate
        ScalarValuesEqual = (CDate(LeftValue) = CDate(RightValue))
    End Select
End Function

Private Sub ValidateObject(ByVal InputValue As Variant, ByVal Path As String, ByVal Issues As Collection)
    If mFields Is Nothing Or mFieldOrder Is Nothing Then
        Err.Raise vbObjectError + 2299, "VSchema.ValidateObject", "Object schema state is not initialized."
    End If

    Dim inputDictionary As Object
    Set inputDictionary = GetDictionary(InputValue, "VSchema.ValidateObject")
    If inputDictionary Is Nothing Then
        AddIssue Issues, Path, "invalid_type", "The value has an invalid type.", "Object(Dictionary)", DescribeValue(InputValue)
        Exit Sub
    End If

    Dim fieldIndex As Long
    Dim fieldName As String
    Dim fieldSchema As VSchema
    Dim matchedKey As Variant
    Dim fieldValue As Variant
    Dim childPath As String

    For fieldIndex = 1 To mFieldOrder.Count
        fieldName = CStr(mFieldOrder.Item(fieldIndex))
        Set fieldSchema = mFields.Item(fieldName)
        If TryFindDictionaryKey(inputDictionary, fieldName, matchedKey) Then
            ReadDictionaryValue inputDictionary, matchedKey, "VSchema.ValidateObject", fieldValue
            childPath = AppendObjectPath(Path, fieldName)
            Call fieldSchema.InternalValidateAtPath(fieldValue, childPath, Issues)
        Else
            Call fieldSchema.InternalValidateMissing(AppendObjectPath(Path, fieldName), Issues)
        End If
    Next fieldIndex

    Dim invalidKeys As Collection
    Set invalidKeys = New Collection
    Dim unknownKeys As Collection
    Set unknownKeys = New Collection
    CollectSpecialKeys inputDictionary, invalidKeys, unknownKeys

    Dim sortedValue As Variant
    For Each sortedValue In SortedStrings(invalidKeys)
        AddIssue Issues, Path, "invalid_key", "Object keys must be strings.", "String key", CStr(sortedValue)
    Next sortedValue

    If mStrict Then
        For Each sortedValue In SortedStrings(unknownKeys)
            AddIssue Issues, AppendObjectPath(Path, CStr(sortedValue)), "unknown_field", "Unknown field.", "Declared field", DescribeValue(CStr(sortedValue))
        Next sortedValue
    End If
End Sub

Private Sub ValidateArray(ByVal InputValue As Variant, ByVal Path As String, ByVal Issues As Collection)
    If mItemSchema Is Nothing Then
        Err.Raise vbObjectError + 2299, "VSchema.ValidateArray", "Array item schema is not initialized."
    End If

    Dim elementCount As Long
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim arrayState As Long
    Dim collectionValue As Object
    Dim isNativeArray As Boolean

    If IsArray(InputValue) Then
        isNativeArray = True
        arrayState = GetNativeArrayState(InputValue, lowerBound, upperBound)
        If arrayState = arrayStateMultiDimensional Then
            AddIssue Issues, Path, "invalid_array_rank", "Only one-dimensional arrays are supported.", "Array rank = 1", DescribeValue(InputValue)
            Exit Sub
        End If
        If arrayState = arrayStateOneDimensional Then
            elementCount = upperBound - lowerBound + 1
        Else
            elementCount = 0
        End If
    Else
        Set collectionValue = GetCollection(InputValue)
        If collectionValue Is Nothing Then
            AddIssue Issues, Path, "invalid_type", "The value has an invalid type.", "Array", DescribeValue(InputValue)
            Exit Sub
        End If
        On Error GoTo CollectionCountUnavailable
        elementCount = CLng(collectionValue.Count)
        On Error GoTo 0
    End If

    If Not ValidateArrayLength(elementCount, Path, InputValue, Issues) Then
        Exit Sub
    End If

    Dim logicalIndex As Long
    Dim elementValue As Variant
    For logicalIndex = 0 To elementCount - 1
        If isNativeArray Then
            ReadArrayValue InputValue, lowerBound + logicalIndex, "VSchema.ValidateArray", elementValue
        Else
            ReadCollectionValue collectionValue, logicalIndex + 1, "VSchema.ValidateArray", elementValue
        End If
        Call mItemSchema.InternalValidateAtPath(elementValue, AppendArrayPath(Path, logicalIndex), Issues)
    Next logicalIndex
    Exit Sub

CollectionCountUnavailable:
    On Error GoTo 0
    Err.Raise vbObjectError + 2200, "VSchema.ValidateArray", "Collection count is unavailable."
End Sub

Private Function ValidateArrayLength(ByVal ElementCount As Long, ByVal Path As String, ByVal InputValue As Variant, ByVal Issues As Collection) As Boolean
    If mHasLength Then
        If ElementCount <> mLengthValue Then
            AddIssue Issues, Path, "invalid_length", "Array length must equal " & CStr(mLengthValue) & ".", "Array length = " & CStr(mLengthValue), DescribeValue(InputValue)
            Exit Function
        End If
    End If
    If mHasMin Then
        If ElementCount < CLng(mMinValue) Then
            AddIssue Issues, Path, "too_small", "Array length is below the minimum.", "Array length >= " & CStr(CLng(mMinValue)), DescribeValue(InputValue)
            Exit Function
        End If
    End If
    If mHasMax Then
        If ElementCount > CLng(mMaxValue) Then
            AddIssue Issues, Path, "too_big", "Array length is above the maximum.", "Array length <= " & CStr(CLng(mMaxValue)), DescribeValue(InputValue)
            Exit Function
        End If
    End If
    ValidateArrayLength = True
End Function

Private Function GetCollection(ByVal InputValue As Variant) As Object
    If Not IsObject(InputValue) Then
        Exit Function
    End If
    If StrComp(TypeName(InputValue), "Collection", vbBinaryCompare) <> 0 Then
        Exit Function
    End If

    On Error GoTo CollectionInvalid
    Dim collectionValue As Object
    Set collectionValue = InputValue
    If collectionValue Is Nothing Then
        Exit Function
    End If
    Dim itemCount As Long
    itemCount = CLng(collectionValue.Count)
    If itemCount < 0 Then
        Exit Function
    End If
    Set GetCollection = collectionValue
    Exit Function

CollectionInvalid:
    Set GetCollection = Nothing
End Function

Private Function GetNativeArrayState(ByVal InputValue As Variant, ByRef LowerBound As Long, ByRef UpperBound As Long) As Long
    If Not TryGetPrimaryArrayBounds(InputValue, LowerBound, UpperBound) Then
        GetNativeArrayState = arrayStateUninitialized
        Exit Function
    End If

    If HasSecondArrayDimension(InputValue) Then
        GetNativeArrayState = arrayStateMultiDimensional
    Else
        GetNativeArrayState = arrayStateOneDimensional
    End If
End Function

Private Function TryGetPrimaryArrayBounds(ByVal InputValue As Variant, ByRef LowerBound As Long, ByRef UpperBound As Long) As Boolean
    On Error GoTo NotAllocated
    LowerBound = LBound(InputValue, 1)
    UpperBound = UBound(InputValue, 1)
    TryGetPrimaryArrayBounds = True
    Exit Function

NotAllocated:
    Err.Clear
    TryGetPrimaryArrayBounds = False
End Function

Private Function HasSecondArrayDimension(ByVal InputValue As Variant) As Boolean
    On Error GoTo NoSecondDimension
    HasSecondArrayDimension = (LBound(InputValue, 2) = LBound(InputValue, 2))
    On Error GoTo 0
    Exit Function

NoSecondDimension:
    Err.Clear
    HasSecondArrayDimension = False
    On Error GoTo 0
End Function

Private Sub ReadArrayValue(ByVal InputValue As Variant, ByVal Index As Long, ByVal Source As String, ByRef OutputValue As Variant)
    On Error GoTo ArrayUnavailable
    If IsObject(InputValue(Index)) Then
        Set OutputValue = InputValue(Index)
    Else
        OutputValue = InputValue(Index)
    End If
    Exit Sub

ArrayUnavailable:
    Err.Raise vbObjectError + 2299, Source, "Native array element access failed."
End Sub

Private Sub ReadCollectionValue(ByVal CollectionValue As Object, ByVal Index As Long, ByVal Source As String, ByRef OutputValue As Variant)
    On Error GoTo CollectionUnavailable
    If CollectionValue Is Nothing Then
        Err.Raise vbObjectError + 2200, Source, "Collection item access failed."
    End If
    If IsObject(CollectionValue.Item(Index)) Then
        Set OutputValue = CollectionValue.Item(Index)
    Else
        OutputValue = CollectionValue.Item(Index)
    End If
    Exit Sub

CollectionUnavailable:
    Err.Raise vbObjectError + 2200, Source, "Collection item access failed."
End Sub

Private Function AppendArrayPath(ByVal BasePath As String, ByVal LogicalIndex As Long) As String
    AppendArrayPath = BasePath & "[" & CStr(LogicalIndex) & "]"
End Function

Private Sub AddIssue(ByVal Issues As Collection, ByVal Path As String, ByVal Code As String, ByVal Message As String, ByVal Expected As String, ByVal Received As String)
    If Issues Is Nothing Then
        Err.Raise vbObjectError + 2299, "VSchema.AddIssue", "Issue collection is required."
    End If

    Dim issue As Object
    Set issue = NewIssueDictionary("VSchema.AddIssue")
    issue.Add "path", Path
    issue.Add "code", Code
    issue.Add "message", Message
    issue.Add "expected", Expected
    issue.Add "received", Received
    Issues.Add issue
End Sub

Private Function NewDictionary(ByVal Source As String) As Object
    On Error GoTo DictionaryUnavailable

    Dim dictionary As Object
    Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.CompareMode = vbBinaryCompare
    Set NewDictionary = dictionary
    Exit Function

DictionaryUnavailable:
    Err.Raise vbObjectError + 2200, Source, "Scripting.Dictionary is required for VBA-Schema."
End Function

Private Function NewIssueDictionary(ByVal Source As String) As Object
    Set NewIssueDictionary = NewDictionary(Source)
End Function

Private Function GetDictionary(ByVal InputValue As Variant, ByVal Source As String) As Object
    If Not IsObject(InputValue) Then
        Exit Function
    End If
    If StrComp(TypeName(InputValue), "Dictionary", vbBinaryCompare) <> 0 Then
        Exit Function
    End If

    On Error GoTo DictionaryInvalid

    Dim dictionary As Object
    Set dictionary = InputValue
    Dim itemCount As Long
    If dictionary Is Nothing Then
        Exit Function
    End If
    itemCount = CLng(dictionary.Count)
    Dim probe As Boolean
    probe = dictionary.Exists("__vba_schema_probe__")
    If probe Then itemCount = itemCount
    Dim keys As Variant
    keys = dictionary.Keys
    If itemCount < 0 Then
        Err.Raise vbObjectError + 2200, Source, "Scripting.Dictionary returned an invalid Count."
    End If
    Set GetDictionary = dictionary
    Exit Function

DictionaryInvalid:
    Set GetDictionary = Nothing
End Function

Private Function TryFindDictionaryKey(ByVal dictionary As Object, ByVal FieldName As String, ByRef MatchedKey As Variant) As Boolean
    If dictionary.Count = 0 Then
        Exit Function
    End If

    Dim keys As Variant
    keys = dictionary.Keys
    If Not IsArray(keys) Then
        Exit Function
    End If
    Dim index As Long
    For index = LBound(keys) To UBound(keys)
        If VarType(keys(index)) = vbString Then
            If StrComp(CStr(keys(index)), FieldName, vbBinaryCompare) = 0 Then
                MatchedKey = keys(index)
                TryFindDictionaryKey = True
                Exit Function
            End If
        End If
    Next index
End Function

Private Sub ReadDictionaryValue(ByVal dictionary As Object, ByVal Key As Variant, ByVal Source As String, ByRef OutputValue As Variant)
    On Error GoTo DictionaryUnavailable

    If IsObject(dictionary.Item(Key)) Then
        Set OutputValue = dictionary.Item(Key)
    Else
        OutputValue = dictionary.Item(Key)
    End If
    Exit Sub

DictionaryUnavailable:
    Err.Raise vbObjectError + 2200, Source, "Scripting.Dictionary item access failed."
End Sub

Private Sub CollectSpecialKeys(ByVal dictionary As Object, ByVal InvalidKeys As Collection, ByVal UnknownKeys As Collection)
    If dictionary.Count = 0 Then
        Exit Sub
    End If

    Dim keys As Variant
    keys = dictionary.Keys
    If Not IsArray(keys) Then
        Exit Sub
    End If
    Dim index As Long
    Dim keyValue As Variant
    For index = LBound(keys) To UBound(keys)
        keyValue = keys(index)
        If VarType(keyValue) <> vbString Then
            InvalidKeys.Add DescribeValue(keyValue)
        ElseIf Not mFields.Exists(CStr(keyValue)) Then
            UnknownKeys.Add CStr(keyValue)
        End If
    Next index
End Sub

Private Function SortedStrings(ByVal Values As Collection) As Collection
    Dim sorted As Collection
    Set sorted = New Collection

    Dim index As Long
    Dim value As String
    Dim position As Long
    Dim inserted As Boolean
    For index = 1 To Values.Count
        value = CStr(Values.Item(index))
        inserted = False
        For position = 1 To sorted.Count
            If StrComp(value, CStr(sorted.Item(position)), vbBinaryCompare) < 0 Then
                sorted.Add value, , position
                inserted = True
                Exit For
            End If
        Next position
        If Not inserted Then
            sorted.Add value
        End If
    Next index

    Set SortedStrings = sorted
End Function

Private Function ExpectedForKind() As String
    Select Case mKind
    Case schemaKindAny
        ExpectedForKind = "Any"
    Case schemaKindText
        ExpectedForKind = "Text"
    Case schemaKindNumber
        ExpectedForKind = "Number"
    Case schemaKindBool
        ExpectedForKind = "Boolean"
    Case schemaKindDateTime
        ExpectedForKind = "Date"
    Case schemaKindObject
        ExpectedForKind = "Object(Dictionary)"
    Case schemaKindArray
        ExpectedForKind = "Array"
    Case schemaKindLiteral
        ExpectedForKind = "Literal"
    Case schemaKindEnum
        ExpectedForKind = "Enum"
    Case schemaKindUnion
        ExpectedForKind = "Union"
    Case Else
        ExpectedForKind = "Schema"
    End Select
End Function

Private Function DescribeValue(ByVal InputValue As Variant) As String
    If IsNull(InputValue) Then
        DescribeValue = "Null"
        Exit Function
    End If
    If IsEmpty(InputValue) Then
        DescribeValue = "Empty"
        Exit Function
    End If
    If IsError(InputValue) Then
        DescribeValue = "Error(" & ErrorNumberText(InputValue) & ")"
        Exit Function
    End If
    If IsObject(InputValue) Then
        DescribeValue = "Object(" & TypeName(InputValue) & ")"
        Exit Function
    End If
    If IsArray(InputValue) Then
        DescribeValue = "Array(" & TypeName(InputValue) & ")"
        Exit Function
    End If

    Select Case VarType(InputValue)
    Case vbString
        DescribeValue = "String(" & Chr$(34) & EscapeAndTruncate(CStr(InputValue)) & Chr$(34) & ")"
    Case vbBoolean
        If CBool(InputValue) Then
            DescribeValue = "Boolean(True)"
        Else
            DescribeValue = "Boolean(False)"
        End If
    Case vbDate
        DescribeValue = "Date(" & Format$(CDate(InputValue), "yyyy-mm-dd") & "T" & Format$(CDate(InputValue), "hh:nn:ss") & ")"
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency
        DescribeValue = "Number(" & NumberText(InputValue) & ")"
    Case Else
        DescribeValue = "Object(" & TypeName(InputValue) & ")"
    End Select
End Function

Private Function ErrorNumberText(ByVal InputValue As Variant) As String
    On Error GoTo Failed

    Dim text As String
    text = Trim$(CStr(InputValue))
    Dim separator As Long
    separator = InStrRev(text, " ")
    If separator > 0 Then
        text = Mid$(text, separator + 1)
    End If
    If Len(text) = 0 Then
        text = "0"
    End If
    ErrorNumberText = text
    Exit Function

Failed:
    ErrorNumberText = "0"
End Function

Private Function EscapeAndTruncate(ByVal InputText As String) As String
    Dim escaped As String
    escaped = Replace$(InputText, Chr$(92), Chr$(92) & Chr$(92))
    escaped = Replace$(escaped, Chr$(34), Chr$(92) & Chr$(34))
    escaped = Replace$(escaped, vbCrLf, "\r\n")
    escaped = Replace$(escaped, vbCr, "\r")
    escaped = Replace$(escaped, vbLf, "\n")
    escaped = Replace$(escaped, vbTab, "\t")
    If Len(escaped) > 80 Then
        escaped = Left$(escaped, 77) & "..."
    End If
    EscapeAndTruncate = escaped
End Function

Private Function NumberText(ByVal InputValue As Variant) As String
    On Error GoTo InvalidNumber
    NumberText = Trim$(Str$(InputValue))
    Exit Function

InvalidNumber:
    Err.Raise vbObjectError + 2100, "VSchema.NumberText", "Numeric value cannot be represented deterministically."
End Function

Private Function IsSupportedNumber(ByVal InputValue As Variant) As Boolean
    If IsObject(InputValue) Or IsArray(InputValue) Or IsNull(InputValue) Or IsEmpty(InputValue) Or IsError(InputValue) Then
        Exit Function
    End If

    Select Case VarType(InputValue)
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency
        IsSupportedNumber = True
    End Select
End Function

Private Function TryDecimal(ByVal InputValue As Variant, ByRef DecimalValue As Variant) As Boolean
    If Not IsSupportedNumber(InputValue) Then
        Exit Function
    End If

    On Error GoTo Failed
    DecimalValue = CDec(InputValue)
    Select Case VarType(InputValue)
    Case vbByte, vbInteger, vbLong
        If CLng(DecimalValue) <> CLng(InputValue) Then Exit Function
    Case vbSingle
        If CSng(DecimalValue) <> CSng(InputValue) Then Exit Function
    Case vbDouble
        If CDbl(DecimalValue) <> CDbl(InputValue) Then Exit Function
    Case vbCurrency
        If CCur(DecimalValue) <> CCur(InputValue) Then Exit Function
    End Select
    TryDecimal = True
    Exit Function

Failed:
    DecimalValue = Empty
    TryDecimal = False
End Function

Private Function CompareNumeric(ByVal LeftValue As Variant, ByVal RightValue As Variant, ByVal Source As String) As Long
    Dim leftDecimal As Variant
    Dim rightDecimal As Variant
    If Not TryDecimal(LeftValue, leftDecimal) Or Not TryDecimal(RightValue, rightDecimal) Then
        Err.Raise vbObjectError + 2100, Source, "Numeric comparison would lose precision or overflow."
    End If

    If leftDecimal < rightDecimal Then
        CompareNumeric = -1
    ElseIf leftDecimal > rightDecimal Then
        CompareNumeric = 1
    Else
        CompareNumeric = 0
    End If
End Function

Private Function NormalizeLengthBoundary(ByVal Boundary As Variant, ByVal Source As String) As Long
    Dim decimalValue As Variant
    If Not TryDecimal(Boundary, decimalValue) Then
        Err.Raise vbObjectError + 2100, Source, "Length boundary must be a non-negative whole number in Long range."
    End If
    If decimalValue < 0 Or decimalValue <> Fix(decimalValue) Then
        Err.Raise vbObjectError + 2100, Source, "Length boundary must be a non-negative whole number in Long range."
    End If

    On Error GoTo OutOfRange
    NormalizeLengthBoundary = CLng(decimalValue)
    Exit Function

OutOfRange:
    Err.Raise vbObjectError + 2100, Source, "Length boundary must be a non-negative whole number in Long range."
End Function

Private Sub ValidateNumericBoundary(ByVal Boundary As Variant, ByVal Source As String)
    Dim decimalValue As Variant
    If Not TryDecimal(Boundary, decimalValue) Then
        Err.Raise vbObjectError + 2100, Source, "Number boundary must be a supported finite numeric value."
    End If
End Sub

Private Function IsWholeNumberValue(ByVal InputValue As Variant) As Boolean
    Dim decimalValue As Variant
    If Not TryDecimal(InputValue, decimalValue) Then
        Exit Function
    End If
    IsWholeNumberValue = (decimalValue = Fix(decimalValue))
End Function

Private Function AppendObjectPath(ByVal BasePath As String, ByVal FieldName As String) As String
    If IsPathIdentifier(FieldName) Then
        AppendObjectPath = BasePath & "." & FieldName
    Else
        AppendObjectPath = BasePath & "[" & Chr$(34) & EscapePathSegment(FieldName) & Chr$(34) & "]"
    End If
End Function

Private Function IsPathIdentifier(ByVal FieldName As String) As Boolean
    Dim index As Long
    Dim code As Long
    If Len(FieldName) = 0 Then
        Exit Function
    End If

    code = AscW(Mid$(FieldName, 1, 1))
    If Not ((code >= AscW("A") And code <= AscW("Z")) Or (code >= AscW("a") And code <= AscW("z")) Or code = AscW("_")) Then
        Exit Function
    End If

    For index = 2 To Len(FieldName)
        code = AscW(Mid$(FieldName, index, 1))
        If Not ((code >= AscW("A") And code <= AscW("Z")) Or (code >= AscW("a") And code <= AscW("z")) Or (code >= AscW("0") And code <= AscW("9")) Or code = AscW("_")) Then
            Exit Function
        End If
    Next index
    IsPathIdentifier = True
End Function

Private Function EscapePathSegment(ByVal FieldName As String) As String
    Dim escaped As String
    Dim index As Long
    Dim code As Long
    Dim character As String

    For index = 1 To Len(FieldName)
        character = Mid$(FieldName, index, 1)
        code = AscW(character)
        Select Case character
        Case Chr$(92)
            escaped = escaped & Chr$(92) & Chr$(92)
        Case Chr$(34)
            escaped = escaped & Chr$(92) & Chr$(34)
        Case vbCr
            escaped = escaped & "\r"
        Case vbLf
            escaped = escaped & "\n"
        Case vbTab
            escaped = escaped & "\t"
        Case Else
            If code < 32 Then
                escaped = escaped & "\u" & Right$("0000" & Hex$(code), 4)
            Else
                escaped = escaped & character
            End If
        End Select
    Next index
    EscapePathSegment = escaped
End Function

Private Sub EnsureAcyclicGraph(ByVal ProcedureName As String)
    Dim active As Collection
    Set active = New Collection
    VisitSchemaGraph active, ProcedureName
End Sub

Private Sub VisitSchemaGraph(ByVal ActiveSchemas As Collection, ByVal ProcedureName As String)
    Dim active As Collection
    Set active = ActiveSchemas
    If active Is Nothing Then
        Err.Raise vbObjectError + 2299, ProcedureName, "Active schema collection is required."
    End If

    If mFields Is Nothing Or mFieldOrder Is Nothing Then
        If mKind = schemaKindObject Then
            Err.Raise vbObjectError + 2299, ProcedureName, "Object schema state is not initialized."
        End If
    End If

    Dim index As Long
    Dim fieldName As String
    Dim fieldSchema As VSchema

    If SchemaCollectionContains(ActiveSchemas, Me) Then
        Err.Raise vbObjectError + 2103, ProcedureName, "Schema graph contains a cycle."
    End If
    active.Add Me

    If mKind = schemaKindObject Then
        Dim objectFieldOrder As Collection
        Dim objectFields As Object
        Set objectFieldOrder = mFieldOrder
        Set objectFields = mFields
        If objectFieldOrder Is Nothing Or objectFields Is Nothing Then
            Err.Raise vbObjectError + 2299, ProcedureName, "Object schema state is not initialized."
        End If
        For index = 1 To objectFieldOrder.Count
            fieldName = CStr(objectFieldOrder.Item(index))
            Set fieldSchema = objectFields.Item(fieldName)
            Call fieldSchema.InternalVisitSchemaGraph(active, ProcedureName)
        Next index
    ElseIf mKind = schemaKindArray Then
        If mItemSchema Is Nothing Then
            Err.Raise vbObjectError + 2299, ProcedureName, "Array item schema is not initialized."
        End If
        Call mItemSchema.InternalVisitSchemaGraph(active, ProcedureName)
    ElseIf mKind = schemaKindUnion Then
        If mUnionSchemas Is Nothing Then
            Err.Raise vbObjectError + 2299, ProcedureName, "Union schema state is not initialized."
        End If
        If mUnionSchemas.Count = 0 Then
            Err.Raise vbObjectError + 2299, ProcedureName, "Union schema state is not initialized."
        End If
        For index = 1 To mUnionSchemas.Count
            Set fieldSchema = mUnionSchemas.Item(index)
            Call fieldSchema.InternalVisitSchemaGraph(active, ProcedureName)
        Next index
    End If

    If active Is Nothing Then
        Exit Sub
    End If
    active.Remove active.Count
End Sub

Private Function SchemaCollectionContains(ByVal ActiveSchemas As Collection, ByVal TargetSchema As VSchema) As Boolean
    If ActiveSchemas Is Nothing Then
        Exit Function
    End If
    Dim index As Long
    Dim activeSchema As VSchema
    For index = 1 To ActiveSchemas.Count
        Set activeSchema = ActiveSchemas.Item(index)
        If activeSchema Is TargetSchema Then
            SchemaCollectionContains = True
            Exit Function
        End If
    Next index
End Function

Private Sub EnsureSupportedKind(ByVal ProcedureName As String)
    If mKind < schemaKindAny Or mKind > schemaKindUnion Then
        RaiseNotImplemented ProcedureName
    End If
End Sub

Private Sub EnsureInitialized(ByVal ProcedureName As String)
    If Not mInitialized Then
        Err.Raise vbObjectError + 2102, ProcedureName, "Schema has not been initialized."
    End If
End Sub

Private Sub RaiseNotImplemented(ByVal ProcedureName As String)
    Err.Raise vbObjectError + 2299, ProcedureName, "VBA-Schema v1 implementation is not complete."
End Sub
