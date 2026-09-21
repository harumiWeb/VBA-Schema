Attribute VB_Name = "PublicApiCompile"
Option Explicit

' This procedure is compile-only. It intentionally is not named Test*.
Public Sub CompilePublicApi()
    Dim scalarSchema As VSchema
    Dim objectSchema As VSchema
    Dim arraySchema As VSchema
    Dim literalSchema As VSchema
    Dim enumSchema As VSchema
    Dim unionSchema As VSchema
    Dim branches As Collection
    Dim result As VValidationResult
    Dim typedValues() As String
    Dim lateBoundObject As Object
    Dim scalarValue As Variant
    Dim objectValue As Object
    Dim success As Boolean
    Dim issues As Collection
    Dim errorText As String

    If Not CompileOnlyGuard() Then Exit Sub

    Set scalarSchema = Schema.Text().Min(1).Max(100).Length(10).Pattern("^[A-Z]+$").Email().OptionalField().Nullable()
    Set scalarSchema = Schema.Number().WholeNumber().Min(0).Max(100)
    Set scalarSchema = Schema.Bool().Nullable()
    Set scalarSchema = Schema.DateTime()

    Set objectSchema = Schema.ObjectSchema() _
        .Field("name", Schema.Text().Min(1).Max(100)) _
        .Field("age", Schema.Number().WholeNumber().Min(0).OptionalField()) _
        .Field("enabled", Schema.Bool().Nullable()) _
        .Field("createdAt", Schema.DateTime()) _
        .Field("tags", Schema.ArrayOf(Schema.Text()).Length(3)) _
        .Strict()

    Set arraySchema = Schema.ArrayOf(Schema.Text())
    Set literalSchema = Schema.Literal("active")
    Set enumSchema = Schema.EnumOf(Array("pending", "active", "disabled"))

    Set branches = New Collection
    branches.Add Schema.Text()
    branches.Add Schema.Number()
    Set unionSchema = Schema.UnionOf(branches)

    ReDim typedValues(0 To 0)
    Set result = arraySchema.SafeParse(typedValues)

    Set lateBoundObject = CreateObject("Scripting.Dictionary")
    Set result = objectSchema.SafeParse(lateBoundObject)

    scalarValue = result.Value
    Set objectValue = result.Value
    success = result.Success
    Set issues = result.Issues
    errorText = result.ErrorText

    KeepCompileReferences scalarSchema, literalSchema, enumSchema, unionSchema, scalarValue, objectValue, success, issues, errorText

    CompileInternalInitialize
    CompileInternalHooks
End Sub

Private Sub CompileInternalInitialize()
    Dim directSchema As VSchema
    Set directSchema = New VSchema
    Set directSchema = directSchema.InternalInitialize(1)
End Sub

Private Sub CompileInternalHooks()
    Dim directSchema As VSchema
    Dim issues As Collection
    Dim activeSchemas As Collection
    Set directSchema = New VSchema
    Set directSchema = directSchema.InternalInitialize(27183)
    Set issues = New Collection
    Set activeSchemas = New Collection
    Set directSchema = directSchema.InternalEnsureComposed("PublicApiCompile")
    Set directSchema = directSchema.InternalValidateAtPath("compile", "$", issues)
    Set directSchema = directSchema.InternalValidateMissing("$", issues)
    Set directSchema = directSchema.InternalVisitSchemaGraph(activeSchemas, "PublicApiCompile")
End Sub

Private Sub KeepCompileReferences(ByVal scalarSchema As VSchema, ByVal literalSchema As VSchema, ByVal enumSchema As VSchema, ByVal unionSchema As VSchema, ByVal scalarValue As Variant, ByVal objectValue As Object, ByVal success As Boolean, ByVal issues As Collection, ByVal errorText As String)
    If scalarSchema Is Nothing Then Exit Sub
    If literalSchema Is Nothing Then Exit Sub
    If enumSchema Is Nothing Then Exit Sub
    If unionSchema Is Nothing Then Exit Sub
    If IsEmpty(scalarValue) Then Exit Sub
    If objectValue Is Nothing Then Exit Sub
    If Not success Then Exit Sub
    If issues Is Nothing Then Exit Sub
    If Len(errorText) < 0 Then Exit Sub
End Sub

Private Function CompileOnlyGuard() As Boolean
    CompileOnlyGuard = False
End Function
