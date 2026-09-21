Attribute VB_Name = "TestArray"
Option Explicit

Public Sub Test_Array_AcceptsNativeArraysAndCollections()
    Dim arraySchema As VSchema
    Set arraySchema = Schema.ArrayOf(Schema.Text())

    Dim zeroBased(0 To 1) As String
    zeroBased(0) = "a"
    zeroBased(1) = "b"
    Dim result As VValidationResult
    Set result = arraySchema.SafeParse(zeroBased)
    XlflowAssert.AssertTrue result.Success

    Dim oneBased(1 To 2) As String
    oneBased(1) = "c"
    oneBased(2) = "d"
    Set result = arraySchema.SafeParse(oneBased)
    XlflowAssert.AssertTrue result.Success

    Dim negativeBounded(-2 To - 1) As String
    negativeBounded(-2) = "g"
    negativeBounded(-1) = "h"
    Set result = arraySchema.SafeParse(negativeBounded)
    XlflowAssert.AssertTrue result.Success

    Dim items As Collection
    Set items = New Collection
    items.Add "e"
    items.Add "f"
    Set result = arraySchema.SafeParse(items)
    XlflowAssert.AssertTrue result.Success

    Dim typedObjects(0 To 0) As Object
    Dim objectValue As Object
    Set objectValue = CreateObject("Scripting.Dictionary")
    objectValue.Add "id", 1
    Set typedObjects(0) = objectValue
    Dim objectArraySchema As VSchema
    Set objectArraySchema = Schema.ArrayOf(Schema.AnyValue())
    Set result = objectArraySchema.SafeParse(typedObjects)
    XlflowAssert.AssertTrue result.Success
End Sub

Public Sub Test_Array_UsesLogicalIndexAndNestedValidation()
    Dim arraySchema As VSchema
    Set arraySchema = Schema.ArrayOf(Schema.Text().Length(2))

    Dim values(5 To 6) As String
    values(5) = "ok"
    values(6) = "x"

    Dim result As VValidationResult
    Set result = arraySchema.SafeParse(values)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count

    Dim issue As Object
    Set issue = result.Issues.Item(1)
    If issue Is Nothing Then
        Err.Raise vbObjectError + 2299, "TestArray", "Issue is required."
    End If
    XlflowAssert.AssertStrictEquals "$[1]", issue.Item("path")
    XlflowAssert.AssertStrictEquals "invalid_length", issue.Item("code")
End Sub

Public Sub Test_Array_UsesCollectionAndLengthConstraints()
    Dim arraySchema As VSchema
    Set arraySchema = Schema.ArrayOf(Schema.Number()).Length(2).Min(1).Max(3)

    Dim items As Collection
    Set items = New Collection
    items.Add 1
    items.Add 2

    Dim result As VValidationResult
    Set result = arraySchema.SafeParse(items)
    XlflowAssert.AssertTrue result.Success
    If items Is Nothing Then
        Err.Raise vbObjectError + 2299, "TestArray", "Collection is required."
    End If
    XlflowAssert.AssertEquals 2, items.Count
    items.Add 3
    items.Add 4
    Set result = arraySchema.SafeParse(items)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertStrictEquals "invalid_length", result.Issues.Item(1).Item("code")

    Dim emptyItems As Collection
    Set emptyItems = New Collection
    Set result = Schema.ArrayOf(Schema.Number()).Length(0).SafeParse(emptyItems)
    XlflowAssert.AssertTrue result.Success
End Sub

Public Sub Test_Array_HandlesUninitializedAndMultidimensionalArrays()
    Dim arraySchema As VSchema
    Set arraySchema = Schema.ArrayOf(Schema.Text()).Length(0)

    Dim uninitialized() As String
    Dim result As VValidationResult
    Set result = arraySchema.SafeParse(uninitialized)
    XlflowAssert.AssertTrue result.Success

    Dim matrix(0 To 0, 0 To 0) As String
    Set result = arraySchema.SafeParse(matrix)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertStrictEquals "invalid_array_rank", result.Issues.Item(1).Item("code")

    Dim dictionary As Object
    Set dictionary = CreateObject("Scripting.Dictionary")
    Set result = arraySchema.SafeParse(dictionary)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "invalid_type", result.Issues.Item(1).Item("code")
End Sub

Public Sub Test_Array_PreservesObjectAndNothingElements()
    Dim arraySchema As VSchema
    Set arraySchema = Schema.ArrayOf(Schema.AnyValue())

    Dim values(0 To 1) As Variant
    Dim dictionary As Object
    Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.Add "id", 1
    Set values(0) = dictionary
    Dim nothingObject As Object
    Set nothingObject = Nothing
    Set values(1) = nothingObject

    Dim result As VValidationResult
    Set result = arraySchema.SafeParse(values)
    XlflowAssert.AssertTrue result.Success
End Sub

Public Sub Test_Array_SupportsNestedSequences()
    Dim inner As Collection
    Set inner = New Collection
    inner.Add "ok"

    Dim outer As Collection
    Set outer = New Collection
    outer.Add inner

    Dim nestedSchema As VSchema
    Set nestedSchema = Schema.ArrayOf(Schema.ArrayOf(Schema.Text()))
    Dim result As VValidationResult
    Set result = nestedSchema.SafeParse(outer)
    XlflowAssert.AssertTrue result.Success

    inner.Add 42
    Set result = nestedSchema.SafeParse(outer)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertStrictEquals "$[0][1]", result.Issues.Item(1).Item("path")
End Sub

Public Sub Test_Array_CollectionPreservesSpecialElements()
    Dim items As Collection
    Set items = New Collection

    Dim objectValue As Object
    Set objectValue = CreateObject("Scripting.Dictionary")
    objectValue.Add "id", 1
    items.Add objectValue

    Dim errorValue As Variant
    errorValue = CVErr(2042)
    items.Add errorValue

    Dim nothingValue As Variant
    Set nothingValue = Nothing
    items.Add nothingValue

    Dim result As VValidationResult
    Set result = Schema.ArrayOf(Schema.AnyValue()).SafeParse(items)
    XlflowAssert.AssertTrue result.Success
End Sub
