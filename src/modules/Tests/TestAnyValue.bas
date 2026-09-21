Attribute VB_Name = "TestAnyValue"
Option Explicit

Public Sub Test_AnyValue_AcceptsScalarAndSpecialVariants()
    Dim result As VValidationResult
    Dim inputValue As Variant

    Set result = Schema.AnyValue().SafeParse("value")
    XlflowAssert.AssertTrue result.Success, "AnyValue should accept String"
    XlflowAssert.AssertStrictEquals "value", result.Value
    XlflowAssert.AssertEquals 0, result.Issues.Count
    XlflowAssert.AssertStrictEquals "", result.ErrorText

    inputValue = Null
    Set result = Schema.AnyValue().SafeParse(inputValue)
    XlflowAssert.AssertTrue result.Success, "AnyValue should accept Null"
    XlflowAssert.AssertNull result.Value

    inputValue = Empty
    Set result = Schema.AnyValue().SafeParse(inputValue)
    XlflowAssert.AssertTrue result.Success, "AnyValue should accept Empty"
    XlflowAssert.AssertEmpty result.Value

    inputValue = CVErr(2042)
    Set result = Schema.AnyValue().SafeParse(inputValue)
    XlflowAssert.AssertTrue result.Success, "AnyValue should accept Error Variant"
    XlflowAssert.AssertTrue IsError(result.Value), "AnyValue should preserve Error Variant"
End Sub

Public Sub Test_AnyValue_PreservesObjectAndArrayReferences()
    Dim dictionary As Object
    Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.Add "name", "vba"

    Dim result As VValidationResult
    Set result = Schema.AnyValue().SafeParse(dictionary)
    XlflowAssert.AssertTrue result.Success, "AnyValue should accept objects"

    Dim output As Object
    Set output = result.Value
    XlflowAssert.AssertSame dictionary, output

    Dim values(0 To 1) As Long
    values(0) = 3
    values(1) = 5
    Set result = Schema.AnyValue().SafeParse(values)
    XlflowAssert.AssertTrue result.Success, "AnyValue should accept native arrays"
    XlflowAssert.AssertTrue IsArray(result.Value), "AnyValue should preserve native array"

    Dim items As Collection
    Set items = New Collection
    items.Add "item"
    Set result = Schema.AnyValue().SafeParse(items)
    XlflowAssert.AssertTrue result.Success, "AnyValue should accept Collection objects"

    Dim nothingObject As Object
    Set nothingObject = Nothing
    Set result = Schema.AnyValue().SafeParse(nothingObject)
    XlflowAssert.AssertTrue result.Success, "AnyValue should accept Nothing object references"
    Dim outputNothing As Object
    Set outputNothing = result.Value
    XlflowAssert.AssertIsNothing outputNothing
End Sub

Public Sub Test_AnyValue_UsesRootPathForFailureFreeResult()
    Dim result As VValidationResult
    Set result = Schema.AnyValue().SafeParse("root")
    XlflowAssert.AssertStrictEquals "", result.ErrorText
    XlflowAssert.AssertEquals 0, result.Issues.Count
End Sub
