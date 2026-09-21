Attribute VB_Name = "TestResult"
Option Explicit

Public Sub Test_Result_SuccessAndFailureInvariants()
    Dim result As VValidationResult
    Set result = Schema.Text().SafeParse("ok")
    XlflowAssert.AssertTrue result.Success
    XlflowAssert.AssertStrictEquals "ok", result.Value
    XlflowAssert.AssertEquals 0, result.Issues.Count
    XlflowAssert.AssertStrictEquals "", result.ErrorText

    Set result = Schema.Text().SafeParse(42)
    XlflowAssert.AssertFalse result.Success
    XlflowAssert.AssertEmpty result.Value
    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertTrue Len(result.ErrorText) > 0
End Sub

Public Sub Test_Result_IssuesAreSnapshots()
    Dim result As VValidationResult
    Set result = Schema.Text().Min(3).SafeParse("x")

    Dim firstSnapshot As Collection
    Set firstSnapshot = result.Issues
    Dim firstIssue As Object
    Set firstIssue = firstSnapshot.Item(1)
    firstIssue.Item("message") = "changed"
    firstSnapshot.Remove 1

    XlflowAssert.AssertEquals 1, result.Issues.Count
    XlflowAssert.AssertContains "Text length is below the minimum", result.ErrorText
End Sub

Public Sub Test_Result_PreservesSuccessfulObjectValue()
    Dim inputObject As Object
    Set inputObject = CreateObject("Scripting.Dictionary")
    inputObject.Add "id", 1

    Dim result As VValidationResult
    Set result = Schema.AnyValue().SafeParse(inputObject)

    Dim outputObject As Object
    Set outputObject = result.Value
    XlflowAssert.AssertSame inputObject, outputObject
End Sub
