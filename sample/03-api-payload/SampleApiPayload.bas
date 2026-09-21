Attribute VB_Name = "SampleApiPayload"
Option Explicit

Public Sub RunApiPayloadSample()
    Dim responseSchema As VSchema
    Set responseSchema = BuildResponseSchema()

    Dim validPayload As Object
    Dim validResult As VValidationResult
    Set validPayload = BuildValidPayload()
    Set validResult = responseSchema.SafeParse(validPayload)
    PrintResult "valid response", validResult

    Dim invalidPayload As Object
    Dim invalidResult As VValidationResult
    Set invalidPayload = BuildInvalidPayload()
    Set invalidResult = responseSchema.SafeParse(invalidPayload)
    PrintResult "invalid response", invalidResult
End Sub

Private Function BuildResponseSchema() As VSchema
    Dim contactBranches As Collection
    Set contactBranches = New Collection
    contactBranches.Add Schema.Text().Email()
    contactBranches.Add Schema.Number().WholeNumber().Min(1)

    Dim itemSchema As VSchema
    Set itemSchema = Schema.ObjectSchema() _
        .Field("id", Schema.Number().WholeNumber().Min(1)) _
        .Field("name", Schema.Text().Length(1).Max(100)) _
        .Field("active", Schema.Bool())

    Set BuildResponseSchema = Schema.ObjectSchema() _
        .Field("status", Schema.Literal("ok")) _
        .Field("requestId", Schema.Text().Length(36)) _
        .Field("contact", Schema.UnionOf(contactBranches)) _
        .Field("items", Schema.ArrayOf(itemSchema).Min(1)) _
        .Field("nextPage", Schema.Number().WholeNumber().Nullable().OptionalField()) _
        .Strict()
End Function

Private Function BuildValidPayload() As Object
    Dim payload As Object
    Set payload = NewDictionary()
    payload.Add "status", "ok"
    payload.Add "requestId", "123e4567-e89b-12d3-a456-426614174000"
    payload.Add "contact", "alice@example.com"
    payload.Add "nextPage", Null

    Dim items As Collection
    Set items = New Collection

    Dim item As Object
    Set item = BuildItem(1001, "VBA-Schema", True)
    items.Add item
    payload.Add "items", items

    Set BuildValidPayload = payload
End Function

Private Function BuildInvalidPayload() As Object
    Dim payload As Object
    Set payload = NewDictionary()
    payload.Add "status", "error"
    payload.Add "requestId", "short-id"
    payload.Add "contact", True

    Dim items As Collection
    Set items = New Collection

    Dim item As Object
    Set item = BuildItem(0, "", "yes")
    items.Add item
    payload.Add "items", items
    payload.Add "traceId", "debug-only"

    Set BuildInvalidPayload = payload
End Function

Private Function BuildItem(ByVal ItemId As Long, ByVal ItemName As String, ByVal IsActive As Variant) As Object
    Dim item As Object
    Set item = NewDictionary()
    item.Add "id", ItemId
    item.Add "name", ItemName
    item.Add "active", IsActive
    Set BuildItem = item
End Function

Private Function NewDictionary() As Object
    Dim dictionary As Object
    Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.CompareMode = vbBinaryCompare
    Set NewDictionary = dictionary
End Function

Private Sub PrintResult(ByVal Label As String, ByVal Result As VValidationResult)
    If Result.Success Then
        Debug.Print Label & ": OK"
    Else
        Debug.Print Label & ": FAIL"
        Debug.Print Result.ErrorText
    End If
End Sub
