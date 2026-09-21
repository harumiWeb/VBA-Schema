Attribute VB_Name = "SampleOrderImport"
Option Explicit

Public Sub RunOrderImportSample()
    Dim orderSchema As VSchema
    Set orderSchema = BuildOrderSchema()

    Dim validOrder As Object
    Dim validResult As VValidationResult
    Set validOrder = BuildValidOrder()
    Set validResult = orderSchema.SafeParse(validOrder)
    PrintResult "valid order", validResult

    Dim invalidOrder As Object
    Dim invalidResult As VValidationResult
    Set invalidOrder = BuildInvalidOrder()
    Set invalidResult = orderSchema.SafeParse(invalidOrder)
    PrintResult "invalid order", invalidResult
End Sub

Private Function BuildOrderSchema() As VSchema
    Dim statuses As Variant
    statuses = Array("pending", "paid", "cancelled")

    Dim itemSchema As VSchema
    Set itemSchema = Schema.ObjectSchema() _
        .Field("sku", Schema.Text().Pattern("[A-Z]{2}-[0-9]{4}")) _
        .Field("quantity", Schema.Number().WholeNumber().Min(1).Max(999)) _
        .Field("unitPrice", Schema.Number().Min(0))

    Set BuildOrderSchema = Schema.ObjectSchema() _
        .Field("orderId", Schema.Text().Pattern("ORD-[0-9]{6}")) _
        .Field("email", Schema.Text().Email()) _
        .Field("status", Schema.EnumOf(statuses)) _
        .Field("createdAt", Schema.DateTime()) _
        .Field("items", Schema.ArrayOf(itemSchema).Min(1).Max(50)) _
        .Field("note", Schema.Text().Max(200).OptionalField()) _
        .Strict()
End Function

Private Function BuildValidOrder() As Object
    Dim order As Object
    Set order = NewDictionary()
    order.Add "orderId", "ORD-202609"
    order.Add "email", "buyer@example.com"
    order.Add "status", "paid"
    order.Add "createdAt", DateSerial(2026, 9, 21) + TimeSerial(9, 30, 0)
    order.Add "note", "Deliver after 18:00"

    Dim items As Collection
    Set items = New Collection

    Dim item As Object
    Set item = BuildItem("VB-0001", 2, 1980#)
    items.Add item
    order.Add "items", items

    Set BuildValidOrder = order
End Function

Private Function BuildInvalidOrder() As Object
    Dim order As Object
    Set order = NewDictionary()
    order.Add "orderId", "ORD-ABC123"
    order.Add "email", "buyer.example.com"
    order.Add "status", "refunded"
    order.Add "createdAt", "2026-09-21"

    Dim items As Collection
    Set items = New Collection

    Dim item As Object
    Set item = BuildItem("broken", 0, -1)
    items.Add item
    order.Add "items", items
    order.Add "debug", True

    Set BuildInvalidOrder = order
End Function

Private Function BuildItem(ByVal Sku As String, ByVal Quantity As Long, ByVal UnitPrice As Double) As Object
    Dim item As Object
    Set item = NewDictionary()
    item.Add "sku", Sku
    item.Add "quantity", Quantity
    item.Add "unitPrice", UnitPrice
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
