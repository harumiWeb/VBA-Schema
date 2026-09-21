Attribute VB_Name = "Schema"
Option Explicit

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

' InternalInitialize accepts an encoded kind so ordinary direct calls fail.
Private Const internalKindKey As Long = 27183

Public Function AnyValue() As VSchema
    Set AnyValue = CreateSchema(schemaKindAny)
End Function

Public Function Text() As VSchema
    Set Text = CreateSchema(schemaKindText)
End Function

Public Function Number() As VSchema
    Set Number = CreateSchema(schemaKindNumber)
End Function

Public Function Bool() As VSchema
    Set Bool = CreateSchema(schemaKindBool)
End Function

Public Function DateTime() As VSchema
    Set DateTime = CreateSchema(schemaKindDateTime)
End Function

Public Function ObjectSchema() As VSchema
    Set ObjectSchema = CreateSchema(schemaKindObject)
End Function

Public Function ArrayOf(ByVal ItemSchema As VSchema) As VSchema
    Dim newSchema As VSchema
    Set newSchema = CreateSchema(schemaKindArray)
    Set ArrayOf = newSchema.InternalSetItemSchema(ItemSchema)
End Function

Public Function Literal(ByVal ExpectedValue As Variant) As VSchema
    Dim newSchema As VSchema
    Set newSchema = CreateSchema(schemaKindLiteral)
    Set Literal = newSchema.InternalSetLiteral(ExpectedValue)
End Function

Public Function EnumOf(ByVal Values As Variant) As VSchema
    Dim newSchema As VSchema
    Set newSchema = CreateSchema(schemaKindEnum)
    Set EnumOf = newSchema.InternalSetEnum(Values)
End Function

Public Function UnionOf(ByVal Schemas As Collection) As VSchema
    Dim newSchema As VSchema
    Set newSchema = CreateSchema(schemaKindUnion)
    Set UnionOf = newSchema.InternalSetUnion(Schemas)
End Function

Private Function CreateSchema(ByVal KindCode As Long) As VSchema
    Dim newSchema As VSchema
    Set newSchema = New VSchema
    Set CreateSchema = newSchema.InternalInitialize(EncodeKind(KindCode))
End Function

Private Function EncodeKind(ByVal KindCode As Long) As Long
    EncodeKind = KindCode + internalKindKey
End Function
