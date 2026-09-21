Attribute VB_Name = "SampleSettingsValidation"
Option Explicit

Public Sub RunSettingsValidationSample()
    Dim settingsSchema As VSchema
    Set settingsSchema = BuildSettingsSchema()

    Dim validSettings As Object
    Dim validResult As VValidationResult
    Set validSettings = BuildValidSettings()
    Set validResult = settingsSchema.SafeParse(validSettings)
    PrintResult "valid settings", validResult

    Dim invalidSettings As Object
    Dim invalidResult As VValidationResult
    Set invalidSettings = BuildInvalidSettings()
    Set invalidResult = settingsSchema.SafeParse(invalidSettings)
    PrintResult "invalid settings", invalidResult
End Sub

Private Function BuildSettingsSchema() As VSchema
    Dim environments As Variant
    environments = Array("development", "production")

    Set BuildSettingsSchema = Schema.ObjectSchema() _
        .Field("environment", Schema.EnumOf(environments)) _
        .Field("apiBaseUrl", Schema.Text().Pattern("https://[A-Za-z0-9.-]+")) _
        .Field("port", Schema.Number().WholeNumber().Min(1).Max(65535)) _
        .Field("retryCount", Schema.Number().WholeNumber().Min(0).Max(5).OptionalField()) _
        .Field("proxy", Schema.Text().Nullable().OptionalField()) _
        .Field("tags", Schema.ArrayOf(Schema.Text().Length(1).Max(20)).OptionalField()) _
        .Strict()
End Function

Private Function BuildValidSettings() As Object
    Dim settings As Object
    Set settings = NewDictionary()
    settings.Add "environment", "development"
    settings.Add "apiBaseUrl", "https://api.example.com"
    settings.Add "port", 443
    settings.Add "retryCount", 2
    settings.Add "proxy", Null
    settings.Add "tags", Array("vba", "validation")
    Set BuildValidSettings = settings
End Function

Private Function BuildInvalidSettings() As Object
    Dim settings As Object
    Set settings = NewDictionary()
    settings.Add "environment", "staging"
    settings.Add "apiBaseUrl", "http://api.example.com"
    settings.Add "port", 0
    settings.Add "tags", Array("ok", 17)
    settings.Add "debugMode", True
    Set BuildInvalidSettings = settings
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
