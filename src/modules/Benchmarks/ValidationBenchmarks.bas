Attribute VB_Name = "ValidationBenchmarks"
Option Explicit

Private Const benchmarkWarmupIterations As Long = 1
Private Const benchmarkIterations As Long = 5
Private Const objectFieldCount As Long = 1000
Private Const objectPassesPerIteration As Long = 10
Private Const scalarValidationCount As Long = 10000

''' Runs the Windows 64-bit runtime benchmark and emits the JSON result to xlflow.
'''
''' The benchmark module is development-only and is excluded from the release payload.
Public Sub RunValidationBenchmark()
    Dim objectSchema As VSchema
    Set objectSchema = BuildObjectSchema()
    Dim objectSuccessInput As Object
    Set objectSuccessInput = BuildObjectInput(True)
    Dim objectIssueInput As Object
    Set objectIssueInput = BuildObjectInput(False)

    Dim scalarSchema As VSchema
    Set scalarSchema = BuildScalarSchema()

    Dim objectSuccessJson As String
    MeasureObjectFixture "object_1000_fields_success", 500#, objectSchema, objectSuccessInput, True, objectSuccessJson

    Dim objectIssueJson As String
    MeasureObjectFixture "object_1000_fields_issues", 0#, objectSchema, objectIssueInput, False, objectIssueJson

    Dim scalarSuccessJson As String
    MeasureScalarFixture "scalar_10000_success", 1000#, scalarSchema, 123#, True, scalarSuccessJson

    Dim scalarIssueJson As String
    MeasureScalarFixture "scalar_10000_issues", 0#, scalarSchema, 123.5, False, scalarIssueJson

    Dim benchmarkJson As String
    benchmarkJson = "{" & _
        """schema_version"":1," & _
        """generated_at"":""" & JsonEscape(Format$(Now, "yyyy-mm-dd\Thh:nn:ss")) & """," & _
        """machine"":""" & JsonEscape(Environ$("COMPUTERNAME")) & """," & _
        """office_bitness"":""" & OfficeBitness() & """," & _
        """excel_version"":""" & JsonEscape(Application.Version) & """," & _
        """excel_build"":""unknown""," & _
        """command"":""ValidationBenchmarks.RunValidationBenchmark""," & _
        """warmup_iterations"":" & CStr(benchmarkWarmupIterations) & "," & _
        """iterations"":" & CStr(benchmarkIterations) & "," & _
        """fixtures"": [" & objectSuccessJson & "," & objectIssueJson & "," & scalarSuccessJson & "," & scalarIssueJson & "]" & _
        "}"

    XlflowDebug.Log "VBA_SCHEMA_BENCHMARK_JSON=" & benchmarkJson
End Sub

Private Function BuildObjectSchema() As VSchema
    Dim objectSchema As VSchema
    Set objectSchema = Schema.ObjectSchema()

    Dim index As Long
    Dim fieldSchema As VSchema
    For index = 0 To objectFieldCount - 1
        Set fieldSchema = Schema.Text().Length(5)
        Call objectSchema.Field("field_" & CStr(index), fieldSchema)
    Next index

    Set BuildObjectSchema = objectSchema
End Function

Private Function BuildObjectInput(ByVal IsValid As Boolean) As Object
    Dim dictionary As Object
    Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.CompareMode = vbBinaryCompare

    Dim index As Long
    Dim fieldName As String
    For index = 0 To objectFieldCount - 1
        fieldName = "field_" & CStr(index)
        If IsValid Or index Mod 10 <> 0 Then
            dictionary.Add fieldName, "value"
        Else
            dictionary.Add fieldName, index
        End If
    Next index

    Set BuildObjectInput = dictionary
End Function

Private Function BuildScalarSchema() As VSchema
    Set BuildScalarSchema = Schema.Number().WholeNumber().Min(0).Max(1000000)
End Function

Private Sub MeasureObjectFixture( _
        ByVal FixtureId As String, _
        ByVal TargetMilliseconds As Double, _
        ByVal SourceSchema As VSchema, _
        ByVal InputObject As Object, _
        ByVal ExpectedSuccess As Boolean, _
        ByRef OutputJson As String)

    Dim warmupIndex As Long
    For warmupIndex = 1 To benchmarkWarmupIterations
        Call RunObjectPass(SourceSchema, InputObject, ExpectedSuccess)
    Next warmupIndex

    Dim samples(1 To benchmarkIterations) As Double
    Dim totalIssues As Long
    Dim iterationIndex As Long
    Dim passIndex As Long
    Dim startedAt As Double
    For iterationIndex = 1 To benchmarkIterations
        startedAt = CDbl(Timer)
        For passIndex = 1 To objectPassesPerIteration
            totalIssues = totalIssues + RunObjectPass(SourceSchema, InputObject, ExpectedSuccess)
        Next passIndex
        samples(iterationIndex) = ElapsedMilliseconds(startedAt)
    Next iterationIndex

    OutputJson = BuildFixtureJson(FixtureId, TargetMilliseconds, objectFieldCount * benchmarkIterations * objectPassesPerIteration, totalIssues, samples)
End Sub

Private Sub MeasureScalarFixture( _
        ByVal FixtureId As String, _
        ByVal TargetMilliseconds As Double, _
        ByVal SourceSchema As VSchema, _
        ByVal InputValue As Variant, _
        ByVal ExpectedSuccess As Boolean, _
        ByRef OutputJson As String)

    Dim warmupIndex As Long
    For warmupIndex = 1 To benchmarkWarmupIterations
        Call RunScalarPass(SourceSchema, InputValue, ExpectedSuccess)
    Next warmupIndex

    Dim samples(1 To benchmarkIterations) As Double
    Dim totalIssues As Long
    Dim iterationIndex As Long
    Dim startedAt As Double
    For iterationIndex = 1 To benchmarkIterations
        startedAt = CDbl(Timer)
        totalIssues = totalIssues + RunScalarPass(SourceSchema, InputValue, ExpectedSuccess)
        samples(iterationIndex) = ElapsedMilliseconds(startedAt)
    Next iterationIndex

    OutputJson = BuildFixtureJson(FixtureId, TargetMilliseconds, scalarValidationCount * benchmarkIterations, totalIssues, samples)
End Sub

Private Function RunObjectPass(ByVal SourceSchema As VSchema, ByVal InputObject As Object, ByVal ExpectedSuccess As Boolean) As Long
    Dim result As VValidationResult
    Set result = SourceSchema.SafeParse(InputObject)
    If result.Success <> ExpectedSuccess Then
        Err.Raise vbObjectError + 2299, "ValidationBenchmarks.RunObjectPass", "Object fixture returned an unexpected result."
    End If
    If Not result.Success Then
        RunObjectPass = result.Issues.Count
    End If
End Function

Private Function RunScalarPass(ByVal SourceSchema As VSchema, ByVal InputValue As Variant, ByVal ExpectedSuccess As Boolean) As Long
    Dim index As Long
    Dim result As VValidationResult
    For index = 1 To scalarValidationCount
        Set result = SourceSchema.SafeParse(InputValue)
        If result.Success <> ExpectedSuccess Then
            Err.Raise vbObjectError + 2299, "ValidationBenchmarks.RunScalarPass", "Scalar fixture returned an unexpected result."
        End If
        If Not result.Success Then
            RunScalarPass = RunScalarPass + result.Issues.Count
        End If
    Next index
End Function

Private Function BuildFixtureJson( _
        ByVal FixtureId As String, _
        ByVal TargetMilliseconds As Double, _
        ByVal ValidationCount As Long, _
        ByVal IssueCount As Long, _
        ByRef Samples() As Double) As String

    Dim sampleJson As String
    Dim index As Long
    For index = LBound(Samples) To UBound(Samples)
        If index > LBound(Samples) Then
            sampleJson = sampleJson & ","
        End If
        sampleJson = sampleJson & InvariantNumber(Samples(index))
    Next index

    BuildFixtureJson = "{" & _
        """fixture_id"":""" & JsonEscape(FixtureId) & """," & _
        """warmup_iterations"":" & CStr(benchmarkWarmupIterations) & "," & _
        """iterations"":" & CStr(benchmarkIterations) & "," & _
        """raw_ms"": [" & sampleJson & "]," & _
        """median_ms"":" & InvariantNumber(Median(Samples)) & "," & _
        """validation_count"":" & CStr(ValidationCount) & "," & _
        """issue_count"":" & CStr(IssueCount) & "," & _
        """regexp_creation_count"":0," & _
        """target_ms"":" & InvariantNumber(TargetMilliseconds) & _
        "}"
End Function

Private Function Median(ByRef Samples() As Double) As Double
    Dim sorted(1 To benchmarkIterations) As Double
    Dim index As Long
    Dim position As Long
    Dim currentValue As Double

    For index = 1 To benchmarkIterations
        sorted(index) = Samples(index)
    Next index

    For index = 2 To benchmarkIterations
        currentValue = sorted(index)
        position = index - 1
        Do While position >= 1
            If sorted(position) <= currentValue Then
                Exit Do
            End If
            sorted(position + 1) = sorted(position)
            position = position - 1
        Loop
        sorted(position + 1) = currentValue
    Next index

    Median = sorted((benchmarkIterations + 1) \ 2)
End Function

Private Function ElapsedMilliseconds(ByVal StartedAt As Double) As Double
    Dim finishedAt As Double
    finishedAt = CDbl(Timer)
    If finishedAt < StartedAt Then
        finishedAt = finishedAt + 86400#
    End If
    ElapsedMilliseconds = (finishedAt - StartedAt) * 1000#
End Function

Private Function InvariantNumber(ByVal Value As Double) As String
    InvariantNumber = Trim$(Str$(Value))
End Function

Private Function OfficeBitness() As String
    #If Win64 Then
    OfficeBitness = "x64"
    #Else
    OfficeBitness = "x86"
    #End If
End Function

Private Function JsonEscape(ByVal Value As String) As String
    Value = Replace$(Value, "\", "\\")
    Value = Replace$(Value, Chr$(34), Chr$(92) & Chr$(34))
    Value = Replace$(Value, vbCrLf, "\n")
    Value = Replace$(Value, vbCr, "\n")
    Value = Replace$(Value, vbLf, "\n")
    Value = Replace$(Value, vbTab, "\t")
    JsonEscape = Value
End Function
