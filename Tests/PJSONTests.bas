Attribute VB_Name = "PJSONTests"
'@IgnoreModule LineLabelNotUsed, AssignmentNotUsed, VariableNotUsed
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

#If DEV Then
  Private Assert As Rubberduck.AssertClass
  Private Fakes As Rubberduck.FakesProvider
#Else
  Private Assert As Object
  Private Fakes As Object
#End If

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
  #If DEV Then
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
  #Else
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
  #End If
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Uncategorized")
Private Sub EmptyJSON_Test()
    On Error GoTo TestFail
    
    'Arrange:
    Dim JSONString As String: JSONString = "{}"
  #If DEV Then
    Dim Result As Dictionary: Set Result = PJSON.Parse(JSONString)
  #Else
    Dim Result As Object: Set Result = PJSON.Parse(JSONString)
  #End If
    
    'Act:
    Assert.AreEqual Conversion.CLng(0), Result.Count, "Check keys count."
    
    'Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub SimpleObject_Test()
    On Error GoTo TestFail
    
    'Arrange:
    Dim JSONString As String: JSONString = "{ ""key"": ""value"" }"
  #If DEV Then
    Dim Result As Dictionary: Set Result = PJSON.Parse(JSONString)
  #Else
    Dim Result As Object: Set Result = PJSON.Parse(JSONString)
  #End If
    
    'Act:
    Assert.AreEqual Conversion.CLng(1), Result.Count, "Check keys count."
    Assert.AreEqual "key", Result.Keys()(0), "Check key."
    Assert.AreEqual "value", Result("key"), "Check value."
    
    'Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub MultilevelObject_Test()
    On Error GoTo TestFail
    
    'Arrange:
    Dim JSONString As String: JSONString = "{ ""key"": { ""key2"": ""value"" } }"
  #If DEV Then
    Dim Result As Dictionary: Set Result = PJSON.Parse(JSONString)
  #Else
    Dim Result As Object: Set Result = PJSON.Parse(JSONString)
  #End If
  
    'Act:
    Assert.AreEqual Conversion.CLng(1), Result.Count, "Check keys count."
    Assert.AreEqual "key", Result.Keys()(0), "Check key."
    Assert.IsTrue Information.IsObject(Result("key")), "Check nested object."

    Assert.AreEqual Conversion.CLng(1), Result("key").Count, "Check nested object keys count."
    Assert.AreEqual "key2", Result("key").Keys()(0), "Check nested key."
    Assert.AreEqual "value", Result("key")("key2"), "Check nested value."
    
    'Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Array_Test()
    On Error GoTo TestFail
    
    'Arrange:
    Dim JSONString As String: JSONString = "{ ""array"": [1, 2, 3] }"
  #If DEV Then
    Dim Result As Dictionary: Set Result = PJSON.Parse(JSONString)
  #Else
    Dim Result As Object: Set Result = PJSON.Parse(JSONString)
  #End If
    
    'Act:
    Assert.IsTrue Information.IsObject(Result("array")), "Check is array."
    Assert.AreEqual Conversion.CLng(3), Result("array").Count, "Check array size."
    Assert.AreEqual Conversion.CDbl(1), Result("array").Item(1), "Check array element 1."
    Assert.AreEqual Conversion.CDbl(2), Result("array").Item(2), "Check array element 2."
    Assert.AreEqual Conversion.CDbl(3), Result("array").Item(3), "Check array element 3."

    'Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub DifferentTypes_Test()
    On Error GoTo TestFail
    
    'Arrange:

    Dim JSONString As String
    JSONString = "{""string"": ""value"", ""number"": 123, ""boolean"": true, ""null"": null}"
  #If DEV Then
    Dim Result As Dictionary: Set Result = PJSON.Parse(JSONString)
  #Else
    Dim Result As Object: Set Result = PJSON.Parse(JSONString)
  #End If
    
    'Act:
    Assert.AreEqual "value", Result("string"), "Check string value."
    Assert.IsTrue Information.VarType(Result("string")) = VbVarType.vbString, "Check string type."

    Assert.AreEqual Conversion.CDbl(123), Result("number"), "Check number value."
    Assert.IsTrue Information.IsNumeric(Result("number")), "Check is number."
    Assert.IsTrue Information.VarType(Result("number")) = VbVarType.vbDouble, "Check number type."

    Assert.AreEqual True, Result("boolean"), "Check boolean value."
    Assert.IsTrue Information.VarType(Result("boolean")) = VbVarType.vbBoolean, "Check boolean type."

    Assert.AreEqual Null, Result("null"), "Check null value."
    Assert.IsTrue Information.VarType(Result("null")) = VbVarType.vbNull, "Check null type."

    'Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub InvalidValue_Test()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim JSONString As String: JSONString = "{ ""key"": value }"
  #If DEV Then
    Dim Result As Dictionary: Set Result = PJSON.Parse(JSONString)
  #Else
    Dim Result As Object: Set Result = PJSON.Parse(JSONString)
  #End If
    
    'Act:
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Uncategorized")
Private Sub MultiNested_Test()
    On Error GoTo TestFail
    
    'Arrange:
    Dim JSONString As String
    JSONString = "{""object"": {""array"": [1, 2, {""nested"": ""object""}]}, ""bool"": false}"
  #If DEV Then
    Dim Result As Dictionary: Set Result = PJSON.Parse(JSONString)
  #Else
    Dim Result As Object: Set Result = PJSON.Parse(JSONString)
  #End If
    
    'Act:
    Assert.IsTrue Utils.IsDictionary(Result("object")), "Check if is object."
    Assert.IsTrue Utils.IsCollection(Result("object")("array")), "Check if is array."
    Assert.IsTrue Utils.IsDictionary(Result("object")("array")(3)), "Check if is object."
    Assert.IsTrue Information.VarType(Result("bool")) = VbVarType.vbBoolean, "Check if is boolean."
    
Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Float_Test()
    On Error GoTo TestFail
    
    'Arrange:
    Dim JSONString As String
    JSONString = "{""float"": 123.456}"
  #If DEV Then
    Dim Result As Dictionary: Set Result = PJSON.Parse(JSONString)
  #Else
    Dim Result As Object: Set Result = PJSON.Parse(JSONString)
  #End If
    
    'Act:
    Assert.IsTrue Information.VarType(Result("float")) = VbVarType.vbDouble, "Check if is double."
    Assert.AreEqual 123.456, Result("float"), "Check value."
    
Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub ArrayOfObjects_Test()
    On Error GoTo TestFail
    
    'Arrange:
    Dim JSONString As String
    JSONString = "[{""obj1"": ""value1""}, {""obj2"": ""value2""}]"
  #If DEV Then
    Dim Result As Dictionary: Set Result = PJSON.Parse(JSONString)
  #Else
    Dim Result As Object: Set Result = PJSON.Parse(JSONString)
  #End If
    
    'Act:
    Assert.IsTrue Utils.IsCollection(Result), "Check if is array."
    Assert.IsTrue Utils.IsDictionary(Result(1)), "Check if is object 1."
    Assert.IsTrue Utils.IsDictionary(Result(2)), "Check if is object 2."
    Assert.AreEqual "value1", Result(1)("obj1"), "Check value of object 1."
    Assert.AreEqual "value2", Result(2)("obj2"), "Check value of object 2."
    
Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Stringify_Test()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Expected As String
    Expected = "{""object"": {""array"": [1, 2, {""nested"": ""object""}]}, ""bool"": false}"
  #If DEV Then
    Dim JSONRoot As Dictionary: Set JSONRoot = NewDictionary()
    Dim JSONObject As Dictionary: Set JSONObject = NewDictionary()
    Dim JSONArray As Collection: Set JSONArray = New Collection
    Dim JSONNested As Dictionary: Set JSONNested = NewDictionary()
  #Else
    Dim JSONRoot As Object: Set JSONRoot = NewDictionary()
    Dim JSONObject As Object: Set JSONObject = NewDictionary()
    Dim JSONArray As Collection: Set JSONArray = New Collection
    Dim JSONNested As Object: Set JSONNested = NewDictionary()
  #End If

    JSONNested("nested") = "object"
    JSONArray.Add 1
    JSONArray.Add 2
    JSONArray.Add JSONNested
    JSONObject.Add "array", JSONArray
    JSONRoot.Add "object", JSONObject
    JSONRoot.Add "bool", False
    Dim Result As String: Result = Strings.Join(Strings.Split(PJSON.Stringify(JSONRoot), vbNewLine), "")
    
    'Act:
    Assert.AreEqual Expected, Result, "Check stringify."
    
Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
