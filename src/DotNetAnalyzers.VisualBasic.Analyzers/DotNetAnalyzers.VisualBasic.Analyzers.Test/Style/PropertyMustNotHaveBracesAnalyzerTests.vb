Imports DotNetAnalyzers.VisualBasic.Analyzers.Test.TestHelper
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports NUnit.Framework

Public Class PropertyMustNotHaveBracesAnalyzerTests
	Inherits TestHelper.DiagnosticVerifier

	Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
		Return New PropertyMustNotHaveBracesAnalyzer()
	End Function

	Public Class AnalyzeMethod
		Inherits PropertyMustNotHaveBracesAnalyzerTests

		Private _testClassBuilder As VBTestClassBuilder

		<SetUp>
		Public Sub SetUp()
			_testClassBuilder = New VBTestClassBuilder()
		End Sub

		<Test>
		Public Sub Throws_Error_If_Property_Assignment_Is_Done_With_Braces()
			Dim testClass = _testClassBuilder.
							WithAutoProperty("TestProperty", "Integer").
							WithSub("TestSub", "TestProperty() = 1").
							BuildClass("TestClass")

			Dim expected = CreateExpectedDiagnosticResult(6, 3, PropertyMustNotHaveBracesAnalyzer.MessageFormatPropertyCall)

			VerifyBasicDiagnostic(testClass, expected)
		End Sub

		<Test>
		Public Sub Throws_Error_If_Property_Declaration_Is_Done_With_Braces()
			Dim testClass = _testClassBuilder.
							WithAutoProperty("TestProperty()", "Integer").
							BuildClass("Testclass")

			Dim expected = CreateExpectedDiagnosticResult(3, 18, PropertyMustNotHaveBracesAnalyzer.MessageFormatPropertyDeclaration)

			VerifyBasicDiagnostic(testClass, expected)
		End Sub

		<Test>
		Public Sub Throws_Error_If_Property_Is_Accessed_With_Braces()
			Dim testClass = _testClassBuilder.
							WithAutoProperty("TestProperty", "Integer").
							WithSub("TestSub", "TestProperty().Trim()").
							BuildClass("TestClass")

			Dim expected = CreateExpectedDiagnosticResult(6, 3, PropertyMustNotHaveBracesAnalyzer.MessageFormatPropertyCall)

			VerifyBasicDiagnostic(testClass, expected)
		End Sub

		<Test>
		Public Sub Is_Valid_For_Default_Indexers_Declaration()
			Dim testClass = _testClassBuilder.
							WithReadOnlyProperty("TestProperty (index As Integer)", "String", "ABC").
							BuildClass("TestClass")

			VerifyBasicDiagnostic(testClass)
		End Sub

		<Test>
		Public Sub Is_valid_If_Property_Assignment_Is_Done_Without_Braces()
			Dim testClass = _testClassBuilder.
							WithAutoProperty("TestProperty", "Integer").
							WithSub("TestSub", "TestProperty = 1").
							BuildClass("TestClass")

			VerifyBasicDiagnostic(testClass)
		End Sub

		<Test>
		Public Sub Is_Valid_If_Property_Declaration_Is_Done_Without_Braces()
			Dim testClass = _testClassBuilder.
							WithAutoProperty("TestProperty", "Integer").
							BuildClass("Testclass")

			VerifyBasicDiagnostic(testClass)
		End Sub

		<Test>
		Public Sub Is_Valid_If_Property_Is_Accessed_By_Indexer()
			Dim listDeclaration = "Dim list = New List(Of Integer)()" & vbCrLf &
									"list(0) = 0"

			Dim testClass = _testClassBuilder.
							WithSub("TestSub", listDeclaration).
							BuildClass("TestClass")

			VerifyBasicDiagnostic(testClass)
		End Sub

		<Test>
		Public Sub Is_Valid_If_System_Property_Is_Accessed_Without_Braces()
			Dim subContent = "Dim str As String = """"" & vbCrLf &
									"Console.Write(str.Length)"

			Dim testClass = _testClassBuilder.
							WithSub("TestSub", subContent).
							BuildClass("TestClass")

			VerifyBasicDiagnostic(testClass)

		End Sub

		<Test>
		Public Sub Is_Valid_If_Property_Is_Accessed_Without_Braces()
			Dim testClass = _testClassBuilder.
							WithAutoProperty("TestProperty", "Integer").
							WithSub("TestSub", "TestProperty.Trim()").
							BuildClass("TestClass")

			VerifyBasicDiagnostic(testClass)
		End Sub

		Private Function CreateExpectedDiagnosticResult(line As Integer, column As Integer, baseMessage As String) As DiagnosticResult
			Return New DiagnosticResult With {
												.Id = PropertyMustNotHaveBracesAnalyzer.DiagnosticId,
												.Message = String.Format(baseMessage, "TestProperty"),
												.Severity = DiagnosticSeverity.Warning,
												.Locations = New DiagnosticResultLocation() {New DiagnosticResultLocation("Test0.vb", line, column)}
											}
		End Function

	End Class

End Class

