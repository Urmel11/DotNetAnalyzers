Imports DotNetAnalyzers.VisualBasic.Analyzers.Test.TestHelper
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports NUnit.Framework

Public Class MethodCallMustHaveBracesAnalyzerTests
	Inherits DiagnosticVerifier

	Protected Overrides Function GetBasicDiagnosticAnalyzer() As DiagnosticAnalyzer
		Return New MethodCallMustHaveBracesAnalyzer()
	End Function

	Public Class AnalyzeMethod
		Inherits MethodCallMustHaveBracesAnalyzerTests

		Private _testClassBuilder As VBTestClassBuilder

		<SetUp>
		Public Sub SetUp()
			_testClassBuilder = New VBTestClassBuilder()
		End Sub

		<Test>
		Public Sub Is_Valid_If_Object_Creation_Is_Done_With_Braces()
			Dim testClass = _testClassBuilder.
							WithSub("TestSub", "Dim test = New System.Text.StringBuilder()").
							BuildClass("TestClass")

			VerifyBasicDiagnostic(testClass)
		End Sub

		<Test>
		Public Sub Is_Valid_If_Method_Call_Is_Done_With_Braces()
			Dim testClass = _testClassBuilder.
							WithSub("TestSub", "AnotherSub()").
							WithSub("AnotherSub", "").
							BuildClass("TestClass")

			VerifyBasicDiagnostic(testClass)
		End Sub

		<Test>
		Public Sub Throws_Error_If_Method_Call_Has_No_Braces()
			Dim testClass = _testClassBuilder.
							WithSub("TestSub", "AnotherSub").
							WithSub("AnotherSub", "").
							BuildClass("TestClass")

			Dim errorMessage = String.Format(MethodCallMustHaveBracesAnalyzer.MethodCallMessage, "AnotherSub")
			Dim expected = CreateExpectedDiagnosticResult(4, 3, errorMessage)

			VerifyBasicDiagnostic(testClass, expected)
		End Sub

		<Test>
		Public Sub Thorws_Error_If_Object_Creation_Is_Done_Without_Braces()
			Dim testClass = _testClassBuilder.
										WithSub("TestSub", "Dim test = New System.Text.StringBuilder").
										BuildClass("TestClass")

			Dim expected = CreateExpectedDiagnosticResult(4, 14, MethodCallMustHaveBracesAnalyzer.ConstructorCallMessage)

			VerifyBasicDiagnostic(testClass, expected)
		End Sub

		Private Function CreateExpectedDiagnosticResult(line As Integer, column As Integer, message As String) As DiagnosticResult
			Return New DiagnosticResult With {
												.Id = MethodCallMustHaveBracesAnalyzer.DiagnosticId,
												.Message = message,
												.Severity = DiagnosticSeverity.Warning,
												.Locations = New DiagnosticResultLocation() {New DiagnosticResultLocation("Test0.vb", line, column)}
											}
		End Function

	End Class

End Class
