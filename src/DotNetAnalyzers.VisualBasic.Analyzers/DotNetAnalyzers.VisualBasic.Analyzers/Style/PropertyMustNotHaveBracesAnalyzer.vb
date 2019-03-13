Imports System.Collections.Immutable
Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

<DiagnosticAnalyzer(LanguageNames.VisualBasic)>
Public Class PropertyMustNotHaveBracesAnalyzer
	Inherits DiagnosticAnalyzer

	Public Const DiagnosticId = "PropertyMustNotHaveBraces"
	Friend Shared ReadOnly TitlePropertyDeclaration As String = "Property declaration must not have braces"
	Public Shared ReadOnly MessageFormatPropertyDeclaration As String = "Remove the braces of property declaration '{0}'."
	Friend Shared ReadOnly TitlePropertyCall As String = "Property call must not habe braces"
	Public Shared ReadOnly MessageFormatPropertyCall As String = "Remove braces while calling property '{0}'."
	Friend Const Category = "Style"

	Friend Shared PropertyDeclarationRule As New DiagnosticDescriptor(DiagnosticId, TitlePropertyDeclaration, MessageFormatPropertyDeclaration, Category, DiagnosticSeverity.Warning, True)
	Friend Shared PropertyCallRule As New DiagnosticDescriptor(DiagnosticId, TitlePropertyCall, MessageFormatPropertyCall, Category, DiagnosticSeverity.Warning, True)

	Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
		Get
			Return ImmutableArray.Create(PropertyDeclarationRule, PropertyCallRule)
		End Get
	End Property

	Public Overrides Sub Initialize(context As AnalysisContext)
		context.RegisterSyntaxNodeAction(AddressOf PropertyCallHasNoBracesAnalyzer, SyntaxKind.InvocationExpression)
		context.RegisterSyntaxNodeAction(AddressOf PropertyDeclarationHasNoBracesAnalyzer, SyntaxKind.PropertyStatement)
	End Sub

	Private Sub PropertyCallHasNoBracesAnalyzer(context As SyntaxNodeAnalysisContext)
		Dim expression = TryCast(context.Node, InvocationExpressionSyntax)

		If (expression Is Nothing) Then Return

		Dim symbol = context.SemanticModel.GetSymbolInfo(expression)
		Dim propertyInfo = TryCast(symbol.Symbol, IPropertySymbol)
		If (propertyInfo Is Nothing) Then Return

		If (propertyInfo.Parameters.Any()) Then Return

		Dim diagnosticResult = Diagnostic.Create(PropertyCallRule, expression.GetLocation(), propertyInfo.Name)
		context.ReportDiagnostic(diagnosticResult)
	End Sub


	Private Sub PropertyDeclarationHasNoBracesAnalyzer(context As SyntaxNodeAnalysisContext)
		Dim expression = TryCast(context.Node, PropertyStatementSyntax)

		If (expression Is Nothing) Then Return

		If (expression.ParameterList IsNot Nothing) Then
			If (expression.ParameterList.Parameters.Any() = False) Then
				Dim diagnosticResult = Diagnostic.Create(PropertyDeclarationRule, expression.Identifier.GetLocation(), expression.Identifier.GetIdentifierText())
				context.ReportDiagnostic(diagnosticResult)
			End If
		End If
	End Sub


End Class
