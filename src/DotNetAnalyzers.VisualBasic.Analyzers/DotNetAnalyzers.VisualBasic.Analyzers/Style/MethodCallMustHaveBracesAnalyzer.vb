Imports System.Collections.Immutable
Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

<DiagnosticAnalyzer(LanguageNames.VisualBasic)>
Public Class MethodCallMustHaveBracesAnalyzer
	Inherits DiagnosticAnalyzer

	Public Const DiagnosticId = "MethodCallMustHaveBraces"
	Friend Shared ReadOnly MethodCallTitle As String = "Method call must have braces"
	Public Shared ReadOnly MethodCallMessage As String = "Add braces to method call '{0}'"
	Friend Shared ReadOnly ConstructorCallTitle As String = "Constructor must habe braces"
	Public Shared ReadOnly ConstructorCallMessage As String = "Add braces to constructor"
	Friend Shared ReadOnly MethodDeclarationTitle As String = "Method declarations must have braces"
	Public Shared ReadOnly MethodDeclarationMessage As String = "Add braces to constructor"

	Friend Const Category = "Style"

	Friend Shared MethodCallRule As New DiagnosticDescriptor(DiagnosticId, MethodCallTitle, MethodCallMessage, Category, DiagnosticSeverity.Warning, True)
	Friend Shared ConstructorCallRule As New DiagnosticDescriptor(DiagnosticId, ConstructorCallTitle, ConstructorCallMessage, Category, DiagnosticSeverity.Warning, True)
	Friend Shared MethodDeclarationRule As New DiagnosticDescriptor(DiagnosticId, MethodDeclarationTitle, MethodDeclarationMessage, Category, DiagnosticSeverity.Warning, True)

	Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
		Get
			Return ImmutableArray.Create(MethodCallRule, ConstructorCallRule, MethodDeclarationRule)
		End Get
	End Property

	Public Overrides Sub Initialize(context As AnalysisContext)
		context.RegisterSyntaxNodeAction(AddressOf ObjectCreationAnalyzer, SyntaxKind.ObjectCreationExpression)
		context.RegisterSyntaxNodeAction(AddressOf MethodCallBraceAnalyzer, SyntaxKind.InvocationExpression)
        context.RegisterSyntaxNodeAction(AddressOf MethodDeclaration, SyntaxKind.SubStatement, SyntaxKind.FunctionStatement)
    End Sub

    Private Sub MethodDeclaration(context As SyntaxNodeAnalysisContext)
        Dim methodStatement = TryCast(context.Node, MethodStatementSyntax)

		If(methodStatement Is Nothing) Then Return

		If(methodStatement.ParameterList Is Nothing) Then
			Dim result = Diagnostic.Create(MethodDeclarationRule, methodStatement.GetLocation())
			context.ReportDiagnostic(result)
		End If
    End Sub

    Private Sub MethodCallBraceAnalyzer(context As SyntaxNodeAnalysisContext)
		Dim methodCall = TryCast(context.Node, InvocationExpressionSyntax)

		If (methodCall Is Nothing) Then Return

		Dim model = context.SemanticModel
		Dim symbol = model.GetSymbolInfo(methodCall)

		Dim methodSymbol = TryCast(symbol.Symbol, IMethodSymbol)

		If (methodSymbol Is Nothing) Then Return

		If (methodCall.ArgumentList IsNot Nothing) Then Return
		
		Dim methodName = GetMethodName(methodCall, context.SemanticModel)
		If (methodName?.Equals("InitializeComponent", StringComparison.OrdinalIgnoreCase)) Then Return
		
		Dim result = Diagnostic.Create(MethodCallRule, methodCall.GetLocation(), methodSymbol.Name)
		context.ReportDiagnostic(result)
	End Sub

	Private Sub ObjectCreationAnalyzer(context As SyntaxNodeAnalysisContext)
		Dim createdObject = TryCast(context.Node, ObjectCreationExpressionSyntax)

		If (createdObject Is Nothing) Then Return

		If (createdObject.ArgumentList Is Nothing) Then
			Dim result = Diagnostic.Create(ConstructorCallRule, createdObject.GetLocation())
			context.ReportDiagnostic(result)
		End If
	End Sub

	Private Function GetMethodName(expression As SyntaxNode, model As SemanticModel) As String
		If (expression Is Nothing) Then Return Nothing
		
		Dim methodExpression = TryCast(expression, MethodBlockSyntax)

		If (methodExpression Is Nothing) Then
			Return GetMethodName(expression.Parent, model)
		End If

		Return methodExpression.SubOrFunctionStatement.Identifier.ValueText
	End Function

End Class
