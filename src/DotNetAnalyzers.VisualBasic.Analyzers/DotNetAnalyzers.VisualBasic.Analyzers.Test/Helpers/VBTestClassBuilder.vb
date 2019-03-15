Imports System.Text

Public Class VBTestClassBuilder

    Private _testClass As StringBuilder

    Public Sub New()
        _testClass = New StringBuilder()
    End Sub

    Public Function WithSub(subName As String, content As String) As VBTestClassBuilder
        With _testClass
			.AppendLine("	Public Sub " & subName)
			.AppendLine("		" & content)
			.AppendLine("	End Sub")
		End With

        Return Me
    End Function
	Public Function WithFunction(functionName As String, resultType As String, content As String) As VBTestClassBuilder
        With _testClass
			.AppendLine($"	Public Function {functionName} As {resultType}")
			.AppendLine("		" & content)
			.AppendLine("	End Function")
		End With

        Return Me
    End Function

	Public Function WithAutoProperty(propertyName As String, dataType As String) As VBTestClassBuilder
		_testClass.AppendLine($"	Public Property {propertyName} As {dataType}")
		_testClass.AppendLine("")
		Return Me
	End Function

	Public Function WithReadOnlyProperty(propertyName As String, dataType As String, returnResult As String) As VBTestClassBuilder
		With _testClass
			.AppendLine($"	Public ReadOnly Property {propertyName} As {dataType}")
			.AppendLine("		Get")
			.AppendLine($"			Return {returnResult}")
			.AppendLine("		End Get")
			.AppendLine("	End Property")
		End With

		Return Me
	End Function

	Public Function BuildClass(className As String) As String
        Dim classStub = New StringBuilder()

        With classStub
			.AppendLine("Public Class " & className)
			.AppendLine("")
			.AppendLine(_testClass.ToString())
			.AppendLine("End Class")
			.AppendLine("Public Module TestModule")
			.AppendLine("	Public Sub Main(args As String())")
			.AppendLine("	End Sub")
			.AppendLine("End Module")
		End With

        Return classStub.ToString()
    End Function

End Class
