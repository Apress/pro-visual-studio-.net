Imports EnvDTE
Imports System.Windows.Forms
Imports System.CodeDom
Imports System.CodeDom.Compiler
Imports System.Text
Imports System.IO

Imports Microsoft.CSharp
Imports Microsoft.VisualBasic
Imports Microsoft.MCpp

Namespace Com.Apress.VSNET.Wizard2.CodeSupport

    Public Interface IWizardCodeGenerator

        ''' 
        ''' Returs Form the wizard needs to generate the code
        '''
        Function ShowUI() As Boolean

        '''
        ''' Generates the code if GetForm() returns Nothing or GetForm.ShowDialog() = DialogResult.OK
        '''
        Function Generate(ByVal dte As _DTE, ByRef ContextParams() As Object, ByRef CustomParams() As Object) As EnvDTE.wizardResult

    End Interface

    ''' Interface for generating new Solution
    Public Interface ISolutionGenerator
        Function Create(ByVal dte As _DTE, ByVal name As String, ByVal directory As String) As Solution
    End Interface

    ''' Interface for generating new Project
    Public Interface IProjectGenerator
        Function Create(ByVal dte As _DTE, ByVal solution As Solution, ByVal name As String, ByVal directory As String) As Project
        ReadOnly Property SourceItemExtension() As String
    End Interface

    Public Class VBSolutionGenerator
        Implements ISolutionGenerator

        Function Create(ByVal dte As _DTE, ByVal name As String, ByVal directory As String) As Solution _
            Implements ISolutionGenerator.Create
            Dim sol As Solution = dte.Solution

            ' create new solution
            sol.Create(directory, name)

            Return sol
        End Function
    End Class

    Public Class VBProjectGenerator
        Implements IProjectGenerator

        Public Function Create(ByVal dte As _DTE, ByVal solution As EnvDTE.Solution, ByVal name As String, ByVal directory As String) As EnvDTE.Project Implements IProjectGenerator.Create
            Dim templatesRoot As String = dte.Solution.ProjectItemsTemplatePath(VSLangProj.PrjKind.prjKindVBProject) & "\..\VBWizards"
            Dim projectTemplate = templatesRoot & "\ConsoleApplication\Templates\1033\ConsoleApplication.vbproj"

            ' create project from existing
            Dim project As Project = solution.AddFromTemplate(projectTemplate, directory, name, True)

            Return project
        End Function

        Public ReadOnly Property SourceItemExtension() As String Implements IProjectGenerator.SourceItemExtension
            Get
                Return ".vb"
            End Get
        End Property
    End Class


    ''' 
    ''' Factory class that returns ICodeGenerator for the specific language
    '''
    Public Class CodeProviderFactory

        '''
        ''' Returns ICodeGenerator for specified language
        '''
        Public Shared Function GetCodeProvider(ByVal language As String) As CodeDomProvider
            Dim provider As CodeDomProvider = Nothing
            If (language = "VB") Then
                provider = New VBCodeProvider
            ElseIf (language = "C#") Then
                provider = New CSharpCodeProvider
            ElseIf (language = "CPP") Then
                provider = New MCppCodeProvider
            Else
                Throw New ArgumentException(String.Format("Language {0} is not supported", language))
            End If

            Return provider

        End Function

    End Class

    ''' Factory class for generating ISolutionGenerator 
    Public Class SolutionGeneratorFactory
        Public Shared Function GetSolutionGenerator(ByVal language As String) As ISolutionGenerator
            If (language = "VB") Then
                Return New VBSolutionGenerator
            End If

            Throw New ArgumentException(String.Format("language {0} is not supported", language))
        End Function
    End Class

    ''' Factory class for generating IProjectGenerator
    Public Class ProjectGeneratorFactory
        Public Shared Function GetProjectGenerator(ByVal language As String) As IProjectGenerator
            If (language = "VB") Then
                Return New VBProjectGenerator
            End If

            Throw New ArgumentException(String.Format("language {0} is not supported", language))
        End Function
    End Class

End Namespace
