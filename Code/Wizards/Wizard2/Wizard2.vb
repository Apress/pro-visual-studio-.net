Imports EnvDTE
Imports System.CodeDom
Imports System.CodeDom.Compiler
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Namespace Com.Apress.VSNET.Wizard2

    ''' 
    ''' Main implementing class. Implements Execute method where it gathers input from the user interface
    ''' and finally builds the project
    '''
    <Guid("6B967689-DE9C-4fbc-9410-FC6A7C7C6F06")> _
    Public Class Wizard2
        Implements IDTWizard

        ''' IDTWizard.Execute implementation
        Public Sub Execute(ByVal Application As Object, ByVal hwndOwner As Integer, ByRef ContextParams() As Object, _
            ByRef CustomParams() As Object, ByRef retval As EnvDTE.wizardResult) Implements EnvDTE.IDTWizard.Execute

            ''' Wizard2 main code.
            Dim generator As CodeSupport.IWizardCodeGenerator = New WizardGeneratorForm

            ' check if we should generate the code
            If (generator.ShowUI()) Then
                ' call generator's generate method to do the actual processing.
                retval = generator.Generate(CType(Application, _DTE), ContextParams, CustomParams)
            End If

        End Sub
    End Class

End Namespace
