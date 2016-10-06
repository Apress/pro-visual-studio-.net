Imports EnvDTE
Imports VSLangProj
Imports System.CodeDom
Imports System.CodeDom.Compiler
Imports System.Text
Imports System.IO
Imports Wizard2.Com.Apress.VSNET.Wizard2.CodeSupport

Imports Microsoft.CSharp
Imports Microsoft.VisualBasic
Imports Microsoft.MCpp

Namespace Com.Apress.VSNET.Wizard2

    Public Class WizardGeneratorForm
        Inherits System.Windows.Forms.Form
        Implements IWizardCodeGenerator


#Region " Windows Form Designer generated code "

        Public Sub New()
            MyBase.New()

            'This call is required by the Windows Form Designer.
            InitializeComponent()

            'Add any initialization after the InitializeComponent() call

        End Sub

        'Form overrides dispose to clean up the component list.
        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        Friend WithEvents Label1 As System.Windows.Forms.Label
        Friend WithEvents Label2 As System.Windows.Forms.Label
        Friend WithEvents okButton As System.Windows.Forms.Button
        Friend WithEvents wizardName As System.Windows.Forms.TextBox
        Friend WithEvents wizardDescription As System.Windows.Forms.TextBox
        Friend WithEvents abortButton As System.Windows.Forms.Button
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            Me.Label1 = New System.Windows.Forms.Label
            Me.wizardName = New System.Windows.Forms.TextBox
            Me.Label2 = New System.Windows.Forms.Label
            Me.wizardDescription = New System.Windows.Forms.TextBox
            Me.okButton = New System.Windows.Forms.Button
            Me.abortButton = New System.Windows.Forms.Button
            Me.SuspendLayout()
            '
            'Label1
            '
            Me.Label1.Location = New System.Drawing.Point(14, 18)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(100, 16)
            Me.Label1.TabIndex = 0
            Me.Label1.Text = "Wizard name"
            '
            'wizardName
            '
            Me.wizardName.Location = New System.Drawing.Point(118, 16)
            Me.wizardName.Name = "wizardName"
            Me.wizardName.Size = New System.Drawing.Size(376, 20)
            Me.wizardName.TabIndex = 1
            Me.wizardName.Text = ""
            '
            'Label2
            '
            Me.Label2.Location = New System.Drawing.Point(12, 48)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New System.Drawing.Size(100, 16)
            Me.Label2.TabIndex = 0
            Me.Label2.Text = "Wizard description"
            '
            'wizardDescription
            '
            Me.wizardDescription.Location = New System.Drawing.Point(118, 44)
            Me.wizardDescription.Name = "wizardDescription"
            Me.wizardDescription.Size = New System.Drawing.Size(376, 20)
            Me.wizardDescription.TabIndex = 1
            Me.wizardDescription.Text = ""
            '
            'okButton
            '
            Me.okButton.Location = New System.Drawing.Point(342, 74)
            Me.okButton.Name = "okButton"
            Me.okButton.TabIndex = 3
            Me.okButton.Text = "&OK"
            '
            'abortButton
            '
            Me.abortButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.abortButton.Location = New System.Drawing.Point(422, 74)
            Me.abortButton.Name = "abortButton"
            Me.abortButton.TabIndex = 4
            Me.abortButton.Text = "&Cancel"
            '
            'WizardGeneratorForm
            '
            Me.AcceptButton = Me.okButton
            Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
            Me.ClientSize = New System.Drawing.Size(506, 115)
            Me.Controls.Add(Me.abortButton)
            Me.Controls.Add(Me.okButton)
            Me.Controls.Add(Me.wizardName)
            Me.Controls.Add(Me.Label1)
            Me.Controls.Add(Me.Label2)
            Me.Controls.Add(Me.wizardDescription)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "WizardGeneratorForm"
            Me.ShowInTaskbar = False
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            Me.Text = "Wizard^2"
            Me.ResumeLayout(False)

        End Sub

#End Region

        '''
        ''' Generate the code
        '''
        Public Function Generate(ByVal dte As EnvDTE._DTE, ByRef ContextParams() As Object, _
            ByRef CustomParams() As Object) As EnvDTE.wizardResult _
            Implements IWizardCodeGenerator.Generate

            Dim language As String = "VB"

            For Each o As Object In CustomParams
                Dim params() As String = CType(o, String).Split("=")
                Dim param As String = params(0)
                Dim value As String = params(1)

                If (param = "language") Then
                    language = value
                End If
            Next

            ' assign code names
            Dim className As String = wizardName.Text

            ' get settings from ContextParams
            Dim projectName As String = ContextParams(1)
            Dim projectDirectory As String = ContextParams(2)

            ' get the solution object
            Dim sol As Solution = SolutionGeneratorFactory.GetSolutionGenerator(language).Create(dte, projectName, projectDirectory)

            ' create project from existing
            Dim project As Project = ProjectGeneratorFactory.GetProjectGenerator(language).Create(dte, sol, projectName, projectDirectory)

            ' create wizard namespace, class and main method.
            Dim mainClass As CodeTypeDeclaration = New CodeTypeDeclaration("Demo")

            Dim mainMethod As CodeEntryPointMethod = New CodeEntryPointMethod
            mainMethod.Name = "Main"
            mainMethod.ReturnType = Nothing
            mainMethod.Attributes = MemberAttributes.Static Or MemberAttributes.Public

            mainClass.Members.Add(mainMethod)

            Dim rootNS As System.CodeDom.CodeNamespace = New System.CodeDom.CodeNamespace("com.apress.vsnet.wizard")
            rootNS.Types.Add(mainClass)

            ' generate the code and add it to the file
            Dim tempClassFile As String = Path.GetTempFileName()
            Dim stream As FileStream = New FileStream(tempClassFile, FileMode.Create)
            Dim writer As StreamWriter = New StreamWriter(stream)
            CodeProviderFactory.GetCodeProvider(language).CreateGenerator().GenerateCodeFromNamespace(rootNS, writer, Nothing)
            writer.Close()
            stream.Close()

            ' create a simple new class
            Dim projectItem As ProjectItem
            projectItem = project.ProjectItems.AddFromFileCopy(tempClassFile)
            projectItem.Name = String.Format("Main{0}", ProjectGeneratorFactory.GetProjectGenerator(language).SourceItemExtension)

            File.Delete(tempClassFile)
        End Function

        '''
        ''' Return Me as this form is the wizard form
        '''
        Public Function ShowUI() As Boolean Implements IWizardCodeGenerator.ShowUI
            Return Me.ShowDialog() = DialogResult.OK
        End Function

        Private Sub okButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles okButton.Click
            If (wizardName.Text = String.Empty) Then
                MsgBox("You must specify wizard name", MsgBoxStyle.OKOnly)
                Return
            End If
            If (wizardDescription.Text = String.Empty) Then
                MsgBox("You must specify wizard description", MsgBoxStyle.OKOnly)
                Return
            End If

            ' all OK, set DialogResult to DialogResult.OK
            DialogResult = DialogResult.OK
        End Sub
    End Class

End Namespace
