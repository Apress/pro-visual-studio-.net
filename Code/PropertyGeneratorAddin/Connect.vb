Imports Microsoft.Office.Core
Imports Extensibility
Imports System.Runtime.InteropServices
Imports EnvDTE

#Region " Read me for Add-in installation and setup information. "
' When run, the Add-in wizard prepared the registry for the Add-in.
' At a later time, if the Add-in becomes unavailable for reasons such as:
'   1) You moved this project to a computer other than which is was originally created on.
'   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
'   3) Registry corruption.
' you will need to re-register the Add-in by building the PropertyGeneratorAddinSetup project 
' by right clicking the project in the Solution Explorer, then choosing install.
#End Region

<GuidAttribute("8F4F3750-17C6-45C0-843E-AC162AD53B8E"), ProgIdAttribute("PropertyGeneratorAddin.Connect")> _
Public Class Connect

    Implements Extensibility.IDTExtensibility2
    Implements IDTCommandTarget

    Dim applicationObject As EnvDTE.DTE
    Dim addInInstance As EnvDTE.AddIn

    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnBeginShutdown
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnAddInsUpdate
    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnStartupComplete
    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As Extensibility.ext_DisconnectMode, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnDisconnection
    End Sub

    Public Sub OnConnection(ByVal application As Object, ByVal connectMode As ext_ConnectMode, ByVal addInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection
        applicationObject = CType(application, EnvDTE.DTE)
        addInInstance = CType(addInInst, EnvDTE.AddIn)
        If connectMode = ext_ConnectMode.ext_cm_UISetup Then
            Dim objAddIn As AddIn = CType(addInInst, AddIn)
            Dim CommandObj As Command
            Dim objCommandBar As CommandBar

            'If your command no longer appears on the appropriate command bar, or if you would like to re-create the command,
            ' close all instances of Visual Studio .NET, open a command prompt (MS-DOS window), and run the command 'devenv /setup'.
            objCommandBar = CType(applicationObject.Commands.AddCommandBar("MyMacros", vsCommandBarType.vsCommandBarTypeMenu, applicationObject.CommandBars.Item("Tools")), Microsoft.Office.Core.CommandBar)

            CommandObj = applicationObject.Commands.AddNamedCommand(objAddIn, "PropertyMacro", "PropertyMacro", "TODO: Enter your command description", True, 59, Nothing, 1 + 2)  '1+2 == vsCommandStatusSupported+vsCommandStatusEnabled
            CommandObj.AddControl(objCommandBar)
        Else
            'If you are not using events, you may wish to remove some of these to increase performance.
            EnvironmentEvents.DTEEvents = CType(applicationObject.Events.DTEEvents, EnvDTE.DTEEvents)
            EnvironmentEvents.DocumentEvents = CType(applicationObject.Events.DocumentEvents(Nothing), EnvDTE.DocumentEvents)
            EnvironmentEvents.WindowEvents = CType(applicationObject.Events.WindowEvents(Nothing), EnvDTE.WindowEvents)
            EnvironmentEvents.TaskListEvents = CType(applicationObject.Events.TaskListEvents(""), EnvDTE.TaskListEvents)
            EnvironmentEvents.FindEvents = CType(applicationObject.Events.FindEvents, EnvDTE.FindEvents)
            EnvironmentEvents.OutputWindowEvents = CType(applicationObject.Events.OutputWindowEvents(""), EnvDTE.OutputWindowEvents)
            EnvironmentEvents.SelectionEvents = CType(applicationObject.Events.SelectionEvents, EnvDTE.SelectionEvents)
            EnvironmentEvents.SolutionItemsEvents = CType(applicationObject.Events.SolutionItemsEvents, EnvDTE.ProjectItemsEvents)
            EnvironmentEvents.MiscFilesEvents = CType(applicationObject.Events.MiscFilesEvents, EnvDTE.ProjectItemsEvents)
            EnvironmentEvents.DebuggerEvents = CType(applicationObject.Events.DebuggerEvents, EnvDTE.DebuggerEvents)
        End If
    End Sub

    Public Sub Exec(ByVal cmdName As String, ByVal executeOption As vsCommandExecOption, ByRef varIn As Object, ByRef varOut As Object, ByRef handled As Boolean) Implements IDTCommandTarget.Exec
        handled = False
        If (executeOption = vsCommandExecOption.vsCommandExecOptionDoDefault) Then
            If cmdName = "PropertyGeneratorAddin.Connect.PropertyMacro" Then
                handled = True
                Module1.DTE = applicationObject
                Module1.PropertyMacro()
                Exit Sub
            End If
        End If
    End Sub


    Public Sub QueryStatus(ByVal cmdName As String, ByVal neededText As vsCommandStatusTextWanted, ByRef statusOption As vsCommandStatus, ByRef commandText As Object) Implements IDTCommandTarget.QueryStatus
        statusOption = vsCommandStatus.vsCommandStatusUnsupported
        If neededText = EnvDTE.vsCommandStatusTextWanted.vsCommandStatusTextWantedNone Then
            If cmdName = "PropertyGeneratorAddin.Connect.PropertyMacro" Then
                statusOption = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
            End If

        End If
    End Sub
End Class
