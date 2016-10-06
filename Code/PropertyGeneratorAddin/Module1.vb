Option Strict Off
Imports EnvDTE
Imports System.Diagnostics
Imports System.Text

Public Module Module1
	public Dim DTE as EnvDTE.DTE

    '' Returns true if the member identified by the member argument is 
    '' a valid member for creating a property
    Function IsValidMemberString(ByVal member As String, ByRef name As String, ByRef type As String) As Boolean
        Debug.WriteLine(member)
        Try
            Dim line() As String = member.Split()
            If (line.Length = 4) Then
                'members must be in form of <Access Modifier> <Name> As <Type>
                Dim accessModifier = line(0)
                name = line(1)
                Dim asKeyword = line(2)
                type = line(3)

                Debug.WriteLine(accessModifier & " " & name & " As " & type)

                'member is valid iif 
                ' * accessModifier = [Private, Protected] And
                ' * asKeyword = As

                Return _
                    (accessModifier = "Private" Or accessModifier = "Protected") And _
                    (asKeyword = "As")
            End If
        Catch e As System.Exception
            Debug.Write(e.Message)
        End Try

        Return False
    End Function

    ''' Returns True if the current member under cursor
    ''' is a Private or Protected member
    Function IsValidMember(ByRef name As String, ByRef type As String) As Boolean
        Dim sel As TextSelection = DTE.ActiveWindow().Selection

        Try
            Dim pt As EditPoint = sel.ActivePoint.CreateEditPoint

            '' find starting point of the statement
            pt.MoveToPoint(pt.CodeElement(vsCMElement.vsCMElementDeclareDecl).GetStartPoint(vsCMPart.vsCMPartWhole))
            '' move the selection to the start of the statement
            sel.MoveToPoint(pt, False)
            '' find ending point of the statement
            pt.MoveToPoint(pt.CodeElement(vsCMElement.vsCMElementDeclareDecl).GetEndPoint(vsCMPart.vsCMPartWhole))

            pt.EndOfLine()

            sel.MoveToPoint(pt, True)

            Dim line As String = sel.Text.Trim()
            While (line.EndsWith("_"))
                line = line.Replace("_", "")
                sel.LineDown()
                sel.SelectLine()
                line &= sel.Text.Trim()
            End While

            Return IsValidMemberString(line, name, type)
        Catch e As System.Exception
            Debug.WriteLine(e.Message)
            Return False
        End Try
    End Function

    ''' This method generates the actual code that will be used
    ''' to process the property
    Function GenerateProperty(ByVal memberName As String, ByVal propertyName As String, ByVal type As String) As String
        Dim builder As StringBuilder = New StringBuilder()

        ' generate property declaration
        builder.AppendFormat(vbTab & "Public Property {0}() As {1}" & vbNewLine, propertyName, type)
        ' generate getter code
        builder.AppendFormat(vbTab & vbTab & "Get" & vbNewLine)
        builder.AppendFormat(vbTab & vbTab & vbTab & "Return Me.{0}" & vbNewLine, memberName)
        builder.AppendFormat(vbTab & vbTab & "End Get" & vbNewLine)
        ' generate setter code
        builder.AppendFormat(vbTab & vbTab & "Set(ByVal Value As {0})" & vbNewLine, type)
        builder.AppendFormat(vbTab & vbTab & vbTab & "Me.{0} = Value" & vbNewLine, memberName)
        builder.AppendFormat(vbTab & vbTab & "End Set" & vbNewLine)
        ' generate end property declaration
        builder.AppendFormat(vbTab & "End Property" & vbNewLine)

        ' return generated code
        Return builder.ToString()
    End Function

    ''' This method generates the actual code that will be used
    ''' to process the property
    Function GeneratePreProperty(ByVal memberName As String, ByVal propertyName As String, ByVal type As String) As String
        Dim comments As StringBuilder = New StringBuilder()

        comments.AppendFormat(vbTab & "''' Gets or sets {0}", memberName)
        Return comments.ToString()
    End Function

    ''' inserts the property into the active document
    Function InsertProperty(ByVal memberName As String, ByVal propertyName As String, ByVal type As String)
        Dim sel As TextSelection = DTE.ActiveWindow().Selection
        Dim pt As EditPoint = sel.ActivePoint.CreateEditPoint

        sel.SelectAll()
        If (sel.Text.IndexOf(String.Format("Property {0}()", propertyName)) = -1) Then
            pt.MoveToPoint(pt.CodeElement(vsCMElement.vsCMElementClass).GetEndPoint(vsCMPart.vsCMPartWhole))

            pt.LineUp()
            pt.EndOfLine()
            pt.Insert(vbNewLine)

            Dim code As StringBuilder = New StringBuilder()
            code.Append(GeneratePreProperty(memberName, propertyName, type))
            code.Append(vbNewLine)
            code.Append(GenerateProperty(memberName, propertyName, type))
            pt.Insert(code.ToString())
            GeneratePostProperty(memberName, propertyName, type, pt.Line())
        End If
    End Function

    ''' function that is executed after inserting a property
    Function GeneratePostProperty(ByVal memberName As String, ByVal propertyName As String, ByVal type As String, ByVal line As Integer)

        Dim win As Window = DTE.Windows.Item(EnvDTE.Constants.vsWindowKindTaskList)
        Dim fileName As String = DTE.ActiveWindow.Document.FullName
        Dim taskList As TaskList = win.Object

        taskList.TaskItems.Add("Generated Code", "", "Review the automatically generated code", vsTaskPriority.vsTaskPriorityMedium, vbNull, True, fileName, line - 8, True, True)
    End Function


    ''' this function tries to guess property name
    ''' from the memberName
    Function GetPropertyName(ByVal memberName) As String
        Dim propertyName As String = memberName

        If (memberName.StartsWith("_")) Then
            propertyName = propertyName.Remove(0, 1)
        End If
        ' make sure the first character is in upper case
        Dim pc As Char() = propertyName.ToCharArray()
        pc(0) = pc(0).ToUpper(pc(0))
        propertyName = New String(pc)

        Return propertyName
    End Function

    ''' main macro code
    Sub PropertyMacro()
        Dim memberName As String
        Dim type As String

        If (IsValidMember(memberName, type)) Then
            Debug.WriteLine("Got valid member name " & memberName & ", type " & type)
            ' get propertyName based on memberName
            Dim propertyName As String = GetPropertyName(memberName)
            ' finally, check that memberName and propertyName are not the same
            ' if they are, add P as the first character of propertyName

            If (memberName.ToUpper() = propertyName.ToUpper()) Then
                propertyName = "P" & propertyName
            End If

            ' propertyName variable now holds the correct property value
            Debug.WriteLine("Member name " & memberName & ", Property name " & propertyName)

            InsertProperty(memberName, propertyName, type)
        End If
    End Sub

End Module
