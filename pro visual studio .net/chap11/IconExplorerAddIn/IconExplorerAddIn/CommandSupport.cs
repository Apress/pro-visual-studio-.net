using System;
using Microsoft.Office.Core;
using Extensibility;
using System.Runtime.InteropServices;
using EnvDTE;

using System.Windows.Forms;


namespace Apress.ProVisualStudio.chap11.IconExplorerAddIn
{
	/// <summary>
	/// Summary description for AddReplaceCommandInMenu.
	/// </summary>
	public class AddCommand
	{
		public AddCommand()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		// find a command in the commands collection
		public static EnvDTE.Command FindCommand(EnvDTE.Commands commands, string commandName) 
		{
			// first find this command if it exists, then delete it
			System.Collections.IEnumerator enm = commands.GetEnumerator();
			while(enm.MoveNext()) 
			{
				Command cmd = (Command) enm.Current;
				if (cmd.Name == commandName) 
				{
					return cmd;
				}
			}
			return null;  // command not found
		}

		public static EnvDTE.Command ReplaceCommandInMenu(
			EnvDTE.AddIn addInInstance, 
			EnvDTE._DTE applicationObject, 
			int iconNumber, 
			string progID,
			string commandName, 
			string commandButtonText, 
			string commandDescription,
			string commandBarName
			) 
		{
			object []contextGUIDS = new object[] { };
			Commands commands = applicationObject.Commands;
			_CommandBars commandBars = applicationObject.CommandBars;
			
			try 
			{	
				Command command = FindCommand(commands, progID + "." + commandName);
				if (command != null)  // if we found one
				{
					command.Delete();	// delete it as we're going to replace it
				}
				// call commands.AddNamedCommand to add it
				command = commands.AddNamedCommand(
					addInInstance, commandName, commandButtonText, 
					commandDescription, true, iconNumber, ref contextGUIDS, 
					(int)vsCommandStatus.vsCommandStatusSupported+
					(int)vsCommandStatus.vsCommandStatusEnabled);

				CommandBar commandBar = (CommandBar)commandBars[commandBarName];
				CommandBarControl commandBarControl = command.AddControl(commandBar, 1);

				return command;
			}
			catch(System.Exception e)
			{
				MessageBox.Show("Exception in ReplaceCommandInMenu: " + e.Message);
				return null;
			}
		}
	}
}
