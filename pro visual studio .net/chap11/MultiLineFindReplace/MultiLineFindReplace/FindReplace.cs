using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Diagnostics;

using EnvDTE;	// needed to reference Add-In DTE object.

namespace Apress.ProVisualStudio.chap11.MultiLineFindReplace
{
	/// <summary>
	/// Summary description for FindReplace.
	/// </summary>
	public class FindReplaceForm : System.Windows.Forms.Form
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.CheckBox checkMatchCase;
		private System.Windows.Forms.CheckBox checkMatchWholeWord;
		private System.Windows.Forms.CheckBox checkRegularExpression;
		private System.Windows.Forms.CheckBox checkWildcards;
		private System.Windows.Forms.Label FindLabel;
		private System.Windows.Forms.Label ReplaceLabel;
		private System.Windows.Forms.Label SearchPathLabel;
		private System.Windows.Forms.Label FileTypeLabel;
		private System.Windows.Forms.CheckBox checkDisplayFind2;
		private System.Windows.Forms.CheckBox checkSearchSubfolders;
		private System.Windows.Forms.TextBox findBox;
		private System.Windows.Forms.Button replaceButton;
		private System.Windows.Forms.Button closeButton;
		private System.Windows.Forms.TextBox replaceBox;
		private System.Windows.Forms.Button findButton;
		private System.Windows.Forms.TextBox searchPathBox;
		private System.Windows.Forms.TextBox filesOfTypeBox;
		private System.Windows.Forms.ComboBox lookInComboBox;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button searchPathBrowse;

		private _DTE applicationObject = null;

		private EnvDTE.vsFindTarget findTarget;

		public FindReplaceForm(_DTE applicationObject)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
			this.applicationObject = applicationObject;
			findTarget = EnvDTE.vsFindTarget.vsFindTargetSolution; // default to find/replace in solution

		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.replaceButton = new System.Windows.Forms.Button();
			this.closeButton = new System.Windows.Forms.Button();
			this.findBox = new System.Windows.Forms.TextBox();
			this.replaceBox = new System.Windows.Forms.TextBox();
			this.FindLabel = new System.Windows.Forms.Label();
			this.ReplaceLabel = new System.Windows.Forms.Label();
			this.findButton = new System.Windows.Forms.Button();
			this.checkMatchCase = new System.Windows.Forms.CheckBox();
			this.checkMatchWholeWord = new System.Windows.Forms.CheckBox();
			this.checkRegularExpression = new System.Windows.Forms.CheckBox();
			this.checkWildcards = new System.Windows.Forms.CheckBox();
			this.searchPathBox = new System.Windows.Forms.TextBox();
			this.SearchPathLabel = new System.Windows.Forms.Label();
			this.FileTypeLabel = new System.Windows.Forms.Label();
			this.filesOfTypeBox = new System.Windows.Forms.TextBox();
			this.checkDisplayFind2 = new System.Windows.Forms.CheckBox();
			this.checkSearchSubfolders = new System.Windows.Forms.CheckBox();
			this.lookInComboBox = new System.Windows.Forms.ComboBox();
			this.label1 = new System.Windows.Forms.Label();
			this.searchPathBrowse = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// replaceButton
			// 
			this.replaceButton.Location = new System.Drawing.Point(192, 352);
			this.replaceButton.Name = "replaceButton";
			this.replaceButton.Size = new System.Drawing.Size(88, 40);
			this.replaceButton.TabIndex = 2;
			this.replaceButton.Text = "Replace";
			this.replaceButton.Click += new System.EventHandler(this.ReplaceButton_Click);
			// 
			// closeButton
			// 
			this.closeButton.Location = new System.Drawing.Point(296, 352);
			this.closeButton.Name = "closeButton";
			this.closeButton.Size = new System.Drawing.Size(75, 40);
			this.closeButton.TabIndex = 3;
			this.closeButton.Text = "Close";
			this.closeButton.Click += new System.EventHandler(this.DoneButton_Click);
			// 
			// findBox
			// 
			this.findBox.Location = new System.Drawing.Point(88, 16);
			this.findBox.Multiline = true;
			this.findBox.Name = "findBox";
			this.findBox.Size = new System.Drawing.Size(296, 112);
			this.findBox.TabIndex = 6;
			this.findBox.Text = "";
			// 
			// replaceBox
			// 
			this.replaceBox.Location = new System.Drawing.Point(88, 136);
			this.replaceBox.Multiline = true;
			this.replaceBox.Name = "replaceBox";
			this.replaceBox.Size = new System.Drawing.Size(296, 112);
			this.replaceBox.TabIndex = 7;
			this.replaceBox.Text = "";
			// 
			// FindLabel
			// 
			this.FindLabel.Location = new System.Drawing.Point(24, 56);
			this.FindLabel.Name = "FindLabel";
			this.FindLabel.Size = new System.Drawing.Size(40, 40);
			this.FindLabel.TabIndex = 8;
			this.FindLabel.Text = "Find What:";
			// 
			// ReplaceLabel
			// 
			this.ReplaceLabel.Location = new System.Drawing.Point(24, 184);
			this.ReplaceLabel.Name = "ReplaceLabel";
			this.ReplaceLabel.Size = new System.Drawing.Size(56, 40);
			this.ReplaceLabel.TabIndex = 9;
			this.ReplaceLabel.Text = "Replace with:";
			// 
			// findButton
			// 
			this.findButton.Location = new System.Drawing.Point(88, 352);
			this.findButton.Name = "findButton";
			this.findButton.Size = new System.Drawing.Size(88, 40);
			this.findButton.TabIndex = 10;
			this.findButton.Text = "Find";
			this.findButton.Click += new System.EventHandler(this.FindButton_Click);
			// 
			// checkMatchCase
			// 
			this.checkMatchCase.Location = new System.Drawing.Point(408, 24);
			this.checkMatchCase.Name = "checkMatchCase";
			this.checkMatchCase.TabIndex = 11;
			this.checkMatchCase.Text = "Match Case";
			// 
			// checkMatchWholeWord
			// 
			this.checkMatchWholeWord.Location = new System.Drawing.Point(408, 48);
			this.checkMatchWholeWord.Name = "checkMatchWholeWord";
			this.checkMatchWholeWord.Size = new System.Drawing.Size(128, 24);
			this.checkMatchWholeWord.TabIndex = 12;
			this.checkMatchWholeWord.Text = "Match Whole Word";
			// 
			// checkRegularExpression
			// 
			this.checkRegularExpression.Location = new System.Drawing.Point(408, 72);
			this.checkRegularExpression.Name = "checkRegularExpression";
			this.checkRegularExpression.Size = new System.Drawing.Size(184, 24);
			this.checkRegularExpression.TabIndex = 13;
			this.checkRegularExpression.Text = "Use Regular Expressions";
			// 
			// checkWildcards
			// 
			this.checkWildcards.Location = new System.Drawing.Point(408, 96);
			this.checkWildcards.Name = "checkWildcards";
			this.checkWildcards.Size = new System.Drawing.Size(184, 24);
			this.checkWildcards.TabIndex = 14;
			this.checkWildcards.Text = "Use Wildcards";
			// 
			// searchPathBox
			// 
			this.searchPathBox.Location = new System.Drawing.Point(88, 264);
			this.searchPathBox.Name = "searchPathBox";
			this.searchPathBox.Size = new System.Drawing.Size(296, 20);
			this.searchPathBox.TabIndex = 15;
			this.searchPathBox.Text = "";
			// 
			// SearchPathLabel
			// 
			this.SearchPathLabel.Location = new System.Drawing.Point(24, 256);
			this.SearchPathLabel.Name = "SearchPathLabel";
			this.SearchPathLabel.Size = new System.Drawing.Size(56, 40);
			this.SearchPathLabel.TabIndex = 16;
			this.SearchPathLabel.Text = "Search Path:";
			// 
			// FileTypeLabel
			// 
			this.FileTypeLabel.Location = new System.Drawing.Point(24, 296);
			this.FileTypeLabel.Name = "FileTypeLabel";
			this.FileTypeLabel.Size = new System.Drawing.Size(56, 40);
			this.FileTypeLabel.TabIndex = 18;
			this.FileTypeLabel.Text = "Files of Type:";
			// 
			// filesOfTypeBox
			// 
			this.filesOfTypeBox.Location = new System.Drawing.Point(88, 296);
			this.filesOfTypeBox.Name = "filesOfTypeBox";
			this.filesOfTypeBox.Size = new System.Drawing.Size(296, 20);
			this.filesOfTypeBox.TabIndex = 21;
			this.filesOfTypeBox.Text = "";
			// 
			// checkDisplayFind2
			// 
			this.checkDisplayFind2.Location = new System.Drawing.Point(408, 176);
			this.checkDisplayFind2.Name = "checkDisplayFind2";
			this.checkDisplayFind2.Size = new System.Drawing.Size(184, 24);
			this.checkDisplayFind2.TabIndex = 19;
			this.checkDisplayFind2.Text = "Display in Find 2";
			// 
			// checkSearchSubfolders
			// 
			this.checkSearchSubfolders.Location = new System.Drawing.Point(408, 136);
			this.checkSearchSubfolders.Name = "checkSearchSubfolders";
			this.checkSearchSubfolders.Size = new System.Drawing.Size(184, 24);
			this.checkSearchSubfolders.TabIndex = 20;
			this.checkSearchSubfolders.Text = "Search Subfolders";
			// 
			// lookInComboBox
			// 
			this.lookInComboBox.Items.AddRange(new object[] {
																"Entire Solution",
																"Current Project",
																"Current Document",
																"Current Document Selection",
																"Current Document Function",
																"All Open Documents",
																"Search Path (specified directory)"});
			this.lookInComboBox.Location = new System.Drawing.Point(456, 216);
			this.lookInComboBox.Name = "lookInComboBox";
			this.lookInComboBox.Size = new System.Drawing.Size(128, 21);
			this.lookInComboBox.TabIndex = 22;
			this.lookInComboBox.Text = "Entire Solution";
			this.lookInComboBox.SelectedIndexChanged += new System.EventHandler(this.lookInComboBox_SelectedIndexChanged);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(408, 216);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 40);
			this.label1.TabIndex = 23;
			this.label1.Text = "Look In:";
			// 
			// searchPathBrowse
			// 
			this.searchPathBrowse.Location = new System.Drawing.Point(384, 264);
			this.searchPathBrowse.Name = "searchPathBrowse";
			this.searchPathBrowse.Size = new System.Drawing.Size(24, 24);
			this.searchPathBrowse.TabIndex = 25;
			this.searchPathBrowse.Text = "...";
			this.searchPathBrowse.Click += new System.EventHandler(this.searchPathBrowse_Click);
			// 
			// FindReplaceForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(600, 430);
			this.Controls.Add(this.searchPathBrowse);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lookInComboBox);
			this.Controls.Add(this.checkSearchSubfolders);
			this.Controls.Add(this.checkDisplayFind2);
			this.Controls.Add(this.FileTypeLabel);
			this.Controls.Add(this.filesOfTypeBox);
			this.Controls.Add(this.SearchPathLabel);
			this.Controls.Add(this.searchPathBox);
			this.Controls.Add(this.checkWildcards);
			this.Controls.Add(this.checkRegularExpression);
			this.Controls.Add(this.checkMatchWholeWord);
			this.Controls.Add(this.checkMatchCase);
			this.Controls.Add(this.findButton);
			this.Controls.Add(this.ReplaceLabel);
			this.Controls.Add(this.FindLabel);
			this.Controls.Add(this.replaceBox);
			this.Controls.Add(this.findBox);
			this.Controls.Add(this.closeButton);
			this.Controls.Add(this.replaceButton);
			this.Name = "FindReplaceForm";
			this.Text = "MultiLine Find & Replace in Files";
			this.ResumeLayout(false);

		}
		#endregion

		
		private void FindButton_Click(object sender, System.EventArgs e)
		{
			doFindReplace(EnvDTE.vsFindAction.vsFindActionFindAll);
		}

		private void ReplaceButton_Click(object sender, System.EventArgs e)
		{
			doFindReplace(EnvDTE.vsFindAction.vsFindActionReplaceAll);
		}

		private void doFindReplace(EnvDTE.vsFindAction findAction)
		{
			int findOptions = 0;

			EnvDTE.vsFindResultsLocation findResultsLocation;
			EnvDTE.vsFindResult findResult;

			if (checkMatchCase.Checked) 
			{
				findOptions |= (int) EnvDTE.vsFindOptions.vsFindOptionsMatchCase;  // set this bit in options
			} 
			else
			{
				findOptions &= ~ (int) EnvDTE.vsFindOptions.vsFindOptionsMatchCase; // clear this bit in options

			}

			if (checkMatchWholeWord.Checked) 
			{
				findOptions |= (int) EnvDTE.vsFindOptions.vsFindOptionsMatchWholeWord;  // set this bit in options
			} 
			else
			{
				findOptions &= ~ (int) EnvDTE.vsFindOptions.vsFindOptionsMatchWholeWord; // clear this bit in options

			}

			if (checkRegularExpression.Checked) 
			{
				findOptions |= (int) EnvDTE.vsFindOptions.vsFindOptionsRegularExpression;  // set this bit in options
			} 
			else
			{
				findOptions &= ~ (int) EnvDTE.vsFindOptions.vsFindOptionsRegularExpression; // clear this bit in options
			}

			if (checkWildcards.Checked) 
			{
				findOptions |= (int) EnvDTE.vsFindOptions.vsFindOptionsWildcards;  // set this bit in options
			} 
			else
			{
				findOptions &= ~ (int) EnvDTE.vsFindOptions.vsFindOptionsWildcards; // clear this bit in options
			}	

			if (checkSearchSubfolders.Checked) 
			{
				findOptions |= (int) EnvDTE.vsFindOptions.vsFindOptionsSearchSubfolders;  // set this bit in options
			} 
			else
			{
				findOptions &= ~ (int) EnvDTE.vsFindOptions.vsFindOptionsSearchSubfolders; // clear this bit in options
			}	

			// these options don't make sense in find-in-files context
			// EnvDTE.vsFindOptions.vsFindOptionsBackwards;
			// EnvDTE.vsFindOptions.vsFindOptionsFromStart;
			//EnvDTE.vsFindOptions.vsFindOptionsKeepModifiedDocumentsOpen;

			// these options just didn't seem worthwhile to implement
			//EnvDTE.vsFindOptions.vsFindOptionsNone;
			//EnvDTE.vsFindOptions.vsFindOptionsMatchInHiddenText;

			if (checkDisplayFind2.Checked) 
			{
				findResultsLocation = EnvDTE.vsFindResultsLocation.vsFindResults2;
			} 
			else
			{
				findResultsLocation = EnvDTE.vsFindResultsLocation.vsFindResults1;
			}		
			// not giving option to specify no display of results
			// (EnvDTE.vsFindResultsLocation.vsFindResultsNone)

			try 
			{
				// this method does all the hard work for us
				findResult = applicationObject.DTE.Find.FindReplace(
					findAction, findBox.Text,
					findOptions, replaceBox.Text,
					findTarget, searchPathBox.Text, 
					filesOfTypeBox.Text,
					findResultsLocation);

				Debug.WriteLine("Find Result: " + findResult.ToString());
			} 
			catch (System.Exception e) 
			{
				MessageBox.Show("Exception from FindReplace method: " + e.Message);
			}
		}

		private void DoneButton_Click(object sender, System.EventArgs e)
		{
			// dispose of the find/replace form, we are done for now
			this.Dispose();		
		}

		private void searchPathBrowse_Click(object sender, System.EventArgs e)
		{
			System.Windows.Forms.FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
			folderBrowser.ShowDialog();
			searchPathBox.Text = folderBrowser.SelectedPath;
			//folderBrowser.Dispose();
		}

		private void lookInComboBox_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			Debug.WriteLine("lookInComboBox = " + lookInComboBox.SelectedIndex.ToString());

			switch (lookInComboBox.SelectedIndex) 
			{
				case 0: findTarget = EnvDTE.vsFindTarget.vsFindTargetSolution; return; // search all files in current solution
				case 1: findTarget = EnvDTE.vsFindTarget.vsFindTargetCurrentProject; return; // search all files in current project
				case 2: findTarget = EnvDTE.vsFindTarget.vsFindTargetCurrentDocument; return; // search just current file
				case 3: findTarget = EnvDTE.vsFindTarget.vsFindTargetCurrentDocumentFunction; return; // search just current function in current document
				case 4: findTarget = EnvDTE.vsFindTarget.vsFindTargetCurrentDocumentSelection; return; // search just current selection in current document
				case 5: findTarget = EnvDTE.vsFindTarget.vsFindTargetOpenDocuments; return; // search all open files
				case 6: findTarget = EnvDTE.vsFindTarget.vsFindTargetFiles; return; // search only files specified in search path box
				default: MessageBox.Show("Error in MultiLineFindReplace: " +
										 "findTarget out of range; " +
										  "lookInComboBox.SelectedIndex = " +
											lookInComboBox.SelectedIndex.ToString()
										); return;
							 
			}
		}







		

	
	}
}
