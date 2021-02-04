﻿//************************************************************************************************
// Copyright © 2018 Steven M Cohn.  All rights reserved.
//************************************************************************************************

namespace River.OneMoreAddIn.Commands
{
	using System;
	using Resx = River.OneMoreAddIn.Properties.Resources;


	internal partial class SearchAndReplaceDialog : UI.LocalizableForm
	{
		public SearchAndReplaceDialog ()
		{
			InitializeComponent();

			if (NeedsLocalizing())
			{
				Text = Resx.SearchAndReplaceDialog_Text;

				Localize(new string[]
				{
					"whatLabel",
					"withLabel",
					"matchBox",
					"regBox",
					"okButton",
					"cancelButton"
				});
			}
		}


		public bool MatchCase => matchBox.Checked;

		public string WithText => withBox.Text;

		public bool UseRegex => regBox.Checked;


		public string WhatText
		{
			get => whatBox.Text;
			set => whatBox.Text = value;
		}


		private void SearchAndReplaceDialog_Shown (object sender, EventArgs e)
		{
			UIHelper.SetForegroundWindow(this);
			whatBox.Focus();
		}

		private void WTextChanged (object sender, EventArgs e)
		{
			okButton.Enabled = whatBox.Text.Length > 0;
				//whatBox.Text.Length > 0 && 
				//withBox.Text.Length > 0;
		}
	}
}


