using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace DaveSexton.SettingsSwitcher.Vsix
{
	[ClassInterface(ClassInterfaceType.AutoDual)]
	[ComVisible(true)]
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Interoperability", "CA1408:DoNotUseAutoDualClassInterfaceType", Justification = "Taken from official tutorial and I don't know any better.")]
	public sealed class OptionsDialogPage : DialogPage
	{
		[Category("Export Current Settings")]
		[DisplayName("Overwrite Selected Settings File")]
		[Description("Indicates whether exporting the current settings overwrites the selected settings file or prompts to enter a file path.")]
		[DefaultValue(false)]
		public bool ExportAutomaticallyOverwritesSelectedSettingsFile
		{
			get;
			set;
		}

		[Category("Solution")]
		[DisplayName("Synchronize Selected Settings")]
		[Description("Indicates whether the currently selected settings file is associated with the current solution and is "
							 + "automatically applied whenever the solution is opened.")]
		[DefaultValue(true)]
		public bool SynchronizeSelectedSettingsWithSolution
		{
			get;
			set;
		}

		public OptionsDialogPage()
		{
			SynchronizeSelectedSettingsWithSolution = true;
		}
	}
}