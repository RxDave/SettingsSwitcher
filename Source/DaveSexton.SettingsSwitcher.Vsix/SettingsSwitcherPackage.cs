using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Windows;
using System.Windows.Threading;
using DaveSexton.SettingsSwitcher.Vsix.Properties;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using Ole = Microsoft.VisualStudio.OLE.Interop;

namespace DaveSexton.SettingsSwitcher.Vsix
{
	/// <summary>
	/// Implements the package exposed by this assembly.
	/// </summary>
	/// <remarks>
	/// The minimum requirement for a class to be considered a valid package for Visual Studio
	/// is to implement the IVsPackage interface and register itself with the shell.
	/// This package uses the helper classes defined inside the Managed Package Framework (MPF)
	/// to do it: it derives from the Package class that provides the implementation of the 
	/// IVsPackage interface and uses the registration attributes defined in the framework to 
	/// register itself and its components with the shell.
	/// </remarks>
	[PackageRegistration(UseManagedResourcesOnly = true)]		// This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is a package.
	[InstalledProductRegistration("#110", "#112", GlobalConstants.MajorMinorVersion, IconResourceID = 400)]		// This attribute is used to register the information needed to show this package in the Help/About dialog of Visual Studio.
	[ProvideMenuResource("Menus.ctmenu", 1)]	// This attribute is needed to let the shell know that this package exposes some menus and toolbars.
	[ProvideOptionPage(typeof(OptionsDialogPage), packageTitle, "General", 113, 114, supportsAutomation: true)]
	[ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]
	[ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
	[Guid(Guids.vsixPkgString)]
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling", Justification = "Time constraints.")]
	public sealed class SettingsSwitcherPackage : Package
	{
		public Uri SettingsDirectory
		{
			get
			{
				Contract.Ensures(Contract.Result<Uri>() == null || Contract.Result<Uri>().IsAbsoluteUri);
				Contract.Ensures(Contract.Result<Uri>() == null || Contract.Result<Uri>().IsFile);
				Contract.Ensures(Contract.Result<Uri>() == settingsDirectory);

				return settingsDirectory;
			}
		}

		private Uri CurrentSolutionSettingsFile
		{
			get
			{
				Contract.Ensures(Contract.Result<Uri>() == null || Contract.Result<Uri>().IsAbsoluteUri);
				Contract.Ensures(Contract.Result<Uri>() == null || Contract.Result<Uri>().IsFile);

				var solution = Dte.Solution;

				string solutionFile;
				if (solution == null || !solution.IsOpen || string.IsNullOrWhiteSpace(solutionFile = solution.FullName))
				{
					return null;
				}
				else
				{
					Contract.Assert(!string.IsNullOrWhiteSpace(solutionFile));

					var path = new Uri(SolutionSettingsFile.GetFullName(solutionFile, settingsFileExtension), UriKind.Absolute);

					Contract.Assume(path.IsAbsoluteUri);
					Contract.Assume(path.IsFile);

					return path;
				}
			}
		}

		private EnvDTE.DTE Dte
		{
			get
			{
				Contract.Ensures(Contract.Result<EnvDTE.DTE>() != null);

				if (dte == null)
				{
					dte = (EnvDTE.DTE)GetService(typeof(SDTE));

					Contract.Assume(dte != null);
				}

				return dte;
			}
		}

		private EnvDTE.Property AutoSaveFileProperty
		{
			get
			{
				Contract.Ensures(Contract.Result<EnvDTE.Property>() != null);

				var page = Dte.Properties["Environment", "Import and Export Settings"];

				Contract.Assume(page != null);

				var autoSaveFileProperty = page.Item("AutoSaveFile");

				Contract.Assume(autoSaveFileProperty != null);

				return autoSaveFileProperty;
			}
		}

		private CommandBar ToolBar
		{
			get
			{
				var bars = (CommandBars)Dte.CommandBars;

				if (bars != null)
				{
					try
					{
						var bar = bars[primaryCommandBarCaption];

						if (bar != null)
						{
							bar.Reset();
						}

						return bar;
					}
					catch (COMException)
					{
					}
				}

				return null;
			}
		}

		private OptionsDialogPage Options
		{
			get
			{
				Contract.Ensures(Contract.Result<OptionsDialogPage>() != null);

				if (options == null)
				{
					options = (OptionsDialogPage)GetDialogPage(typeof(OptionsDialogPage));

					Contract.Assume(options != null);
				}

				return options;
			}
		}

		private const string packageTitle = "Settings Switcher";
		private const string primaryCommandBarCaption = packageTitle;
		private const string settingsFileExtension = ".vssettings";
		private const string settingsFilePattern = "*" + settingsFileExtension;
		private const string oldAutoSettingsFilePrefix = "Old.";
		private const string originalSettingsFileName = "OriginalSettings" + settingsFileExtension;

		private static readonly string SettingsOptionKey = typeof(SettingsSwitcherPackage).Name;

		private readonly FileSystemWatcher settingFilesWatcher = new FileSystemWatcher()
		{
			Filter = settingsFilePattern,
			IncludeSubdirectories = true,
			NotifyFilter = NotifyFilters.CreationTime | NotifyFilters.FileName | NotifyFilters.LastWrite
		};

		private readonly object gate = new object();
		private readonly SettingsFileCollection settingsFiles = new SettingsFileCollection();
		private SettingsFileNameManager nameManager;
		private EnvDTE.DTE dte;
		private EnvDTE.SolutionEvents solutionEvents;
		private EnvDTE.DTEEvents dteEvents;
		private SettingsFile selectedSettingsFile, unsavedSelectedSettingsFile;
		private Uri settingsDirectory, autoSettingsFile, oldAutoSettingsFile, originalSettingsFile;
		private bool initialized, executingManageUserSettingsCommand, loadedOptions, unloadingPackage;
		private OptionsDialogPage options;
		private MsoBarPosition toolbarPosition;
		private int toolbarRowIndex;

		/// <summary>
		/// Default constructor of the package.
		/// Inside this method you can place any initialization code that does not require 
		/// any Visual Studio service because at this point the package object is created but 
		/// not sited yet inside Visual Studio environment. The place to do all the other 
		/// initialization is the Initialize method.
		/// </summary>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "The settingFilesWatcher field is disposed in the Dispose method of this class.")]
		public SettingsSwitcherPackage()
		{
			Debug.WriteLine("Entering constructor for: {0}", this);

			settingFilesWatcher.Changed += SettingsFileDeletedCreatedChangedOrRenamed;
			settingFilesWatcher.Created += SettingsFileDeletedCreatedChangedOrRenamed;
			settingFilesWatcher.Deleted += SettingsFileDeletedCreatedChangedOrRenamed;
			settingFilesWatcher.Renamed += SettingsFileDeletedCreatedChangedOrRenamed;
			settingFilesWatcher.Error += SettingFilesWatcherError;
		}

		[ContractInvariantMethod]
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Required for code contracts.")]
		private void ObjectInvariant()
		{
			Contract.Invariant(settingsFiles != null);
			Contract.Invariant(settingFilesWatcher != null);
			Contract.Invariant(settingsDirectory == null || settingsDirectory.IsAbsoluteUri);
			Contract.Invariant(settingsDirectory == null || settingsDirectory.IsFile);
			Contract.Invariant((nameManager == null) == (settingsDirectory == null));
			Contract.Invariant(autoSettingsFile == null || autoSettingsFile.IsAbsoluteUri);
			Contract.Invariant(autoSettingsFile == null || autoSettingsFile.IsFile);
			Contract.Invariant((autoSettingsFile == null) == (settingsDirectory == null));
			Contract.Invariant(oldAutoSettingsFile == null || oldAutoSettingsFile.IsAbsoluteUri);
			Contract.Invariant(oldAutoSettingsFile == null || oldAutoSettingsFile.IsFile);
			Contract.Invariant((oldAutoSettingsFile == null) == (settingsDirectory == null));
			Contract.Invariant(originalSettingsFile == null || originalSettingsFile.IsAbsoluteUri);
			Contract.Invariant(originalSettingsFile == null || originalSettingsFile.IsFile);
			Contract.Invariant((originalSettingsFile == null) == (settingsDirectory == null));
		}

		/// <summary>
		/// Initialization of the package; this method is called right after the package is sited, so this is the place
		/// where you can put all the initialization code that rely on services provided by VisualStudio.
		/// </summary>
		protected override void Initialize()
		{
			Contract.Ensures(initialized);

			Debug.WriteLine("Entering Initialize() of: {0}", this);
			Debug.WriteLine("Default selected settings file: " + Settings.Default.SelectedSettingsFile);

			base.Initialize();

			AddOptionKey(SettingsOptionKey);

			ReadSettingsDirectory();

			if (settingsDirectory != null)
			{
				var events = Dte.Events;

				Contract.Assume(events != null);

				// A reference to solutionEvents must be stored in a field so that event handlers aren't collected by the GC.
				solutionEvents = events.SolutionEvents;
				dteEvents = events.DTEEvents;

				Contract.Assume(solutionEvents != null);
				Contract.Assume(dteEvents != null);

				solutionEvents.Opened += SolutionOpened;
				solutionEvents.Renamed += SolutionRenamed;
				solutionEvents.BeforeClosing += SolutionClosing;

				dteEvents.OnBeginShutdown += BeginShutdown;

				settingFilesWatcher.Path = settingsDirectory.LocalPath;
				settingFilesWatcher.EnableRaisingEvents = true;

				lock (gate)
				{
					LoadSettingsFiles();

					if (selectedSettingsFile == null)
					{
						var defaultFile = Settings.Default.SelectedSettingsFile;

						if (string.IsNullOrWhiteSpace(defaultFile) || !File.Exists(defaultFile) || !settingsDirectory.IsBaseOf(new Uri(defaultFile, UriKind.Absolute)))
						{
							if (File.Exists(originalSettingsFile.LocalPath))
							{
								Debug.WriteLine("Importing original settings file: " + originalSettingsFile.LocalPath);

								ImportSettings(originalSettingsFile.LocalPath);
							}
							else
							{
								Debug.WriteLine("Exporting original settings file: " + originalSettingsFile.LocalPath);

								ExportCurrentSettings(originalSettingsFile.LocalPath);
							}
						}
						else
						{
							Debug.WriteLine("Default settings file found: " + defaultFile);

							selectedSettingsFile = settingsFiles.FindByFullName(defaultFile);
						}
					}
				}
			}

			var menus = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;

			if (menus != null)
			{
				foreach (var command in InitializeCommands())
				{
					menus.AddCommand(command);
				}
			}

			InitializePriorityCommands();
			InitializeToolBar();

			Contract.Assert(settingsDirectory == null || settingsDirectory.IsAbsoluteUri);
			Contract.Assert(settingsDirectory == null || settingsDirectory.IsFile);
			Contract.Assert(autoSettingsFile == null || autoSettingsFile.IsAbsoluteUri);
			Contract.Assert(autoSettingsFile == null || autoSettingsFile.IsFile);
			Contract.Assert(oldAutoSettingsFile == null || oldAutoSettingsFile.IsAbsoluteUri);
			Contract.Assert(oldAutoSettingsFile == null || oldAutoSettingsFile.IsFile);
			Contract.Assert(originalSettingsFile == null || originalSettingsFile.IsAbsoluteUri);
			Contract.Assert(originalSettingsFile == null || originalSettingsFile.IsFile);

			initialized = true;
		}

		private IEnumerable<MenuCommand> InitializeCommands()
		{
			Contract.Ensures(Contract.Result<IEnumerable<MenuCommand>>() != null);
			Contract.Ensures(settingsDirectory == Contract.OldValue(settingsDirectory));
			Contract.Ensures(autoSettingsFile == Contract.OldValue(autoSettingsFile));
			Contract.Ensures(oldAutoSettingsFile == Contract.OldValue(oldAutoSettingsFile));
			Contract.Ensures(originalSettingsFile == Contract.OldValue(originalSettingsFile));

			var switcherCommand = new OleMenuCommand(
				SettingsSelector,
				new CommandID(Guids.VsixCmdSet, (int)PackageCommandIds.SettingsSelector));

			switcherCommand.BeforeQueryStatus += ImportSelectedSettingsQueryStatus;

			yield return switcherCommand;

			yield return new OleMenuCommand(
				SettingsSelectorGetItems,
				new CommandID(Guids.VsixCmdSet, (int)PackageCommandIds.SettingsSelectorGetItems));

			var exportCommand = new OleMenuCommand(
				ExportCurrentSettings,
				new CommandID(Guids.VsixCmdSet, (int)PackageCommandIds.ExportCurrentSettings));

			exportCommand.BeforeQueryStatus += ExportCurrentSettingsQueryStatus;

			yield return exportCommand;

			var exportSolutionCommand = new OleMenuCommand(
				ExportSolutionSettings,
				new CommandID(Guids.VsixCmdSet, (int)PackageCommandIds.ExportSolutionSettings));

			exportSolutionCommand.BeforeQueryStatus += ExportSolutionSettingsQueryStatus;

			yield return exportSolutionCommand;

			var formatSolutionCommand = new OleMenuCommand(
				FormatSolution,
				new CommandID(Guids.VsixCmdSet, (int)PackageCommandIds.FormatSolution));

			formatSolutionCommand.BeforeQueryStatus += FormatSolutionQueryStatus;

			yield return formatSolutionCommand;
		}

		private void InitializePriorityCommands()
		{
			var priority = GetService(typeof(SVsRegisterPriorityCommandTarget)) as IVsRegisterPriorityCommandTarget;

			if (priority != null)
			{
				uint cookie;
				ErrorHandler.ThrowOnFailure(priority.RegisterPriorityCommandTarget(
					0,
					new AnonymousCommandTarget(
						(groupId, commands, text) => (int)Ole.Constants.OLECMDERR_E_UNKNOWNGROUP,
						(groupId, commandId, opts, input, output) =>
						{
							if (commandId == (uint)VSConstants.VSStd2KCmdID.ManageUserSettings
								&& groupId == Guids.VsStd2kCmdId
								&& !executingManageUserSettingsCommand)
							{
								var result = VsShellUtilities.ShowMessageBox(
									this,
									Resources.ImportAndExportCommandIsIncompatible,
									Resources.ImportAndExportCommandIsIncompatibleTitle,
									OLEMSGICON.OLEMSGICON_WARNING,
									OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL,
									OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_SECOND);

								ErrorHandler.ThrowOnFailure(result);

								if (result == (int)VSConstants.MessageBoxResult.IDCANCEL)
								{
									return (int)VSConstants.S_OK;
								}
							}

							return (int)Ole.Constants.OLECMDERR_E_NOTSUPPORTED;
						}),
					out cookie));
			}
		}

		private void InitializeToolBar()
		{
			ShowToolBar();
		}

		private void ShowToolBar()
		{
			Contract.Ensures(settingsDirectory == Contract.OldValue(settingsDirectory));

			var dispatcher = Dispatcher.CurrentDispatcher;

			var afterInitialization = initialized;

			// During testing, assigning toolbar.Visible sometimes had no effect unless it's done asynchronously.
			dispatcher.InvokeAsync(() =>
			{
				var toolbar = ToolBar;

				if (toolbar != null)
				{
					if (toolbar.Visible)
					{
						toolbarPosition = toolbar.Position;
						toolbarRowIndex = toolbar.RowIndex;

						Debug.WriteLine("Settings switcher toolbar visible at {0}, index {1}.  After init: {2}", toolbarPosition, toolbarRowIndex, afterInitialization);
					}
					else
					{
						toolbar.Visible = true;

						Debug.WriteLine("Settings switcher toolbar currently reports visibility at {0}, index {1}.", toolbar.Position, toolbar.RowIndex);

						// During testing in VS 2012, setting toolbar.Visible to true after initialization sometimes caused the toolbar to reappear in a different location.
						if (afterInitialization)
						{
							Debug.WriteLine("Attempting to show settings switcher toolbar at {0}, index {1}.", toolbarPosition, toolbarRowIndex);

							if (toolbar.Position != toolbarPosition || toolbar.RowIndex != toolbarRowIndex)
							{
								// Another async invocation was required while testing in VS 2012, on a fairly slow computer.
								dispatcher.InvokeAsync(() =>
								{
									// Must recreate the toolbar reference; otherwise, setting the following properties has no effect.
									var t = ToolBar;

									Debug.WriteLine("Settings switcher toolbar currently reports visibility (before update) at {0}, index {1}.", t.Position, t.RowIndex);

									t.Position = toolbarPosition;
									t.RowIndex = toolbarRowIndex;

									Debug.WriteLine("Settings switcher toolbar visibility set to {0}, index {1}; Result: {2}, index {3}.", toolbarPosition, toolbarRowIndex, t.Position, t.RowIndex);
								},
								DispatcherPriority.Background);
							}
						}
						else
						{
							Debug.WriteLine("Saving settings switcher toolbar state.");

							toolbarPosition = toolbar.Position;
							toolbarRowIndex = toolbar.RowIndex;
						}
					}
				}
			},
			DispatcherPriority.Background);
		}

		private void ReadSettingsDirectory()
		{
			var autoSaveFileProperty = AutoSaveFileProperty;
			var autoSaveFile = autoSaveFileProperty.Value as string;

			if (string.IsNullOrWhiteSpace(autoSaveFile))
			{
				autoSaveFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
			}

			var settingsDirectoryPath = Path.GetDirectoryName(autoSaveFile) + '\\';

			settingsDirectory = new Uri(settingsDirectoryPath, UriKind.Absolute);
			autoSettingsFile = new Uri(autoSaveFile, UriKind.Absolute);
			oldAutoSettingsFile = new Uri(Path.Combine(settingsDirectoryPath, oldAutoSettingsFilePrefix + Path.GetFileName(autoSaveFile)), UriKind.Absolute);
			originalSettingsFile = new Uri(Path.Combine(settingsDirectoryPath, originalSettingsFileName), UriKind.Absolute);

			//// TODO: Special case team settings file too.

			AssumeReadSettingsDirectory();

			nameManager = new SettingsFileNameManager(autoSettingsFile, oldAutoSettingsFile, originalSettingsFile);

			AssumeReadSettingsDirectory();

			Debug.WriteLine("Settings directory: " + settingsDirectory);
			Debug.WriteLine("Auto-settings: " + autoSettingsFile);
			Debug.WriteLine("Old auto-settings: " + oldAutoSettingsFile);
			Debug.WriteLine("Original settings: " + originalSettingsFile);
		}

		[ContractVerification(false)]
		private void AssumeReadSettingsDirectory()
		{
			Contract.Ensures(settingsDirectory != null);
			Contract.Ensures(settingsDirectory.IsAbsoluteUri);
			Contract.Ensures(settingsDirectory.IsFile);
			Contract.Ensures(autoSettingsFile != null);
			Contract.Ensures(autoSettingsFile.IsAbsoluteUri);
			Contract.Ensures(autoSettingsFile.IsFile);
			Contract.Ensures(oldAutoSettingsFile != null);
			Contract.Ensures(oldAutoSettingsFile.IsAbsoluteUri);
			Contract.Ensures(oldAutoSettingsFile.IsFile);
			Contract.Ensures(originalSettingsFile != null);
			Contract.Ensures(originalSettingsFile.IsAbsoluteUri);
			Contract.Ensures(originalSettingsFile.IsFile);
		}

		private void LoadSettingsFiles(bool ignoreSolutionSettingsFile = false)
		{
			Contract.Requires(settingsDirectory != null);
			Contract.Ensures(settingsDirectory != null);

			Contract.Assert(autoSettingsFile != null);

			settingsFiles.Clear();
			nameManager.Clear();

			var previousSelection = selectedSettingsFile;

			selectedSettingsFile = null;

			Debug.WriteLine("Loading settings files.  Current selection: " + (previousSelection == null ? null : previousSelection.FullPath));

			try
			{
				var currentSolutionSettingsFile = CurrentSolutionSettingsFile;

				var files = from file in Directory.EnumerateFiles(settingsDirectory.LocalPath, settingsFilePattern, SearchOption.AllDirectories)
										let normalized = new Uri(file, UriKind.Absolute)
										where normalized != currentSolutionSettingsFile		// Probably not required, since solution settings files are next to .sln files, not in settingsDirectory.
										select new SettingsFile(normalized, settingsDirectory, nameManager);

				if (!ignoreSolutionSettingsFile && currentSolutionSettingsFile != null && File.Exists(currentSolutionSettingsFile.LocalPath))
				{
					Contract.Assert(currentSolutionSettingsFile.IsAbsoluteUri);
					Contract.Assert(currentSolutionSettingsFile.IsFile);

					files = files.Concat(new[] { new SolutionSettingsFile(currentSolutionSettingsFile, nameManager) });
				}

				foreach (var settingsFile in from settingsFile in files
																		 orderby settingsFile.LastAccessed descending
																		 select settingsFile)
				{
					settingsFiles.Add(settingsFile);
				}
			}
			catch (IOException ex)
			{
				Debug.WriteLine("Failed to iterate files: " + ex, packageTitle);
			}
			catch (SecurityException ex)
			{
				Debug.WriteLine("Failed to iterate files: " + ex, packageTitle);
			}

			if (previousSelection != null)
			{
				selectedSettingsFile = settingsFiles.FirstOrDefault(file => file.FullPath == previousSelection.FullPath);
			}

			Debug.WriteLine("Settings files loaded.  Current selection: " + (selectedSettingsFile == null ? null : selectedSettingsFile.FullPath));
		}

		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Contracts", "Invariant-123-259", Justification = "False positive.")]
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Contracts", "Ensures-settingsDirectory != null", Justification = "False positive.")]
		private SolutionSettingsFile TryEnsureCurrentSolutionSettingsFile(bool mustAddForCurrentSolution = false)
		{
			Contract.Requires(settingsDirectory != null);
			Contract.Requires(nameManager != null);
			Contract.Ensures(settingsDirectory != null);

			lock (gate)
			{
				var currentSolutionSettingsFile = CurrentSolutionSettingsFile;

				var existing = currentSolutionSettingsFile == null ? null : settingsFiles.FindByFullName(currentSolutionSettingsFile.LocalPath) as SolutionSettingsFile;

				if (existing != null)
				{
					AssumeReadSettingsDirectory();

					return existing;
				}

				existing = settingsFiles.OfType<SolutionSettingsFile>().FirstOrDefault();

				if (existing != null)
				{
					settingsFiles.Remove(existing);
					nameManager.RemoveFile(existing);
				}

				if (currentSolutionSettingsFile != null && (mustAddForCurrentSolution || File.Exists(currentSolutionSettingsFile.LocalPath)))
				{
					var file = new SolutionSettingsFile(currentSolutionSettingsFile, nameManager);

					settingsFiles.Add(file);

					AssumeReadSettingsDirectory();

					return file;
				}

				AssumeReadSettingsDirectory();

				return null;
			}
		}

		private void UnloadSolutionSettingsFile()
		{
			Contract.Requires(settingsDirectory != null);
			Contract.Requires(nameManager != null);
			Contract.Ensures(settingsDirectory != null);

			if (!unloadingPackage)
			{
				unsavedSelectedSettingsFile = selectedSettingsFile;
			}

			var solutionSettingsFile = settingsFiles.OfType<SolutionSettingsFile>().FirstOrDefault();

			if (solutionSettingsFile != null)
			{
				var importRequired = selectedSettingsFile == solutionSettingsFile;

				if (importRequired)
				{
					selectedSettingsFile = settingsFiles.FindByFullName(Settings.Default.SelectedSettingsFile);
				}

				if (!unloadingPackage)
				{
					LoadSettingsFiles(ignoreSolutionSettingsFile: true);
				}

				if (importRequired)
				{
					ImportSelectedSettings(saveSelectedAsDefault: false);
				}
			}
		}

		private void SelectSettingsFile(string file)
		{
			Contract.Requires(settingsDirectory != null);
			Contract.Requires(file != null);

			lock (gate)
			{
				selectedSettingsFile = settingsFiles.FindByFullName(file);

				if (selectedSettingsFile == null)
				{
					LoadSettingsFiles();

					selectedSettingsFile = settingsFiles.FindByFullName(file);
				}

				if (selectedSettingsFile != null)
				{
					selectedSettingsFile.Accessed();
				}
			}

			RefreshCommands();
		}

		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Contracts", "Ensures-settingsDirectory == Contract.OldValue(settingsDirectory)", Justification = "False positive.")]
		private void RefreshCommands()
		{
			Contract.Ensures(settingsDirectory == Contract.OldValue(settingsDirectory));

			var shell = (IVsUIShell)GetService(typeof(SVsUIShell));

			Contract.Assume(shell != null);

			ErrorHandler.ThrowOnFailure(shell.UpdateCommandUI(VSConstants.S_FALSE));
		}

		protected override void OnLoadOptions(string key, Stream stream)
		{
			Contract.Assume(key == SettingsOptionKey);
			Contract.Assume(stream != null);

			loadedOptions = true;

			if (settingsDirectory != null)
			{
				AssumeReadSettingsDirectory();

				var solutionSettingsFile = TryEnsureCurrentSolutionSettingsFile();
				var hasAssociatedFile = false;

				if (Options.SynchronizeSelectedSettingsWithSolution)
				{
					var reader = new StreamReader(stream, Encoding.UTF8);

					var path = reader.ReadToEnd();

					if (!string.IsNullOrEmpty(path))
					{
						lock (gate)
						{
							var file = settingsFiles.FindByFullName(path);

							hasAssociatedFile = file != null;

							if (hasAssociatedFile && file != selectedSettingsFile)
							{
								selectedSettingsFile = file;

								ImportSelectedSettings(saveSelectedAsDefault: false);
							}
						}

						RefreshCommands();
					}
				}

				if (!hasAssociatedFile && solutionSettingsFile != null && solutionSettingsFile.Exists)
				{
					AssumeReadSettingsDirectory();

					ImportSettings(solutionSettingsFile.FullPath.LocalPath);
				}
			}

			base.OnLoadOptions(key, stream);
		}

		protected override void OnSaveOptions(string key, Stream stream)
		{
			Contract.Assume(key == SettingsOptionKey);
			Contract.Assume(stream != null);

			if (Options.SynchronizeSelectedSettingsWithSolution)
			{
				/* The unsavedSelectedSettingsFile field is used instead of selectedSettingsFile since closing the 
				 * solution (or Visual Studio) causes this package to switch VS back to the original settings, thus 
				 * changing selectedSettingsFile to a different reference.
				 */
				var file = unsavedSelectedSettingsFile;

				if (file != null)
				{
					unsavedSelectedSettingsFile = null;

					var writer = new StreamWriter(stream, Encoding.UTF8);

					writer.Write(file is SolutionSettingsFile ? string.Empty : file.FullPath.LocalPath);
					writer.Flush();
				}
			}

			base.OnSaveOptions(key, stream);
		}

		private void SolutionOpened()
		{
			Contract.Requires(settingsDirectory != null);

			if (!loadedOptions)
			{
				var solutionSettingsFile = TryEnsureCurrentSolutionSettingsFile();

				if (solutionSettingsFile != null && solutionSettingsFile.Exists)
				{
					ImportSettings(solutionSettingsFile.FullPath.LocalPath);
				}
			}
			else
			{
				loadedOptions = false;
			}
		}

		private void SolutionRenamed(string oldName)
		{
			Contract.Requires(settingsDirectory != null);
			Contract.Requires(oldName != null);
			Contract.Ensures(settingsDirectory != null);

			var solutionSettingsFile = settingsFiles.OfType<SolutionSettingsFile>().FirstOrDefault();

			if (solutionSettingsFile != null)
			{
				var solution = Dte.Solution;

				Contract.Assume(solution != null);

				var oldFile = solutionSettingsFile.FullPath.LocalPath;

				Contract.Assume(!string.IsNullOrEmpty(oldFile));

				settingsFiles.Remove(solutionSettingsFile);
				nameManager.RemoveFile(solutionSettingsFile);

				solutionSettingsFile = TryEnsureCurrentSolutionSettingsFile(mustAddForCurrentSolution: true);

				Contract.Assume(solutionSettingsFile != null);

				var newFile = solutionSettingsFile.FullPath.LocalPath;

				Contract.Assume(!string.IsNullOrEmpty(newFile));

				var item = solution.FindProjectItem(oldFile);

				if (item != null)
				{
					item.Save(newFile);
				}
				else
				{
					File.Copy(oldFile, newFile, overwrite: true);
				}
			}
		}

		/// <remarks>
		/// Called when the solution is closing, and when the package is unloaded while a solution is opened; e.g., when closing VS.
		/// </remarks>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Contracts", "Invariant-123-69", Justification = "False positive.")]
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Contracts", "Ensures-settingsDirectory != null", Justification = "False positive.")]
		private void SolutionClosing()
		{
			Contract.Requires(settingsDirectory != null);
			Contract.Ensures(settingsDirectory != null);

			lock (gate)
			{
				UnloadSolutionSettingsFile();
			}
		}

		private void ExecuteManageUserSettingsCommand(string arguments)
		{
			Contract.Requires(arguments != null);
			Contract.Ensures(settingsDirectory == Contract.OldValue(settingsDirectory));

			try
			{
				executingManageUserSettingsCommand = true;

				Contract.Assume(Dte.Commands != null);

				Dte.Commands.Raise(Guids.vsStd2kCmdIdString, (int)VSConstants.VSStd2KCmdID.ManageUserSettings, arguments, null);
			}
			finally
			{
				executingManageUserSettingsCommand = false;
			}
		}

		private void ImportSelectedSettings(bool saveSelectedAsDefault = true)
		{
			Contract.Requires(settingsDirectory != null);
			Contract.Ensures(settingsDirectory != null);

			if (selectedSettingsFile != null)
			{
				ImportSettings(selectedSettingsFile.FullPath.LocalPath);

				if (saveSelectedAsDefault && !(selectedSettingsFile is SolutionSettingsFile))
				{
					Contract.Assume(selectedSettingsFile != null);

					Settings.Default.SelectedSettingsFile = selectedSettingsFile.FullPath.LocalPath;
				}
			}
			else if (saveSelectedAsDefault)
			{
				Settings.Default.SelectedSettingsFile = null;
			}

			if (saveSelectedAsDefault)
			{
				Settings.Default.Save();
			}
		}

		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Contracts", "Ensures-settingsDirectory != null", Justification = "False positive.")]
		private void ImportSettings(string file)
		{
			Contract.Requires(settingsDirectory != null);
			Contract.Requires(file != null);
			Contract.Ensures(settingsDirectory != null);

			SelectSettingsFile(file);

			var autoSaveFileProperty = AutoSaveFileProperty;
			var originalPath = autoSaveFileProperty.Value;
			var toolbar = ToolBar;
			var toolbarVisible = toolbar != null && toolbar.Visible;

			try
			{
				ExecuteManageUserSettingsCommand("-import:\"" + file + "\"");
			}
			finally
			{
				autoSaveFileProperty.Value = originalPath;

				if (toolbarVisible)
				{
					ShowToolBar();

					AssumeReadSettingsDirectory();

					Contract.Assume(settingsDirectory != null);
					Contract.Assume(nameManager != null);
				}
			}
		}

		private void ExportCurrentSettings(object sender, EventArgs e)
		{
			Contract.Requires(settingsDirectory != null);

			if (Options.ExportAutomaticallyOverwritesSelectedSettingsFile
				&& selectedSettingsFile != null
				&& selectedSettingsFile.FullPath != autoSettingsFile
				&& selectedSettingsFile.FullPath != oldAutoSettingsFile)
			{
				ExportCurrentSettings(selectedSettingsFile.FullPath.LocalPath);
			}
			else
			{
				var dialog = new SaveFileDialog()
				{
					AddExtension = true,
					DefaultExt = settingsFileExtension,
					Filter = Resources.SettingsFilesGroupName + " (" + settingsFilePattern + ")|" + settingsFilePattern,
					InitialDirectory = settingsDirectory.LocalPath,
					OverwritePrompt = true,
					Title = Resources.ExportCurrentSettingsDialogTitle
				};

				if (dialog.ShowDialog(Application.Current.MainWindow) ?? false)
				{
					ExportCurrentSettings(dialog.FileName);
				}
			}
		}

		private void ExportCurrentSettings(string file)
		{
			Contract.Requires(settingsDirectory != null);
			Contract.Requires(file != null);

			ExecuteManageUserSettingsCommand("-export:\"" + file + "\"");

			SelectSettingsFile(file);
		}

		private void ExportSolutionSettings(object sender, EventArgs e)
		{
			Contract.Requires(settingsDirectory != null);

			var file = TryEnsureCurrentSolutionSettingsFile(mustAddForCurrentSolution: true);

			if (file != null)
			{
				var existed = file.Exists;

				ExportCurrentSettings(file.FullPath.LocalPath);

				if (!existed && file.Exists)
				{
					Dte.ItemOperations.AddExistingItem(file.FullPath.LocalPath);
				}
			}
		}

		private void ImportSelectedSettingsQueryStatus(object sender, EventArgs e)
		{
			Contract.Requires(sender != null);

			var command = (OleMenuCommand)sender;

			command.Enabled = settingsDirectory != null;
		}

		private void ExportCurrentSettingsQueryStatus(object sender, EventArgs e)
		{
			Contract.Requires(sender != null);

			var command = (OleMenuCommand)sender;

			command.Enabled = settingsDirectory != null;
		}

		private void ExportSolutionSettingsQueryStatus(object sender, EventArgs e)
		{
			Contract.Requires(sender != null);

			var command = (OleMenuCommand)sender;

			command.Enabled = settingsDirectory != null && CurrentSolutionSettingsFile != null;
		}

		private void FormatSolutionQueryStatus(object sender, EventArgs e)
		{
			Contract.Requires(sender != null);

			var command = (MenuCommand)sender;

			command.Enabled = Dte.Solution != null && Dte.Solution.IsOpen;
		}

		[ContractVerification(false)]
		private void FormatSolution(object sender, EventArgs e)
		{
			var solution = Dte.Solution;

			if (solution != null
				&& solution.IsOpen
				&& ShouldFormatSolution(solution))
			{
				var projects = solution.Projects;

				Contract.Assume(projects != null);

				foreach (EnvDTE.Project project in projects)
				{
					if (project != null)
					{
						FormatSolutionRecursive(project.ProjectItems);
					}
				}
			}
		}

		private bool ShouldFormatSolution(EnvDTE.Solution solution)
		{
			Contract.Requires(solution != null);

			var sourceControl = Dte.SourceControl;

			int result;

			if (sourceControl != null && sourceControl.IsItemUnderSCC(solution.FullName))
			{
				result = VsShellUtilities.ShowMessageBox(
					this,
					Resources.ConfirmFormatSourceControlledSolution,
					Resources.ConfirmFormatSolutionTitle,
					OLEMSGICON.OLEMSGICON_QUERY,
					OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL,
					OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_SECOND);
			}
			else
			{
				result = VsShellUtilities.ShowMessageBox(
					this,
					Resources.ConfirmFormatSolution,
					Resources.ConfirmFormatSolutionTitle,
					OLEMSGICON.OLEMSGICON_QUERY,
					OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL,
					OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_SECOND);
			}

			ErrorHandler.ThrowOnFailure(result);

			return result == (int)VSConstants.MessageBoxResult.IDOK;
		}

		[ContractVerification(false)]
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "It's safe to bail out here for any catchable exception.")]
		private void FormatSolutionRecursive(EnvDTE.ProjectItems projectItems)
		{
			if (projectItems != null && projectItems.Count > 0)
			{
				try
				{
					foreach (EnvDTE.ProjectItem item in projectItems)
					{
						if (item != null)
						{
							if (item.Kind == VSConstants.ItemTypeGuid.PhysicalFile_string
								&& item.FileCodeModel != null)
							{
								FormatDocument(item);
							}

							var children = item.ProjectItems;

							if (children == null && item.SubProject != null)
							{
								// Required while testing in Visual Studio 2013 against a GIT-sourced solution.
								children = item.SubProject.ProjectItems;
							}

							FormatSolutionRecursive(children);
						}
					}
				}
				catch (Exception ex)
				{
					Debug.WriteLine("Failed to iterate items in project.  Error: " + ex, packageTitle);
				}
			}
		}

		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification = "It's safe to bail out here for any catchable exception.")]
		private void FormatDocument(EnvDTE.ProjectItem item)
		{
			Contract.Requires(item != null);

			try
			{
				var fileName = item.FileNames[0];

				Contract.Assume(Dte.ItemOperations != null);

				var shouldClose = fileName == null || !Dte.ItemOperations.IsFileOpen(fileName, VSConstants.LOGVIEWID.Code_string);

				var window = item.Open(VSConstants.LOGVIEWID.Code_string);

				if (window != null)
				{
					window.Activate();

					Contract.Assume(Dte.Commands != null);

					Dte.Commands.Raise(Guids.vsStd2kCmdIdString, (int)VSConstants.VSStd2KCmdID.FORMATDOCUMENT, null, null);

					if (shouldClose)
					{
						window.Close(EnvDTE.vsSaveChanges.vsSaveChangesYes);
					}
					else
					{
						item.Save();
					}
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine("Failed to format item.  Error: " + ex, packageTitle);
			}
		}

		private void SettingsSelector(object sender, EventArgs e)
		{
			Contract.Requires(settingsDirectory != null);

			var args = e as OleMenuCmdEventArgs;

			if (args != null)
			{
				var input = args.InValue;
				var output = args.OutValue;

				if (input != null)
				{
					lock (gate)
					{
						selectedSettingsFile = settingsFiles[input as string];

						Debug.WriteLine("User selected settings file: " + (selectedSettingsFile == null ? null : selectedSettingsFile.FullPath));

						ImportSelectedSettings();
					}
				}
				else if (output != IntPtr.Zero)
				{
					Debug.WriteLine("Synchronizing combo box control with selected settings file: " + (selectedSettingsFile == null ? null : selectedSettingsFile.FullPath));

					if (selectedSettingsFile == null)
					{
						Marshal.GetNativeVariantForObject(null, output);
					}
					else
					{
						Marshal.GetNativeVariantForObject(selectedSettingsFile.DisplayName, output);
					}
				}
			}

			AssumeReadSettingsDirectory();
		}

		private void SettingsSelectorGetItems(object sender, EventArgs e)
		{
			var args = e as OleMenuCmdEventArgs;

			if (args != null)
			{
				var output = args.OutValue;

				if (output != IntPtr.Zero)
				{
					Marshal.GetNativeVariantForObject(settingsFiles.DisplayNames, output);
				}
			}
		}

		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Contracts", "Invariant-123-93", Justification = "False positive.")]
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Contracts", "Ensures-settingsDirectory == Contract.OldValue(settingsDirectory)", Justification = "False positive.")]
		private void SettingsFileDeletedCreatedChangedOrRenamed(object sender, EventArgs e)
		{
			Contract.Ensures(settingsDirectory == Contract.OldValue(settingsDirectory));

			lock (gate)
			{
				if (unloadingPackage)
				{
					return;
				}

				if (settingsDirectory != null)
				{
					LoadSettingsFiles();
				}
			}

			RefreshCommands();
		}

		[ContractVerification(false)]
		private void SettingFilesWatcherError(object sender, ErrorEventArgs e)
		{
			Debug.Fail("Settings Switcher - File Watcher Error", e.GetException().ToString());
		}

		/// <summary>
		/// According to the docs, the OnBeginShutdown event cannot be canceled.
		/// </summary>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Contracts", "Ensures-settingsDirectory == Contract.OldValue(settingsDirectory)", Justification = "False positive.")]
		private void BeginShutdown()
		{
			Contract.Ensures(settingsDirectory == Contract.OldValue(settingsDirectory));

			lock (gate)
			{
				unsavedSelectedSettingsFile = selectedSettingsFile;
				unloadingPackage = true;

				settingFilesWatcher.EnableRaisingEvents = false;

				if (settingsDirectory != null)
				{
					AssumeReadSettingsDirectory();

					UnloadSolutionSettingsFile();
				}
			}
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				settingFilesWatcher.Dispose();
			}

			base.Dispose(disposing);
		}
	}
}