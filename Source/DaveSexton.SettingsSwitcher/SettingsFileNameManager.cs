using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using DaveSexton.SettingsSwitcher.Properties;

namespace DaveSexton.SettingsSwitcher
{
	/// <summary>
	/// Generates unique display names for settings files.
	/// </summary>
	public sealed class SettingsFileNameManager : IEnumerable<SettingsFile>
	{
#if DEBUG
		public bool IsSealed
		{
			get
			{
				Contract.Ensures(Contract.Result<bool>() == isSealed);

				return isSealed;
			}
		}
#endif

		private static StringDictionary PersistedLastAccessTimes
		{
			[MethodImpl(MethodImplOptions.Synchronized)]
			get
			{
				Contract.Ensures(Contract.Result<StringDictionary>() != null);

				var lastAccessTimes = Settings.Default.SettingsFilesLastAccessTimes;

				if (lastAccessTimes == null)
				{
					Settings.Default.SettingsFilesLastAccessTimes = lastAccessTimes = new StringDictionary();
				}

				return lastAccessTimes;
			}
		}

		private readonly List<SettingsFile> files = new List<SettingsFile>();
		private readonly Uri autoSettingsFile, oldAutoSettingsFile, originalSettingsFile;
#if DEBUG
		private bool isSealed;
#endif

		/// <summary>
		/// Constructs a new instance of the <see cref="SettingsFileNameManager"/> class.
		/// </summary>
		/// <param name="autoSettingsFile">The path to the auto-save settings file.  The file does not have to exist.</param>
		/// <param name="oldAutoSettingsFile">The path to the old auto-save settings file.  The file does not have to exist.</param>
		/// <param name="originalSettingsFile">The path to the original auto-saved settings file that was created when the extension was initialized for the first time.  The file does not have to exist.</param>
		public SettingsFileNameManager(Uri autoSettingsFile, Uri oldAutoSettingsFile, Uri originalSettingsFile)
		{
			Contract.Requires(autoSettingsFile != null);
			Contract.Requires(autoSettingsFile.IsAbsoluteUri);
			Contract.Requires(autoSettingsFile.IsFile);
			Contract.Requires(oldAutoSettingsFile != null);
			Contract.Requires(oldAutoSettingsFile.IsAbsoluteUri);
			Contract.Requires(oldAutoSettingsFile.IsFile);
			Contract.Requires(originalSettingsFile != null);
			Contract.Requires(originalSettingsFile.IsAbsoluteUri);
			Contract.Requires(originalSettingsFile.IsFile);

			this.autoSettingsFile = autoSettingsFile;
			this.oldAutoSettingsFile = oldAutoSettingsFile;
			this.originalSettingsFile = originalSettingsFile;
		}

		[ContractInvariantMethod]
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Required for code contracts.")]
		private void ObjectInvariant()
		{
			Contract.Invariant(files != null);
			Contract.Invariant(autoSettingsFile != null);
			Contract.Invariant(autoSettingsFile.IsAbsoluteUri);
			Contract.Invariant(autoSettingsFile.IsFile);
			Contract.Invariant(oldAutoSettingsFile != null);
			Contract.Invariant(oldAutoSettingsFile.IsAbsoluteUri);
			Contract.Invariant(oldAutoSettingsFile.IsFile);
			Contract.Invariant(originalSettingsFile != null);
			Contract.Invariant(originalSettingsFile.IsAbsoluteUri);
			Contract.Invariant(originalSettingsFile.IsFile);
		}

		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Contracts", "Requires-14-129", Justification = "The long value is always written by DateTime itself, so it must be within range, assuming that the XML settings file wasn't changed manually.")]
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Contracts", "Requires-40-129", Justification = "The long value is always written by DateTime itself, so it must be within range, assuming that the XML settings file wasn't changed manually.")]
		internal static DateTime GetLastAccessTime(Uri file)
		{
			Contract.Requires(file != null);
			Contract.Requires(file.IsAbsoluteUri);
			Contract.Requires(file.IsFile);
			Contract.Ensures(file.IsAbsoluteUri);
			Contract.Ensures(file.IsFile);

			var times = PersistedLastAccessTimes;

			Contract.Assume(!string.IsNullOrEmpty(file.LocalPath));

			return times.ContainsKey(file.LocalPath)
					 ? DateTime.FromFileTimeUtc(long.Parse(times[file.LocalPath], CultureInfo.InvariantCulture))
					 : File.GetLastAccessTimeUtc(file.LocalPath);		// In practice, Windows may not update file access times with alacrity, or at all, so it's only used as a fallback.
		}

		internal static void SetLastAccessTime(Uri file)
		{
			Contract.Requires(file != null);
			Contract.Requires(file.IsAbsoluteUri);
			Contract.Requires(file.IsFile);
			Contract.Ensures(file.IsAbsoluteUri);
			Contract.Ensures(file.IsFile);

			PersistedLastAccessTimes[file.LocalPath] = DateTime.Now.ToFileTimeUtc().ToString(CultureInfo.InvariantCulture);

			Debug.WriteLine("Saving last access times.", "SettingsFileNameManager");

			Settings.Default.Save();
		}

		internal void AddFile(SettingsFile file)
		{
#if DEBUG
			Contract.Requires(file is SolutionSettingsFile || !IsSealed);
			Contract.Requires(!(file is SolutionSettingsFile) || !Contract.Exists(this, f => f is SolutionSettingsFile));
#endif
			Contract.Requires(file != null);

			files.Add(file);
		}

		/// <summary>
		/// Removes the specified <paramref name="file"/> from the list.
		/// </summary>
		/// <param name="file">The file to be removed.</param>
		/// <returns><see langword="True"/> if the file was found in the list; otherwise, <see langword="false"/>.</returns>
		public bool RemoveFile(SettingsFile file)
		{
#if DEBUG
			Contract.Requires(file is SolutionSettingsFile || !IsSealed);
#endif
			Contract.Requires(file != null);

			return files.Remove(file);
		}

		/// <summary>
		/// Clears the list of files.
		/// </summary>
		public void Clear()
		{
#if DEBUG
			isSealed = false;
#endif

			files.Clear();
		}

		internal string GetDisplayName(SettingsFile file)
		{
			Contract.Requires(file != null);
			Contract.Ensures(Contract.Result<string>() != null);

#if DEBUG
			isSealed = true;
#endif

			/* Display names must be unique.  It seems that Visual Studio can only display strings in 
			 * combo commands; therefore, the display name is used to find the selected object.
			 */

			string format;

			if (file.FullPath == autoSettingsFile || file.FullPath == oldAutoSettingsFile || file.FullPath == originalSettingsFile)
			{
				format = Resources.SettingsFileAutoSavedDisplayNameFormat;
			}
			else if (HasFileWithSameName(file))
			{
				format = Resources.SettingsFileRelativeDisplayNameFormat;
			}
			else if (file.SpecialDisplayName != null)
			{
				format = Resources.SettingsFileSpecialDisplayNameFormat;
			}
			else
			{
				format = Resources.SettingsFileDefaultDisplayNameFormat;
			}

			return FormatName(file, format);
		}

		private static string FormatName(SettingsFile file, string format)
		{
			Contract.Requires(file != null);
			Contract.Requires(format != null);
			Contract.Ensures(Contract.Result<string>() != null);

			return string.Format(CultureInfo.CurrentCulture, format, file.Name, Path.GetDirectoryName(file.RelativePath.ToString()), file.SpecialDisplayName);
		}

		private bool HasFileWithSameName(SettingsFile file)
		{
			Contract.Requires(file != null);

			return (from other in files
							where other != file
								 && string.Equals(other.Name, file.Name, StringComparison.OrdinalIgnoreCase)
							select other)
							.Any();
		}

		/// <inheritdoc/>
		public IEnumerator<SettingsFile> GetEnumerator()
		{
			return files.GetEnumerator();
		}

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return GetEnumerator();
		}
	}
}