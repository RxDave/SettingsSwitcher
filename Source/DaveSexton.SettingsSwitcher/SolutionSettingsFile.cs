using System;
using System.Diagnostics.Contracts;
using System.IO;
using DaveSexton.SettingsSwitcher.Properties;

namespace DaveSexton.SettingsSwitcher
{
	/// <summary>
	/// Represents a Visual Studio settings file associated with a particular solution file (.sln).
	/// </summary>
	public sealed class SolutionSettingsFile : SettingsFile
	{
		/// <summary>
		/// Gets the full path and name of the solution file.
		/// </summary>
		public string SolutionFileName
		{
			get
			{
				Contract.Ensures(!string.IsNullOrWhiteSpace(Contract.Result<string>()));

				var settingsFile = FullPath.LocalPath;

				var path = Path.Combine(Path.GetDirectoryName(settingsFile), Path.GetFileNameWithoutExtension(settingsFile) + solutionFileExtension);

				Contract.Assume(!string.IsNullOrWhiteSpace(path));

				return path;
			}
		}

		/// <inheritdoc/>
		public override string SpecialDisplayName
		{
			get
			{
				return Resources.SettingsFileSolutionDisplayName;
			}
		}

		private const string solutionFileExtension = ".sln";

		/// <summary>
		/// Constructs a new instance of the <see cref="SolutionSettingsFile"/> class.
		/// </summary>
		/// <param name="fullPath">The full path to this solution settings file.</param>
		/// <param name="manager">An object that generates unique names for settings files.</param>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Contracts", "Requires-24-9", Justification = "False positive.")]
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Contracts", "Requires-35-9", Justification = "False positive.")]
		public SolutionSettingsFile(Uri fullPath, SettingsFileNameManager manager)
			: base(fullPath, GetBasePath(fullPath), manager)
		{
			Contract.Requires(fullPath != null);
			Contract.Requires(fullPath.IsAbsoluteUri);
			Contract.Requires(fullPath.IsFile);
			Contract.Requires(manager != null);
		}

		private static Uri GetBasePath(Uri fullPath)
		{
			Contract.Requires(fullPath != null);
			Contract.Requires(fullPath.IsAbsoluteUri);
			Contract.Requires(fullPath.IsFile);
			Contract.Ensures(Contract.Result<Uri>() != null);
			Contract.Ensures(Contract.Result<Uri>().IsAbsoluteUri);
			Contract.Ensures(Contract.Result<Uri>().IsFile);

			var path = new Uri(Path.GetDirectoryName(fullPath.LocalPath) + '\\', UriKind.Absolute);

			Contract.Assume(path.IsAbsoluteUri);
			Contract.Assume(path.IsFile);

			return path;
		}

		/// <summary>
		/// Gets the full path and name of a solution's settings file.
		/// </summary>
		/// <param name="solutionFile">The solution for which the settings file will be returned.</param>
		/// <param name="settingsFileExtension">The file extension that replaces the .sln extension of the solution file.</param>
		/// <returns>Full path and name of the solution's settings file.</returns>
		public static string GetFullName(string solutionFile, string settingsFileExtension)
		{
			Contract.Requires(!string.IsNullOrWhiteSpace(solutionFile));
			Contract.Requires(!string.IsNullOrWhiteSpace(settingsFileExtension));
			Contract.Ensures(!string.IsNullOrWhiteSpace(Contract.Result<string>()));

			var path = Path.Combine(Path.GetDirectoryName(solutionFile), Path.GetFileNameWithoutExtension(solutionFile) + settingsFileExtension);

			Contract.Assume(!string.IsNullOrWhiteSpace(path));

			return path;
		}
	}
}