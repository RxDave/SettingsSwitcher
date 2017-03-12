using System;
using System.Diagnostics.Contracts;
using System.IO;

namespace DaveSexton.SettingsSwitcher
{
	/// <summary>
	/// Represents a Visual Studio settings file on disc.
	/// </summary>
	public class SettingsFile
	{
		/// <summary>
		/// Gets the full path to this settings file.
		/// </summary>
		public Uri FullPath
		{
			get
			{
				Contract.Ensures(Contract.Result<Uri>() != null);
				Contract.Ensures(Contract.Result<Uri>().IsAbsoluteUri);
				Contract.Ensures(Contract.Result<Uri>().IsFile);

				return fullPath;
			}
		}

		/// <summary>
		/// Gets the path to this settings file relative from the directory in which Visual Studio exports settings files.
		/// </summary>
		public Uri RelativePath
		{
			get
			{
				Contract.Ensures(Contract.Result<Uri>() != null);
				Contract.Ensures(!Contract.Result<Uri>().IsAbsoluteUri);

				return relativePath;
			}
		}

		/// <summary>
		/// Gets the name of this settings file without the extension.
		/// </summary>
		public string Name
		{
			get
			{
				Contract.Ensures(Contract.Result<string>() != null);

				return name ?? (name = Path.GetFileNameWithoutExtension(fullPath.LocalPath));
			}
		}

		/// <summary>
		/// Gets the unique display name of this settings file.
		/// </summary>
		public string DisplayName
		{
			get
			{
				Contract.Ensures(Contract.Result<string>() != null);

				return displayName ?? (displayName = manager.GetDisplayName(this));
			}
		}

		/// <summary>
		/// Gets a special display name to be shown with the actual name, or <see langword="null"/>.
		/// </summary>
		public virtual string SpecialDisplayName
		{
			get
			{
				return null;
			}
		}

		/// <summary>
		/// Gets a value indicating whether the file exists.
		/// </summary>
		public bool Exists
		{
			get
			{
				return File.Exists(fullPath.LocalPath);
			}
		}

		/// <summary>
		/// Gets the date and time of the last read.
		/// </summary>
		public DateTime LastAccessed
		{
			get
			{
				return SettingsFileNameManager.GetLastAccessTime(fullPath);
			}
		}

		private readonly Uri fullPath, relativePath;
		private readonly SettingsFileNameManager manager;
		private string name, displayName;

		/// <summary>
		/// Constructs a new instance of the <see cref="SettingsFile"/> class.
		/// </summary>
		/// <param name="fullPath">The full path to this settings file.</param>
		/// <param name="basePath">The full path to the directory in which Visual Studio exports settings files.</param>
		/// <param name="manager">An object that generates unique names for settings files.</param>
		public SettingsFile(Uri fullPath, Uri basePath, SettingsFileNameManager manager)
		{
			Contract.Requires(fullPath != null);
			Contract.Requires(fullPath.IsAbsoluteUri);
			Contract.Requires(fullPath.IsFile);
			Contract.Requires(basePath != null);
			Contract.Requires(basePath.IsAbsoluteUri);
			Contract.Requires(basePath.IsFile);
			Contract.Requires(manager != null);

			this.fullPath = fullPath;
			this.relativePath = basePath.MakeRelativeUri(fullPath);
			this.manager = manager;

			Contract.Assume(!this.relativePath.IsAbsoluteUri);

#if DEBUG
			Contract.Assume(this is SolutionSettingsFile || !manager.IsSealed);
#endif

			manager.AddFile(this);
		}

		[ContractInvariantMethod]
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Required for code contracts.")]
		private void ObjectInvariant()
		{
			Contract.Invariant(fullPath != null);
			Contract.Invariant(fullPath.IsAbsoluteUri);
			Contract.Invariant(fullPath.IsFile);
			Contract.Invariant(relativePath != null);
			Contract.Invariant(!relativePath.IsAbsoluteUri);
			Contract.Invariant(manager != null);
		}

		/// <summary>
		/// Updates the value of <see cref="LastAccessed"/>.
		/// </summary>
		public void Accessed()
		{
			SettingsFileNameManager.SetLastAccessTime(fullPath);
		}

		/// <summary>
		/// Returns the <see cref="DisplayName"/> for this settings file.
		/// </summary>
		/// <returns>The value of the <see cref="DisplayName"/> property.</returns>
		public override string ToString()
		{
			return DisplayName;
		}
	}
}