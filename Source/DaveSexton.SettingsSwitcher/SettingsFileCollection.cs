using System;
using System.Collections.ObjectModel;
using System.Diagnostics.Contracts;
using System.Linq;

namespace DaveSexton.SettingsSwitcher
{
	/// <summary>
	/// A list of <see cref="SettingsFile"/> objects.
	/// </summary>
	public sealed class SettingsFileCollection : Collection<SettingsFile>
	{
		/// <summary>
		/// Gets the settings file with the specified unique <paramref name="displayName"/>.
		/// </summary>
		/// <param name="displayName">The unique display name of the settings file to be retrieved.</param>
		/// <returns>The settings file with the specified <paramref name="displayName"/>; otherwise, <see langword="null"/> if a match is not found.</returns>
		public SettingsFile this[string displayName]
		{
			get
			{
				if (string.IsNullOrEmpty(displayName))
				{
					return null;
				}
				else
				{
					return this.FirstOrDefault(file => string.Equals(file.DisplayName, displayName, StringComparison.Ordinal));
				}
			}
		}

		/// <summary>
		/// Gets an array of unique display names representing all of the settings files within this collection.
		/// </summary>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1819:PropertiesShouldNotReturnArrays", Justification = "Visual Studio combo commands do not support generic types as the items source. An array was the quickest alternative; it worked on the first try.")]
		public string[] DisplayNames
		{
			get
			{
				Contract.Ensures(Contract.Result<string[]>() != null);

				if (displayNames == null)
				{
					displayNames = (from file in this
													select file.DisplayName)
													.ToArray();
				}

				return displayNames;
			}
		}

		private string[] displayNames;

		/// <summary>
		/// Gets the settings file with the specified full path and name.
		/// </summary>
		/// <param name="file">The full path and name of the settings file to be retrieved.</param>
		/// <returns>The settings file with the specified full path and name; otherwise, <see langword="null"/> if no match is found.</returns>
		public SettingsFile FindByFullName(string file)
		{
			if (string.IsNullOrEmpty(file))
			{
				return null;
			}
			else
			{
				return this.FirstOrDefault(f => string.Equals(f.FullPath.LocalPath, file, StringComparison.OrdinalIgnoreCase));
			}
		}

		/// <summary>
		/// Removes all settings files from the collection.
		/// </summary>
		protected override void ClearItems()
		{
			displayNames = null;

			base.ClearItems();
		}

		/// <summary>
		/// Inserts a settings file into the collection.
		/// </summary>
		/// <param name="index">The index at which to insert the settings file.</param>
		/// <param name="item">The settings file to be inserted.</param>
		protected override void InsertItem(int index, SettingsFile item)
		{
			displayNames = null;

			base.InsertItem(index, item);
		}

		/// <summary>
		/// Removes the specified settings file from the collection.
		/// </summary>
		/// <param name="index">The settings file to be removed.</param>
		protected override void RemoveItem(int index)
		{
			displayNames = null;

			base.RemoveItem(index);
		}

		/// <summary>
		/// Replaces the settings file at the specified index with the specified settings file.
		/// </summary>
		/// <param name="index">The index at which to set the new settings file.</param>
		/// <param name="item">The new settings file.</param>
		protected override void SetItem(int index, SettingsFile item)
		{
			displayNames = null;

			base.SetItem(index, item);
		}
	}
}