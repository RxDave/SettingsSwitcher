using System;
using System.Reflection;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("DaveSexton.SettingsSwitcher")]
[assembly: AssemblyDescription("")]
[assembly: CLSCompliant(false)]
[assembly: ComVisible(false)]

internal static class AssemblyConstants
{
	public const string Version = GlobalConstants.Version;

	/// <summary>
	/// Semantic version for the assembly, indicating a prerelease package in NuGet.
	/// </summary>
	/// <remarks>
	/// <para>
	/// The specified name can be arbitrary, but its mere presence indicates a prerelease package.
	/// To indicate a release package instead, use an empty string.
	/// </para>
	/// <para>
	/// If specified, the value must include a preceding hyphen; e.g., "-alpha", "-beta", "-rc".
	/// </para>
	/// </remarks>
	/// <seealso href="http://docs.nuget.org/docs/reference/versioning#Really_brief_introduction_to_SemVer">
	/// NuGet Semantic Versioning
	/// </seealso>
	public const string PrereleaseVersion = "";
}