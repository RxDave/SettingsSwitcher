using System;
using System.Reflection;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("Visual Studio Settings Switcher")]
[assembly: AssemblyDescription("Provides commands and packaging capabilities for quickly switching settings in Visual Studio.")]
[assembly: CLSCompliant(true)]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("a46da65c-1a30-4c3d-ac6f-9d3f8212f08b")]

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