using System;

namespace DaveSexton.SettingsSwitcher.Vsix
{
	internal static class Guids
	{
		public const string vsixPkgString = "df47a1c4-0804-4f17-9a10-310b0a52892b";
		public const string vsixCmdSetString = "491bca72-7687-41aa-81e1-fc0e3e746393";

		// http://msdn.microsoft.com/en-us/library/microsoft.visualstudio.vsconstants.vsstd2kcmdid.aspx
		public const string vsStd2kCmdIdString = "{1496A755-94DE-11D0-8C3F-00C04FC2AAE2}";

		public static readonly Guid VsixCmdSet = new Guid(vsixCmdSetString);
		public static readonly Guid VsStd2kCmdId = new Guid(vsStd2kCmdIdString);
	}
}