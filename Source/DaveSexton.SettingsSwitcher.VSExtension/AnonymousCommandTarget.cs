using System;
using System.Diagnostics.Contracts;
using Microsoft.VisualStudio.OLE.Interop;

namespace DaveSexton.SettingsSwitcher.Vsix
{
	internal delegate int OleCommandQueryStatus(Guid groupId, OLECMD[] commands, IntPtr text);
	internal delegate int OleCommandExecute(Guid groupId, uint commandId, uint options, IntPtr input, IntPtr output);

	internal sealed class AnonymousCommandTarget : IOleCommandTarget
	{
		private readonly OleCommandQueryStatus queryStatus;
		private readonly OleCommandExecute execute;

		public AnonymousCommandTarget(OleCommandQueryStatus queryStatus, OleCommandExecute execute)
		{
			Contract.Requires(queryStatus != null);
			Contract.Requires(execute != null);

			this.queryStatus = queryStatus;
			this.execute = execute;
		}

		[ContractInvariantMethod]
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Required for code contracts.")]
		private void ObjectInvariant()
		{
			Contract.Invariant(queryStatus != null);
			Contract.Invariant(execute != null);
		}

		[System.Diagnostics.CodeAnalysis.SuppressMessage("StyleCop.CSharp.NamingRules", "SA1305:FieldNamesMustNotUseHungarianNotation", Justification = "Interface implementation.")]
		public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
		{
			return queryStatus(pguidCmdGroup, prgCmds, pCmdText);
		}

		[System.Diagnostics.CodeAnalysis.SuppressMessage("StyleCop.CSharp.NamingRules", "SA1305:FieldNamesMustNotUseHungarianNotation", Justification = "Interface implementation.")]
		public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
		{
			return execute(pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
		}
	}
}