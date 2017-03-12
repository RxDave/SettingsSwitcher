![Toolbar](../../wiki/Images/ToolBar.png)
# Usage
Easily switch to different Visual Studio settings when...
* your laptop is docked, or when it's undocked and you're on the go.
* giving a live presentation.
* recording a how-to video.
* working on open source software with different code formatting requirements.
* an update to Visual Studio or a bad VS extension resets your desired settings.

# Features
* _Solution Settings Files_: Whenever a solution is opened, if a settings file of the same name (Solution_Name.vssettings) exists next to the solution file (Solution_Name.sln), its settings are automatically applied.
* Associates the current settings with the active solution. The associated settings will be applied automatically whenever the solution is opened. (Overrides the active _Solution Settings File_ without modifying it.)
* Provides a toolbar (shown above).
  * Selecting an item in the drop-down list immediately imports those settings into Visual Studio.
  * ![Export Current Settings](../../wiki/Images/Export Current Settings Button.png) _Export Current Settings_ saves the current settings to a file.
  * ![Export Solution Settings](../../wiki/Images/Export Solution Settings Button.png) _Export Solution Settings_ saves the current settings to the active solution's settings file, or creates a new _Solution Settings File_ and adds it to the solution.
  * ![Format Solution ](../../wiki/Images/Format Solution Button.png) _Format Solution_ formats all documents in the solution according to the current settings.

See the [Wiki](../../wiki) for details.

# Thanks
This extension is based on a tip from [Sara Ford's blog](http://blogs.msdn.com/b/saraford/archive/2008/12/05/did-you-know-you-can-create-toolbar-buttons-to-quickly-toggle-your-favorite-vs-settings-371.aspx).
