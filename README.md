# FolderViewPainter

![image](https://github.com/user-attachments/assets/ea142d31-5bc0-4ebc-89b7-67ab78c9770f)
![image](https://github.com/user-attachments/assets/c19f2c26-c6c0-4298-ad90-e1314c1855a0)

**Note**: The above menu is just an example. Your Import View menu will be totally custom, containing the views that you've exported with names that you've entered.

## Export and Import Windows Explorer Folder Views

This program adds a right-click context menu to Windows Explorer that allows you to export, import, and reset folder views.

You can export the current view and give that exported view any name you like. You can export as many different views, from as many different folders, as you like.

Any time you want to change the current folder view, to use one of the exported views, it's as easy as right-clicking and selecting one of those views from the Import View menu.

Selecting Reset View will reset the current folder's view to the default for that folder type, but not change any other folder.

## Why would I need or want this program?

### Scenario 1 - When Five Views Are Not Enough

You've set your preferred default view for each folder type (General, Documents, Music, Pictures, Videos) and you don't want to change those defaults. But, sometimes, you want to display a different combination of columns, column arrangement, column widths, sort order, etc. FolderViewPainter will let you save all those selections and adjustments for a view under a single name that you can select with a simple right-click.

### Scenario 2 - When You Need to Go Beyond Generic

You've set up the Explorer folder views for maximum performance by setting all folders to be Generic and only display the standard file system properties of Size, Date modified, and File type (or Extension). But sometimes you want to see certain metadata, such as Dimensions, Bit Depth, Bit Rate, Artist, Album, Length, etc.  FolderViewPainter will let you quickly select from a variety of different saved views and then, just as easily, set the view back to your basic default.

## How to Download and Install

[![image](https://github.com/user-attachments/assets/75e62417-c8ee-43b1-a8a8-a217ce130c91)Download the installer](https://github.com/LesFerch/FolderViewPainter/releases/download/1.3.1/FolderViewPainter-Setup.exe)

[![image](https://github.com/LesFerch/WinSetView/assets/79026235/0188480f-ca53-45d5-b9ff-daafff32869e)Download the zip file](https://github.com/LesFerch/FolderViewPainter/releases/download/1.3.1/FolderViewPainter.zip)

**Note**: Some antivirus software may falsely detect the download as a virus. This can happen any time you download a new executable and may require extra steps to whitelist the file.

**IMPORTANT**: Before starting with this tool, it is highly recommended to first run [WinSetView](https://lesferch.github.io/WinSetView/) to set all of your default preferred folder views. That will clear out all of the saved views in the registry, so that this tool can perform well.

### Install Using Setup Program

1. Download the installer using the link above.
2. Double-click **FolderViewPainter-Setup.exe** to start the installation.
4. In the SmartScreen window, click **More info** and then **Run anyway**.

**Note**: The installer is only provided in English, but the right-click menu items will be created using your current Windows language, if that language is included in the **Language.ini** file.

The right-click menu items will be created for the user that is currently logged on interactively (i.e. desktop is displayed). If you wish to add the right-click menu items to *other* users, log on as each user and either run **FolderViewPainter-Setup.exe** again or navigate to the **FolderViewPainter** folder and double-click **FolderViewPainter.exe** (see **Install and Remove** below for details).

If you don't have other users to set up, skip down to the [**How to Use**](#how-to-use) section.

### Portable Install

1. Download the zip file using the link above.
2. Extract the contents. You should see **FolderViewPainter.exe**, **FolderViewPainter.ini**, and **Language.ini**.
3. Move the contents to a permanent location of your choice. For example **C:\Tools\FolderViewPainter**.
3. Right-click **FolderViewPainter.exe**, select Properties, check **Unblock**, and click **OK**.
5. Double-click **FolderViewPainter.exe** to open the Install/Remove dialog and click **Install** to add the tools to the Explorer right-click menu.
6. If you skipped step 4, then, in the SmartScreen window, click **More info** and then **Run anyway**.
7. Click **OK** when the **Done** message box appears.

When FolderViewPainter is installed as a portable app, you will NOT see the app listed under **Apps** or **Programs and Files**. 

## Install and Remove

![image](https://github.com/user-attachments/assets/431193a2-6ad1-4a00-b312-bf4670489dbc)

The app's install/remove procedure adds, or removes, the commands to/from the context menu. Those commands all use **FolderViewPainter.exe**, so the files must remain in place after doing the **Install**.

The **Remove** option removes the context menu entries. It does not delete the app files.



Upon completion the following dialog pops up. It may be hidden under another window, but can always be found on the taskbar.

![image](https://github.com/user-attachments/assets/aec6802f-9d55-4362-8855-cc83393020db)


**Note**: If you move **FolderViewPainter.exe** after installing, the context menu entries will do nothing because the exe path will be incorrect. To fix that issue, just run the install again.

### About the context menu

The context menu item is created with registry entries only and simply provides submenus entries for each command. When one of those commands are selected, **FolderViewPainter.exe** is run with the appropriate arguments to open the selected option.

This program does NOT create a context menu handler. That is, there is no code that runs when you right-click a folder. Code only runs when you actually select an action. FolderViewPainter adds no overhead to your context menu, other than the insignificant impact of one more context menu item.

## Language Selection

By default, **Install** will create the context menu items in the current system language if that language is found in the **Language.ini** file. Otherwise, it will default to English. To force the context menu items to be created in a specific language, edit the **FolderViewPainter.ini** file and uncomment (remove the semicolon) and change the **Lang=en** entry to the two letter code of the desired language found in the **Language.ini** file. Then, just double-click **FolderViewPainter.exe** and click **Install** to update the context menu entries to the new language.

## How to Use

**IMPORTANT**: Before starting with this tool, it is highly recommended to first run [WinSetView](https://lesferch.github.io/WinSetView/) to set all of your default preferred folder views. That will clear out all of the saved views in the registry, so that this tool can perform well.

Right-click the background of an open folder and you should see the FolderViewPainter context menu:

![image](https://github.com/user-attachments/assets/ea142d31-5bc0-4ebc-89b7-67ab78c9770f)

Select the action you wish to perform. If nothing happens, then the Exe was likely moved after installing. In that case, just double-click the Exe to re-install.

### Import View

This command pops up a menu to let you select a folder view to import. It changes the view of the current folder to the view you select from the pop-up menu. You must export at least one view first.

![image](https://github.com/user-attachments/assets/c19f2c26-c6c0-4298-ad90-e1314c1855a0)

**Note**: The above menu is just an example. Your Import View menu will be totally custom, containing the views that you've exported with names that you've entered.

### Export View

This command exports the current folder's view settings. It prompts for a name (by default it uses the current folder name) and will let you know if you enter a name that's already been used.

When the option **Include Explorer Settings** is checked, the current Explorer settings for the **Preview**/**Details** pane and the settings found under **Properties** > **View** are also exported. These settings correspond to the registry values `Explorer\Modules\GlobalSettings` and `Explorer\Advanced`. Please understand that such settings are not specific to any given folder, so only check this box if you really want to capture the current Explorer view options. 

Since the exported view is saved a standard Reg file, you may use Notepad (or any text editor) to view or edit the exported view.

Note that this dialog displays the number of folder views that are currently in the registry and the number of views you have exported. FolderViewPainter has to enumerate all of the registry's saved views 2-3 times whenever you export or import a view, so it will slow down over time as the number of registry views increases.

It's recommended to run [WinSetView](https://lesferch.github.io/WinSetView/) when the registry view count gets into the high hundreds or thousands. When you click Submit in WinSetView, your default views get set and the saved views are cleared, bringing that count down to a small number.

![image](https://github.com/user-attachments/assets/87176e3a-f6b6-4b86-8e82-26ef7a421060)


### Reset View

This command resets the current folder's view to the current default view for its folder type.

If you have included Explorer settings with the Reset View function (see **Options** below) then those settings will also be applied when you select **Reset View**. The settings (if any) are applied from the Reg file named `!ResetView.ini` in the `SavedViews` folder.

**Note**: Explorer pulls the current default view from three possible sources, in the following order of precedence:

1.  The view set using the `Apply to Folders` button.
2.  The view set in the `FolderTypes` key located in `HKCU` (i.e. as set by WinSetView).
3.  The view set in the `FolderTypes` key located in `HKLM` (i.e. Windows default).

If you set all your folder default view preferences using WinSetView, and haven't used the `Apply to Folders` button since doing that, the views you set in WinSetView are the views you will get when using the Reset View command.

### Manage

This command opens a standard Explorer window to your SavedViews folder where you can rename and/or delete views.

### Options

This command brings up the Options dialog where you can choose to use the default (quick) import and export dialogs or use standard Windows Open and Save dialogs.

![image](https://github.com/user-attachments/assets/f5fbbdc9-ce89-4e61-8ddc-114f025696d3)

When the option **Include additional Explorer Settings with Reset View** is checked, the current Explorer settings for the **Preview**/**Details** pane and the settings found under **Properties** > **View** are saved to a file named `!ResetView.ini`  in the `SavedViews` folder. Unchecking the option deletes that file. Re-checking the option creates the file with the current Explorer view settings.

Since `!ResetView.ini` is a standard Reg file, you may use Notepad (or any text editor) to view or edit it.

### Help

This command opens this Readme file.


## It's Multilingual

FolderViewPainter will detect your Windows language and use it, as long as it has your language in its **Language.ini** file.

For example, here's FolderViewPainter in German (DE):

![image](https://github.com/user-attachments/assets/b991cc27-28df-4cd1-af16-15a6fca232c6)


## Dark Theme Compatible

It automatically detects and switches to a dark theme. Example:

![image](https://github.com/user-attachments/assets/f095cb26-93a0-4857-8bb2-662e8d564c66)


## Notes

The Explorer window is closed and re-opened whenever you export or import a view. This is necessary because Microsoft does not provide an API for working with Explorer views. The app has to force Explorer to update the view in the registry by closing and re-opening the window for that view.

The *GetBagMRU* function, in the C# source code, is derived from a PowerShell script by Keith Miller posted [here](https://stackoverflow.com/a/61240563/15764378).

\
\
[![image](https://github.com/LesFerch/WinSetView/assets/79026235/63b7acbc-36ef-4578-b96a-d0b7ea0cba3a)](https://github.com/LesFerch/FolderViewPainter)
