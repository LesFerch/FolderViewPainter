# FolderViewPainter

![image](https://github.com/LesFerch/FolderViewPainter/assets/79026235/3b41fba6-e7a6-427c-8ce7-0f48e56850f2)

**Note**: The above menu is just an example. Your Import View menu will be totally custom, containing the views that you've exported with names that you've entered.

## Export and Import Windows Explorer Folder Views

This program adds a right-click context menu to Windows Explorer that allows you to export and import folder views.

You can export the current view and give that exported view any name you like. You can export as many different views, from as many different folders, as you like.

Any time you want to change the current folder view, to use one of the exported views, it's as easy as right-clicking and selecting one of those views from the Import View menu.

## Why would I need or want this program?

### Scenario 1 - When Five Views Are Not Enough

Let's say that you've set your preferred default view for each folder type (General, Documents, Music, Pictures, Videos) and you don't want to change those defaults. But, sometimes, you want to display a different combination of columns, column arrangement, column widths, sort order, etc. Wouldn't it be great if you could save all those selections and adjustments for a view under a single name that you can select with a simple right-click?

### Scenario 2 - When You Need to Go Beyond Generic

Lets say you've set up the Explorer folder views for maximum performance by setting all folders to be Generic and only display the standard file system properties of Size, Date modified, and File type (or Extension). But sometimes you want to see certain metadata, such as Dimensions, Bit Depth, Bit Rate, Artist, Album, Length, etc. Wouldn't it be great if you could quickly select from a variety of different saved views and then, just as easily, set the view back to your basic default?

## How to Download and Install

[![image](https://github.com/LesFerch/WinSetView/assets/79026235/0188480f-ca53-45d5-b9ff-daafff32869e)Download the zip file](https://github.com/LesFerch/FolderViewPainter/releases/download/1.1.3/FolderViewPainter.zip)

**Note**: Some antivirus software may falsely detect the download as a virus. This can happen any time you download a new executable and may require extra steps to whitelist the file.

**IMPORTANT**: Before starting with this tool, it is highly recommended to first run [WinSetView](https://lesferch.github.io/WinSetView/) to set all of your default preferred folder views. That will clear out all of the saved views in the registry, so that this tool can perform well.

1. Download the zip file using the link above.
2. Extract **FolderViewPainter.exe** and, optionally, **Language.ini**.
3. Right-click **FolderViewPainter.exe**, select Properties, check **Unblock**, and click **OK**.
4. Move **FolderViewPainter.exe** to the folder of your choice (you need to keep this file). Keep the **Language.ini** file with the Exe if your language is not English.
5. Double-click **FolderViewPainter.exe** to open the Install/Remove dialog and click **Install** to add the tool to the Explorer right-click menu.
6. If you skipped step 3, then, in the SmartScreen window, click **More info** and then **Run anyway**.
7. Click **OK** when the **Done** message box appears.

**Note**: Some antivirus software may falsely detect the download as a virus. This can happen any time you download a new executable and may require extra steps to whitelist the file.

![image](https://github.com/LesFerch/FolderViewPainter/assets/79026235/13e7486e-ed78-4a6c-acab-7451d402ce51)

![image](https://github.com/LesFerch/FolderViewPainter/assets/79026235/d14ec83a-c608-4650-bf73-10b22979c426)

**Note**: Sometimes the **Done** message pops up behind another open window. If that happens, you should see the FolderViewPainter icon on the taskbar, where you can click to bring that dialog to the front. Alternatively, you can minimize the window(s) that are on top of the dialog.

**Note**: If you move **FolderViewPainter.exe** after installing, the context menu entries will do nothing because the exe path will be incorrect. To fix that issue, just run the install again (as descibed in Step 5 above).

### Command line Installation and Removal

If you wish to install or remove the tool via the command line, use these commands:

`FolderViewPainter /install`

`FolderViewPainter /remove`

### About the context menu

The context menu item is created with registry entries only and simply provides submenus entries for each command. When one of those commands are selected, FolderViewPainter.exe is run with the appropriate arguments to either bring up the import view list, prompt to export a view, manage the saved views, change options, or get help.

This program does NOT create a context menu handler. That is, there is no code that runs when you right-click a folder. Code only runs when you actually select an action. FolderViewPainter will add no overhead to your context menu, other than the insignificant impact of one more context menu item.

## How to Use

**IMPORTANT**: Before starting with this tool, it is highly recommended to first run [WinSetView](https://lesferch.github.io/WinSetView/) to set all of your default preferred folder views. That will clear out all of the saved views in the registry, so that this tool can perform well.

Right-click the background of an open folder and you should see the FolderViewPainter context menu:

![image](https://github.com/LesFerch/FolderViewPainter/assets/79026235/312ed72d-5b50-4d28-bcd4-8f97c8fd098c)

Select the action you wish to perform. If nothing happens, then the Exe was likely moved after installing. In that case, just double-click the Exe to re-install.

### Import View

This command pops up a menu to let you select a folder view to import. It changes the view of the current folder to the view you select from the pop-up menu. You must export at least one view first.

![image](https://github.com/LesFerch/FolderViewPainter/assets/79026235/3b41fba6-e7a6-427c-8ce7-0f48e56850f2)

**Note**: The above menu is just an example. Your Import View menu will be totally custom, containing the views that you've exported with names that you've entered.

### Export View

This command exports the current folder's view settings. It prompts for a name (by default it uses the current folder name) and will let you know if you enter a name that's already been used.

Note that this dialog displays the number of folder views that are currently in the registry and the number of views you have exported. FolderViewPainter has to enumerate all of the registry's saved views 2-3 times whenever you export or import a view, so it will slow down over time as the number of registry views increases.

It's recommended to run [WinSetView](https://lesferch.github.io/WinSetView/) when the registry view count gets into the high hundreds or thousands. When you click Submit in WinSetView, your default views get set and the saved views are cleared, bringing that count down to a small number.

![image](https://github.com/LesFerch/FolderViewPainter/assets/79026235/bbbf9dae-4ed8-43a6-a30b-b49e5895f708)

### Manage

This command opens a standard Explorer window to your SavedViews folder where you can rename and/or delete views.

### Options

This command brings up the Options dialog where you can choose to use the default (quick) import and export dialogs or use standard Windows Open and Save dialogs.

![image](https://github.com/LesFerch/FolderViewPainter/assets/79026235/aac5ae63-240c-4951-810e-ec0a703cf9a2)

### Help

This command opens this Readme file.


## It's Multilingual

FolderViewPainter will detect your Windows language and use it, as long as it has your language in its **Language.ini** file.

Here are some screenshots of FolderViewPainter for the German (DE) language:

![image](https://github.com/LesFerch/FolderViewPainter/assets/79026235/038bc1a7-e605-4091-8019-fc9d3e225a36)

![image](https://github.com/LesFerch/FolderViewPainter/assets/79026235/8771bd2c-4854-45e7-bd29-eab82f3e9ce1)

![image](https://github.com/LesFerch/FolderViewPainter/assets/79026235/8d46dc15-9f74-4ebf-a694-788c03003f58)

## Dark Theme Compatible

It automatically detects and switches to a dark theme. Here are some screenshots:

![image](https://github.com/LesFerch/FolderViewPainter/assets/79026235/77415654-0c17-4099-9a73-894bdd21600d)

![image](https://github.com/LesFerch/FolderViewPainter/assets/79026235/8e7b54a6-f76d-45df-a334-0790eccf0237)

![image](https://github.com/LesFerch/FolderViewPainter/assets/79026235/e0599fdd-316b-47dc-8664-765cbc8b9e8a)

![image](https://github.com/LesFerch/FolderViewPainter/assets/79026235/dfb110e5-1c6d-4b14-a390-64633009f0c6)

## Notes

The Explorer window is closed and re-opened whenever you export or import a view. This is necessary because Microsoft does not provide an API for working with Explorer views. The app has to force Explorer to update the view in the registry by closing and re-opening the window for that view.

The *GetBagMRU* function, in the C# source code, is derived from a PowerShell script by Keith Miller posted [here](https://stackoverflow.com/a/61240563/15764378).

\
\
[![image](https://github.com/LesFerch/WinSetView/assets/79026235/63b7acbc-36ef-4578-b96a-d0b7ea0cba3a)](https://github.com/LesFerch/FolderViewPainter)
