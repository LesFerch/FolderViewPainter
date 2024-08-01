# FolderViewPainter

![image](https://github.com/user-attachments/assets/7d79bf43-cc3c-45af-ad50-32b1a1a041d2)

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

![image](https://github.com/user-attachments/assets/ff7951ff-b13b-4a46-a727-9fdeeea565b4)

![image](https://github.com/user-attachments/assets/2c550f0d-730f-47ca-9f00-246b84cc11cc)

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

![image](https://github.com/user-attachments/assets/9c3c83d0-cdb5-4040-8f33-8709a462f7a3)

Select the action you wish to perform. If nothing happens, then the Exe was likely moved after installing. In that case, just double-click the Exe to re-install.

### Import View

This command pops up a menu to let you select a folder view to import. It changes the view of the current folder to the view you select from the pop-up menu. You must export at least one view first.

![image](https://github.com/user-attachments/assets/7d79bf43-cc3c-45af-ad50-32b1a1a041d2)

**Note**: The above menu is just an example. Your Import View menu will be totally custom, containing the views that you've exported with names that you've entered.

### Export View

This command exports the current folder's view settings. It prompts for a name (by default it uses the current folder name) and will let you know if you enter a name that's already been used.

Note that this dialog displays the number of folder views that are currently in the registry and the number of views you have exported. FolderViewPainter has to enumerate all of the registry's saved views 2-3 times whenever you export or import a view, so it will slow down over time as the number of registry views increases.

It's recommended to run [WinSetView](https://lesferch.github.io/WinSetView/) when the registry view count gets into the high hundreds or thousands. When you click Submit in WinSetView, your default views get set and the saved views are cleared, bringing that count down to a small number.

![image](https://github.com/user-attachments/assets/b9e9cc1f-e021-4362-8fff-5813043f8b10)

### Reset View

This command resets the current folder's view to the current default view for its folder type.

Explorer pulls the current default view from three possible sources, in the following order of precedence:

1.  The view set using the `Apply to Folders` button.
2.  The view set in the `FolderTypes` key located in `HKCU` (i.e. as set by WinSetView).
3.  The view set in the `FolderTypes` key located in `HKLM` (i.e. Windows default).

If you set all your folder default view preferences using WinSetView, and haven't touched the `Apply to Folders` button since doing that, the views you set in WinSetView are the views you will get when using the Reset Views command.

### Manage

This command opens a standard Explorer window to your SavedViews folder where you can rename and/or delete views.

### Options

This command brings up the Options dialog where you can choose to use the default (quick) import and export dialogs or use standard Windows Open and Save dialogs.

![image](https://github.com/user-attachments/assets/bc663cc8-39e3-44a4-94de-786f8f6f83a4)

### Help

This command opens this Readme file.


## It's Multilingual

FolderViewPainter will detect your Windows language and use it, as long as it has your language in its **Language.ini** file.

Here are some screenshots of FolderViewPainter for the German (DE) language:

![image](https://github.com/user-attachments/assets/fe6dcea1-2f67-4ac3-b381-6e130f6b3123)

![image](https://github.com/user-attachments/assets/db6f9cac-8ade-49be-a2c9-cdfc9d3ec740)

![image](https://github.com/user-attachments/assets/ffff9802-89f4-4006-8ec0-c81cd20c2844)

## Dark Theme Compatible

It automatically detects and switches to a dark theme. Here are some screenshots:

![image](https://github.com/user-attachments/assets/c0b4ef2c-82d0-438f-975d-daeb815c5ab2)

![image](https://github.com/user-attachments/assets/0aee026a-4a6c-4d41-8ad7-0b66222b3926)

![image](https://github.com/user-attachments/assets/2096cd09-8c46-4676-b046-6d3bcd9189ff)

![image](https://github.com/user-attachments/assets/b6b42b78-96e0-4389-820e-6bbac4ea2161)

## Notes

The Explorer window is closed and re-opened whenever you export or import a view. This is necessary because Microsoft does not provide an API for working with Explorer views. The app has to force Explorer to update the view in the registry by closing and re-opening the window for that view.

The *GetBagMRU* function, in the C# source code, is derived from a PowerShell script by Keith Miller posted [here](https://stackoverflow.com/a/61240563/15764378).

\
\
[![image](https://github.com/LesFerch/WinSetView/assets/79026235/63b7acbc-36ef-4578-b96a-d0b7ea0cba3a)](https://github.com/LesFerch/FolderViewPainter)
