# Excel File Cataloger
This Excel workbook is a Document Management System. When run, it parses the root directory specified in Settings for all files and folders, then it builds an HTML webpage to display the data in a more user friendly searchable format than the Windows file explorer.


# Getting Started
[Download the workbook](https://github.com/M-Scott-Lassiter/Excel-File-Cataloger/releases/download/v1.1.0/Excel.File.Cataloger.v1.1.0.xlsm) to your computer.

After opening the workbook, press the `Instructions` button to review a self-contained version of these instructions.

## Configuring
Press the `Settings` button to configure the utility.
  - **Webpage Title**: Specifies both the `<title>` and `<h1>` elements in the page.
  - **Root Directory**: The system directory that the utility reads from. Starting this value with ".\" tells the utility to use a relative file path to the workbook's current location. Example: `.\Sample Data Set`.
  - **Webpage Output Location**: Folder the output webpage will save to. The output will have a filename of `[WebpageTitle].html` with capitalization maintained but any spaces removed. This can also use the relative location operator.
  - **Administrator Email**: An email address used for an email quick link at the bottom of the output page. Can be set to any email address. If left blank, this utility omits this page element. There is no error checking for invalid email addresses on this setting.
  - **Omitted File/Folder Prefix**: The utility omits any file or folder beginning with this text prefix. This is useful for withholding files or folders from the utility in a working directory. For example, you can have file folders start with an underscore and set this value to `_`. Any files or folders (and all files within these folders) will not appear on the webpage.

## Creating Tags
Tags optional are pre-built search strings applied during the build process. Files can have multiple tags assigned. Selecting a tag from the webpage dropdown menu bar filters the page and leaves only those tagged files remaining. Any subsequent search then applies only within that tagged group. If no tags are defined, then the webpage will omit this feature.

Use the Tags table to organize tags and apply filters. The webpage displays the value from the "Tag Name" column in the drop down menu. The Filters column contains the search strings. When building the webpage, each of these strings are compared against the file name to determine applicability.

Tag names must be unique. Any tags with a blank name are deleted while building the webpage. The webpage will not display tags with no filter specified.

To apply more than one filter term to a tag, separate them with a comma. If the search term contains a comma, you can enclose it within double quotes. Tags are not case sensitive.

Example Filter Strings:
-  *America, United States, Mexico, Canada*
-  *Gunsmoke, "The Good, The Bad, and the Ugly", High Noon*

## Building the Webpage
Press the `Build Webpage` button. A prompt will ask for confirmation to rebuild the webpage. The utility will overwrite the existing webpage of the same name (if it exists), but does not modify any other files or folders.


# How to Use the Webpage
Use the search bar at the top of the page to type a string. The page will filter out all files that do not contain that text. It is not case sensitive. If a folder has no files that match, the page hides that folder.

![Applying search filters](https://github.com/M-Scott-Lassiter/Excel-File-Cataloger/blob/main/Images/ApplySearchFilters.gif)

Select a tag to show only those files that have already been identified as applicable to that group. Searching now only looks within those tagged files. Press the "Show All" button to clear the search bar and tag selections.

![Applying tag filters and searching](https://github.com/M-Scott-Lassiter/Excel-File-Cataloger/blob/main/Images/ApplyTagsAndFilter.gif)

If using an omit prefix, the webpage will not have any of the omitted files present.

![Showing effect of omit prefixes](https://github.com/M-Scott-Lassiter/Excel-File-Cataloger/blob/main/Images/ApplyingOmitPrefixes.gif)


# Contributing and Outlook
This resource has all intended functionality as of version 1.1. I consider it feature complete, but will continue to provide bug support.

I will in no way turn away additional contributions or expansions if beneficial or needed in the future. All are welcome to open an issue or feature request. If you decide to contribute, here is more information about the utility's structure:

## Directory Management
To parse all the files and folders, this utility makes use of the [VBA Directory Manager](https://github.com/M-Scott-Lassiter/VBA-Directory-Manager) class.

## Settings
The settings are stored in a table on a sheet called `SettingsSheet`. This has visibility set to `xlSheetVeryHidden`, and the only way to unhide it is going into the VBA IDE and changing the setting.

## Modules
- **ContentManagement**: Contains the user interface code and the script to initiate rebuilding the webpage.
- **CSS**: Contains the public script `InsertCSS`. All CSS code injects in the HTML `<head>` element.
- **GNU_GPL_v3**: Contains the text of the [license](./LICENSE). This must be included as a term of the license.
- **HTML**: Contains the public script `InsertHTML` to add each of the HTML elements.
- **Javascript**: Contains the public script `InsertJavascript`. All CSS code injects in the HTML `<head>` element.

All HTML, CSS, and Javascript is self contained within a single file.

## Tags Table and Data Validation
Tags are stored in a table on the WebpageBuilder sheet named "Tags". The "Tag Name" column has data validation applied to prevent duplicate entries. To make the data validation apply dynamically to the whole column, a named range called `TagNameList` is defined in the Name Manager.


# License
Distributed under the [GNU General Pulbic Licence v3.0 (GPLv3)](./LICENSE), copyright 2022.


# Contact
Reach me on [LinkedIn](https://www.linkedin.com/in/mscottlassiter/) or [Twitter](https://twitter.com/MScottLassiter).
