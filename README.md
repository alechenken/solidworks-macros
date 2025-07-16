# solidworks-macros
A collection of my custom macros that I use for Solidworks 2022

# License
Some scripts are slightly modified versions of scripts I found on forums or on your site, and I've lost all attribution. If you find that I've misattributed your code, please notify me and I will correct it immediately.

# Installation

All scripts are located in the /src/ folder.

These files are *.bas files, which is how Solidworks exports macros. There are two methods to install these:
### Method 1: Copy Paste
1. Any of these .bas files can be opened in a text editor like Notepad++.
2. Simply copy and paste the text into a new Solidworks macro and save as desired.

### Method 2: Import
1. Solidworks VBA editor allows importing. Simply import the .bas file inside the File>Import menu.

<img width="341" height="242" alt="image" src="https://github.com/user-attachments/assets/b64ff691-3bb3-4a7c-bf9f-66570481be74" />

# Script Descriptions
## ✏ hide-reference-geometry-sketches
### Works for:
✅ Parts  
✅ Assemblies (Iterates through all parts in an assembly)  
🟥 Drawings

### Functionality
This script will pop up with a window asking which features you would like to hide. I primarily use this to hide all planes, axes, and sketches instead of setting the show/hide functions in the View module. It keeps your part files clean and tidy, and is good for geometry that will be viewed in an offline 3D viewer by say, a customer or other department.

<img width="221" height="357" alt="image" src="https://github.com/user-attachments/assets/d9c2a23b-db66-4c72-b2ff-67b6fa83683c" />

## ✏ pdf-to-path
### Works for:
🟥 Parts  
🟥 Assemblies  
✅ Drawings

### Functionality
The script will instantly export the currently open drawing (all sheets) to a pdf file. The default functionality is to ask for a path every time, but if you have a preferred folder that you always use, you can hard-code the path into the script by changing the commenting scheme in the indicated lines:

Asks for destination folder every time (default behavior):
```vba
'UNCOMMENT THIS LINE TO ASK PATH EACH TIME
outFolder = BrowseForFolder()
'UNCOMMENT THIS LINE TO SPECIFY A SPECIFIC PATH AUTOMATICALLY EVERY TIME
'******************TYPE YOUR PATH IN THE QUOTES BELOW*********************
'outFolder = "C:\Users\alechenken\Documents\Drawing pdf Export"
'*************************************************************************
```

Always exports pdf into the same folder (that you fill in yourself instead of "C:\Users....."):
```vba
'UNCOMMENT THIS LINE TO ASK PATH EACH TIME
'outFolder = BrowseForFolder()
'UNCOMMENT THIS LINE TO SPECIFY A SPECIFIC PATH AUTOMATICALLY EVERY TIME
'******************TYPE YOUR PATH IN THE QUOTES BELOW*********************
outFolder = "C:\Users\alechenken\Documents\Drawing pdf Export"
'*************************************************************************
```

## ✏ remove-transparencies
### Works for:
🟥 Parts  
✅ Assemblies (Iterates through all parts in an assembly)  
🟥 Drawings

### Functionality
Will iterate through all parts in an assembly and change them all to opaque. Transparent parts are frequently used in design to see interferences or fit between parts, or simply to see behind them. This is a quick way to change all your parts back to opaque before you save them.

## ✏ remove-with-thread-callout
### Works for:
✅ Parts  
🟥 Assemblies (Iterates through all parts in an assembly)  
🟥 Drawings

### Functionality
This script will iterate through all Hole Wizard threaded holes in a part and uncheck the box that says "With thread callout", since Solidworks is not reliable at keeping this box unchecked.

<img width="569" height="1013" alt="image" src="https://github.com/user-attachments/assets/925be7a9-14a0-40be-bb87-0dc72a75aef5" />

## ✏ set-all-docfonts
### Works for:
🟥 Parts  
🟥 Assemblies  
✅ Drawings

### Functionality
Will change all fonts in Document Properties to the font of your choice. This includes ALL menus in the Document Properties, including all items in the Annotations, Dimensions, Tables, and Detailing sub-categories. This is useful if you have updated the font in your drawing template, but you have to open an old part that's based on an old template. This will bring the font up to date without clicking a million times. Currently this typeface family is set to "Century Gothic", but you can select your desired typeface family by editing the appropriate line in the script:

```vba
' Loop through each text format and apply the font settings. Replace Century Gothic with your preferred typeface name
Dim i As Integer
For i = 0 To UBound(textFormats)
    Set TextFormatObj = ModelDocExtension.GetUserPreferenceTextFormat(textFormats(i), 0)
    Set swTextFormat = TextFormatObj
    swTextFormat.TypeFaceName = "Century Gothic"
    boolstatus = ModelDocExtension.SetUserPreferenceTextFormat(textFormats(i), 0, swTextFormat)
Next i
```

<img width="521" height="234" alt="image" src="https://github.com/user-attachments/assets/cb9a02d7-5642-48b2-bd6b-ed2e6110b822" />

