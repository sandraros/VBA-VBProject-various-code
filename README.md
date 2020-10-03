# VBA-VBProject-various-code
Subroutines to export and import whole VBA project, and a few other utility subroutines.

I guess they work in all Microsoft Office applications which offer VBA: Word, Excel, PowerPoint, Outlook. (and other?)

Installation in VBA Editor (start your Office application and press Alt+F11):
- Create module `ExportImportVBProject` (because this name is hardcoded in several subroutines, like the export, import and at a few other places, to exclude this module)
- Copy the code given here and paste it into this module
- Add references in menu Runtime > References
  - Microsoft Visual Basic for Applications Extensibility (possibly located at `C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB`)
  - Microsoft Scripting Runtime (possibly located at `C:\Windows\SysWOW64\scrrun.dll`)

Subroutines:
- Export/Import
  - direction
    - Export
    - Import
    - Export and Import (`CleanUpVBProject_VBAEditor`)
      - Possibly useful for solving issue like "User-Defined Type Not Defined" (issue described here: https://stackoverflow.com/questions/19680402/compile-throws-a-user-defined-type-not-defined-error-but-does-not-go-to-the-of)
  - running it from
    - active VBA Editor window
    - active MS Word document
    - active MS Excel document
    - active MS PowerPoint document
    - MS Outlook
- Delete VB modules
- VBA Editor windows
  - Close all of them
  - Open all VB components

Miscellaneous:
- Subroutines essentially don't have dialogs to query the user, except those suffixed `Dialog`. Others send error (numbers starting from arbitrary 64230).
- The subroutines work for a document located in Microsoft **OneDrive**. The code was taken from these two places and adapted to make them work with any Microsoft Office application (not only Excel):
  - https://social.msdn.microsoft.com/Forums/office/en-US/1331519b-1dd1-4aa0-8f4f-0453e1647f57/how-to-get-physical-path-instead-of-url-onedrive
  - https://stackoverflow.com/questions/33734706/excels-fullname-property-with-onedrive/33935405 
