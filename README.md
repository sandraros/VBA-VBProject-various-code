# VBA-VBProject-various-code
Subroutines to export and import whole VBA project, and a few other utility subroutines.

Concept:
- The export exports the VBA components of your document into VBA files (.bas, .frm, .frx, .cls) in the same folder as the document
- The import imports all VBA files (.bas, .frm, .frx, .cls) from the folder of the current document into your document
- There is some logic to inform you that the folder contains files that do not correspond to your project, and to let you confirm the replacement of files and macros
- If needed, you may easily create your own macro by combining the existing subroutines

I guess they work in all Microsoft Office applications that offer VBA: Word, Excel, PowerPoint, Outlook. (and other?)

Install the module in your document with macro enabled:
- Create module `ExportImportVBProject` (because this name is hardcoded in several subroutines, like the export, import and at a few other places, to exclude this module; but you may change the constant)
- Copy the source code given here and paste it into this module
- Give programmatic access to the project so that it can export and import code, as explained here: https://support.microsoft.com/en-us/help/282830/programmatic-access-to-office-vba-project-is-denied

**BE CAREFUL: do a backup of your Office document before running the subroutines, there can be some nasty bugs left! (USE THEM AT YOUR OWN RISK)**

My preferred subroutines:
- **ExportVBProjectDialog_VBAEditor** : I run it from the VBA editor to backup the active project \
  There's a popup to confirm the export if some files exist in the folder which don't correspond to the VB components to export, so that a later import doesn't add these parasite components, for instance: *Folder contains unrelated VB components, they should not be here (at least 'test.bas')*
- **ImportVBProjectDialog_VBAEditor** : run it from the VBA editor to restore the active project \
  There's a popup to confirm the import, for instance: *IMPORT! That will add 3 VB components and replace 5 of them. Do you want to continue? (type "YES" for yes)*
- **CleanUpVBProjectDialog_BackupFolder_VBAEditor** : does an export + import to cleanup compile bugs, and the files are kept in the folder

Miscellaneous:
- Subroutines essentially don't have dialogs to query the user, except those suffixed `Dialog`. Others send error (numbers starting from arbitrary 64230).
- The subroutines work for a document located in Microsoft **OneDrive**. The code was taken from these two places and adapted to make them work with any Microsoft Office application (not only Excel):
  - https://social.msdn.microsoft.com/Forums/office/en-US/1331519b-1dd1-4aa0-8f4f-0453e1647f57/how-to-get-physical-path-instead-of-url-onedrive
  - https://stackoverflow.com/questions/33734706/excels-fullname-property-with-onedrive/33935405 

Main utility subroutines:
- Export/Import
  - ExportVBProject
  - ExportVBProjectDialog
  - ImportVBProject
  - ImportVBProjectDialog
  - CleanUpVBProject_VBAEditor
    - Possibly useful for solving issue like "User-Defined Type Not Defined" (issue described here: https://stackoverflow.com/questions/19680402/compile-throws-a-user-defined-type-not-defined-error-but-does-not-go-to-the-of)
- Export/Import helper
  - CheckFolderFreeOfUnrelatedVBComponentFiles
  - CheckFolderHasNoVBComponentFiles
- Delete VB modules
- VBA Editor windows
  - Close all VB components
  - Open all VB components
