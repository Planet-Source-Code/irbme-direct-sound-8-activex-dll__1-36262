To use the dll in your project or the test project:

EITHER

 - Delete the ".RENAME" part from the end of "DS8Wrapper.dll.RENAME" so that it becomes just "DS8Wrapper.dll"

OR

 - Delete "DS8Wrapper.dll.RENAME" then open up the source and recompile it into a dll

THEN

- Open up your project and click the "Project" menu. Then click on "References..."
- Click the browse button"
- Find the "DS8Wrapper.dll" file and add it to the list then press OK
- Use the following skeleton in new projects:

Private SW as DirectSound8Wrapper

Private Sub Form_Load()
  Set SW = New DirectSound8Wrapper
end sub

Private Sub Form_Unload()
  Set SW = nothing
End Sub