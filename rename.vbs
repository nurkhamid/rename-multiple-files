Set objFso = CreateObject("Scripting.FileSystemObject")

Set Folder = objFSO.GetFolder("C:\Users\{Your-path-here}")

For Each File In Folder.Files

    sNewFile = File.Name

    sNewFile = Replace(sNewFile,"OLD","NEW")

    if (sNewFile<>File.Name) then

        File.Move(File.ParentFolder+"\"+sNewFile)

    end if

Next