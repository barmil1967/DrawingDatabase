<%
  Dim strPathInfo, strPhysicalPath
  Dim objFSO, objFile, objFileItem, objFolder, objFolderContents

	strPhysicalPath = "c:\files"
	Set objFSO = CreateObject("Scripting.FileSystemObject")

%>

    <HTML>

    <HEAD>
        <link rel="stylesheet" href="css/normalize.min.css?version=1.2">
        <link rel="stylesheet" href="css/style.css?version=1.7">

        <style>
            body {
                color: #aaaaaa;
                background: #151515;
            }
        </style>
        <TITLE>Drawing Database - ALL Files</TITLE>
    </HEAD>

    <BODY>
        <p><img src="/FileManager/Images/RedStagLogo.jpg" alt="Red Stag Logo" height="156" width="280" class="center"></p>
        <p>
            <h1>Red Stag Timber Drawing Database</h1>
        </p>
        <p>
            <h2>Under Construction</h2>
        </p>
        <TABLE cellpadding=5>
            <TR align=center>
                <td align=left>File Name</td>
                <td align=left>File Size</td>
                <td align=left>Last Accessed</td>
                <td align=left>Last Modified</td>
                <td align=left>Created</td>
            </TR>
            <%
  ShowSubFolders objFSO.GetFolder(strPhysicalPath)
  
  Sub ShowSubFolders(Folder)
  
  For Each Subfolder in Folder.Subfolders
    Set objFolder = objFSO.GetFolder(Subfolder.Path)
        Response.Write "<TR><TD align=left colspan=5>"
		Response.Write objFolder 
		Response.Write "</TD></tr>"
    
    Set objFile = objFolder.Files
	For Each objFileItem In objFile
		Response.Write "<tr></tr><TD align=left>"
		Response.Write objFileItem.Name 
		Response.Write "</TD><TD align=left>" 
		Response.Write objFileItem.Size 
		Response.Write "</TD><TD align=right>" 
		Response.Write objFileItem.DateLastAccessed 
		Response.Write "</TD><TD align=right>" 
		Response.Write objFileItem.DateLastModified 
		Response.Write "</TD><TD align=right>"
		Response.Write objFileItem.DateCreated 
		Response.Write "</TD></TR>"
	Next
	ShowSubFolders Subfolder
  Next
  End Sub
%>
        </table>
    </body>

    </html>