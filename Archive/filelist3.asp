<%
  Dim strPathInfo, strPhysicalPath
  strPhysicalPath = "C:\Program Files"

  Dim objFSO, objFile, objFileItem, objFolder, objFolderContents

  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objFolder = objFSO.GetFolder(strPhysicalPath)
  Set objFile = objFolder.Files
%>

<HTML>
 <HEAD>
	<link rel="stylesheet" href="css/normalize.min.css?version=1.1">
	<link rel="stylesheet" href="css/style.css?version=1.4">

<style>
body {
color: #aaaaaa;
background: #151515;
}
</style>
  <TITLE>Drawing Database - ALL Files</TITLE>
 </HEAD>
<BODY>
<TABLE cellpadding=5>
 <TR align=center>
  <td align=left>Folder Name</td>
  <td align=left>File Name</td>
  <td align=left>File Size</td>
  <td align=left>Last Accessed</td>
  <td align=left>Last Modified</td>
  <td align=left>Created</td>
</TR>
<%
  For Each objFileItem In objFile
    Response.Write "<TR><TD align=left>"
    Response.Write objFolder 
    Response.Write "</TD><TD align=left>"
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
%>
</table>
  <script src='js/jquery.min.js'></script>
</body>
</html>