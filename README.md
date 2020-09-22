<div align="center">

## Disk Space Usage<br/>by Agnus Imonar


</div>

### Description

This code will tell you how much disk space on your server you're using
 
### More Info
 
Noke

Just paste the code into an .asp page on the root dir of your web

The disk space you're using on your server


<span>             |<span>
---                |---
**Category**       |Miscellaneous
**Level**          |Intermediate
**User Rating**    |5.0 (10 votes by 2 users)|
**Compatibility**  |ASP (Active Server Pages), HTML, VbScript (browser/client side)






### Source Code

<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=windows-1252">
<META name="GENERATOR" content="Microsoft FrontPage 4.0">
<META name="ProgId" content="FrontPage.Editor.Document">
<TITLE>New Page 1</TITLE>
</HEAD>
<BODY>
<P>&lt;%<BR>
'Capture the name of the page as well as directory structure&nbsp;<BR>
script_name=request.servervariables("script_name")<BR>
<BR>
'Split the directory tree into an arry by /<BR>
split_name=split(script_name,"/")<BR>
<BR>
' Sets the number of directory levels down<BR>
num_directory=ubound(split_name)-1<BR>
<BR>
%><BR>
&lt;html><BR>
&lt;title>CodeAve.com(Directory Size)&lt;/title><BR>
&lt;body bgcolor="#FFFFFF"><BR>
<BR>
&lt;table align="center"><BR>
&lt;tr><BR>
&lt;td width=150><BR>
&lt;b>Directory&lt;/b><BR>
&lt;/td><BR>
&lt;td width=150><BR>
&lt;b>Megabytes&lt;/b><BR>
&lt;/td><BR>
&lt;td width=150><BR>
&lt;b>Kilobytes&lt;/b><BR>
&lt;/td><BR>
&lt;td width=150><BR>
&lt;b>Bytes&lt;/b><BR>
&lt;/td><BR>
&lt;/tr><BR>
&lt;%<BR>
' Create a file system object to read all the directories<BR>
' beneath the current directory split_name(num_directory)<BR>
' You can hard code the directory name if you like<BR>
set directory=server.createobject("scripting.filesystemobject")<BR>
set allfiles=directory.getfolder(server.mappath("../"&amp; split_name(num_directory)&amp; "/"))<BR>
<BR>
' Lists all the files found in the directory<BR>
for each directory in allFiles.subfolders<BR>
' Removes certain MSFrontPage was directories&nbsp;<BR>
if right(directory.Name,3) &lt;> "cnf" then&nbsp;<BR>
'Adds the folder sizes up for a total<BR>
total_size=total_size + directory.size %><BR>
<BR>
&lt;tr><BR>
&lt;td width=150><BR>
&lt;%= directory.name %><BR>
&lt;/td><BR>
&lt;td width=150>&lt;%= formatnumber((directory.size/1024/1024),2) %>&lt;/td><BR>
&lt;td width=150>&lt;%= formatnumber((directory.size/1024),0) %>&lt;/td>&nbsp;<BR>
&lt;td width=150>&lt;%= formatnumber(directory.size,0) %>&lt;/td>&nbsp;<BR>
&lt;/tr><BR>
&lt;% end if 'end check for FrontPage directories&nbsp;<BR>
next 'end of the for next loop %><BR>
&lt;tr><BR>
&lt;td width=150>&lt;b>Total&lt;/b>&lt;/td><BR>
&lt;td width=150>&lt;%= formatnumber((total_size/1024/1024),2) %>&lt;/td><BR>
&lt;td width=150>&lt;%= formatnumber((total_size/1024),0) %>&lt;/td>&nbsp;<BR>
&lt;td width=150>&lt;%= formatnumber(total_size,0) %>&lt;/td>&nbsp;<BR>
&lt;/tr><BR>
&lt;/table><BR>
<BR>
&lt;/body><BR>
&lt;/html></P>
</BODY>
</HTML>

