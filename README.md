# Reading a text file in ASP.

How to read a text file on a server using VBScript in ASP

- [Download readfile.zip - 1.2 KB](https://raw.githubusercontent.com/ChrisMaunder/readfile/master/docs/assets/readfile.zip)

One of the most important tasks in any programming language is the ability to read and write files. The steps involved in ASP are no different than many other languages:

1. Specify the location of the file
2. Determine if the file exists
3. Get a handle to the file
4. Read the contents
5. Close the file and release any resources used

File I/O in ASP can be done using the `FileSystemObject `component. When opening a text file you simply open it as a text stream, and it is this text stream that you use to access the contents of the file.

The FileSystemObject allows you to perform all file and folder handling operations. It can either return a file which can then be opened as a text stream, or it can return a text stream object directly.

In the following I present two different methods. The first method gets a file object and uses that to open the text stream, and the second method opens the text stream directly from the `FileSystemObject`.

### Method 1:

```cpp
<% Option Explicit

Const Filename = "/readme.txt"    ' file to read
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

' Create a filesystem object
Dim FSO
set FSO = server.createObject("Scripting.FileSystemObject")

' Map the logical path to the physical system path
Dim Filepath
Filepath = Server.MapPath(Filename)

if FSO.FileExists(Filepath) Then

    ' Get a handle to the file
    Dim file    
    set file = FSO.GetFile(Filepath)

    ' Get some info about the file
    Dim FileSize
    FileSize = file.Size

    Response.Write "<p><b>File: " & Filename & " (size " & FileSize  &_
                   " bytes)</b></p><hr>"
    Response.Write "<pre>"

    ' Open the file
    Dim TextStream
    Set TextStream = file.OpenAsTextStream(ForReading, TristateUseDefault)

    ' Read the file line by line
    Do While Not TextStream.AtEndOfStream
        Dim Line
        Line = TextStream.readline
    
        ' Do something with "Line"
        Line = Line & vbCRLF
    
        Response.write Line 
    Loop

    Response.Write "</pre><hr>"

    Set TextStream = nothing
    
Else

    Response.Write "<h3><i><font color=red> File " & Filename &_
                       " does not exist</font></i></h3>"

End If

Set FSO = nothing
%>
```

### Method 2:

```cpp
<% Option Explicit

Const Filename = "/readme.txt"    ' file to read
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

' Create a filesystem object
Dim FSO
set FSO = server.createObject("Scripting.FileSystemObject")

' Map the logical path to the physical system path
Dim Filepath
Filepath = Server.MapPath(Filename)

if FSO.FileExists(Filepath) Then

    Set TextStream = FSO.OpenTextFile(Filepath, ForReading, False, TristateUseDefault)

    ' Read file in one hit
    Dim Contents
    Contents = TextStream.ReadAll
    Response.write "<pre>" & Contents & "</pre><hr>"
    TextStream.Close
    Set TextStream = nothing
    
Else

    Response.Write "<h3><i><font color=red> File " & Filename &_
                   " does not exist</font></i></h3>"

End If

Set FSO = nothing
%>
```
