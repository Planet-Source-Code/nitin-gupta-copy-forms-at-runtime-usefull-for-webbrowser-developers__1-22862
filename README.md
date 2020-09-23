<div align="center">

## Copy forms at Runtime \(usefull for WebBrowser developers\)


</div>

### Description

I have observed that many browsers are coming up, but none supports the same format when a new window is opened, simply because they use IE control. Here's a code for guys who want to create forms at Runtime (Copy their previous forms onto new ones). Plz vote if u use it !!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Nitin Gupta](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nitin-gupta.md)
**Level**          |Intermediate
**User Rating**    |4.8 (48 globes from 10 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/nitin-gupta-copy-forms-at-runtime-usefull-for-webbrowser-developers__1-22862/archive/master.zip)





### Source Code

<B>For normal coders :</B><BR><BR>
Dim frmCopy As Form1<BR>
Set frmCopy = New Form1<BR>
frmCopy.Visible = True<BR><BR>
<B>For WebBrowser developers :</B><BR><BR>
Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)<BR>
Dim frmNew As Form1<BR>
Set frmNew = New Form1<BR>
frmNew.WebBrowser1.RegisterAsBrowser = True<BR>
Set ppDisp = frmNew.WebBrowser1.Object<BR>
frmNew.Visible = True<BR>
End Sub<BR>

