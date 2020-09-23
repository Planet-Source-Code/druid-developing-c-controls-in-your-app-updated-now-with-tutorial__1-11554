<div align="center">

## C\+\+ Controls in your app \(updated\) \- now with tutorial

<img src="PIC20009191616207156.jpg">
</div>

### Description

REAL C++ CONTROLS IN YOUR APP!!!

Do you know the code "Real C++ Buttons" by Randy Mcdowell?

It sends a message to a VB Button to make it a "C++ Button". But if you look at this button with

Spy++ you will see that it is still a "ThunderCommandButton", the VB Button.

My code lets you use EVERY C++ control in your app!!! And these controls are REAL C++ controls!!!

Think of the fantastic controls you can have: e.g. the cool hotkey control. Or what about a "real"

RICHEDIT control? The only limitation is your fantasy! Control events are also supported.

Well commented, and easy to use even if you don't understand everything that happens.

The second version of this code, now with tutorial and some new stuff.

Have fun with this code, use it in your app, and if you like it VOTE FOR IT!!!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-09-19 22:10:26
**By**             |[Druid Developing](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/druid-developing.md)
**Level**          |Advanced
**User Rating**    |4.6 (37 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD99959192000\.zip](https://github.com/Planet-Source-Code/druid-developing-c-controls-in-your-app-updated-now-with-tutorial__1-11554/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Language" content="de">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Neue Seite 1</title>
<meta name="Microsoft Theme" content="none, default">
</head>
<body>
<p align="center"><font face="Papyrus">(best viewed in 1024 x 768)</font></p>
<p align="center"><font size="6" color="#FF0000" face="Papyrus">C++
Controls in your app - Now with tutorial</font></p>
<p align="center"><font size="5" face="Papyrus">Hi, this is the update to my
code &quot;Real C++ Controls in your app&quot; which I submitted at the
beginning of September.</font></p>
<p align="center"><font size="5" face="Papyrus">Now somebody posted that this
code isn´t explained well and so I wrote this tutorial.</font></p>
<p align="center"><font size="5" face="Papyrus">If there would still be any
problems just <b> E-Mail</b> me at</font></p>
<p align="center"><a href="mailto:druid-developing@gmx.de"><font face="Papyrus" size="5">druid-developing@gmx.de</font></a></p>
<p align="center"><font face="Papyrus" size="2"><b>Important note: This tutorial
is also included in the .zip file. You don´t have to read it here.</b></font></p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Papyrus" size="4" color="#0000FF">1. How to use
this code in your app</font></p>
<p align="center"><font face="Papyrus" color="#000000">First you have to include
the modMain.bas in your project.</font></p>
<p align="center"><font face="Papyrus" color="#000000">Then goto the menu &quot;Project&quot;
and click &quot;Properties of ...&quot;.</font></p>
<p align="center"><font face="Papyrus" color="#000000">In this window set the
Start Object to &quot;Sub Main&quot;.</font></p>
<p align="center"><font face="Papyrus" color="#000000">In the Sub Main which is
in the modMain you can now create the controls.</font></p>
<p align="center"><font face="Papyrus" color="#000000">Call the function like
this:</font></p>
<p align="center">&nbsp;</p>
<p align="center"><b><font face="Tahoma" color="#000000" size="2">Hwnd of the
control = CreateControl( &quot;Edit&quot; (Classname) , &quot;This is a TextBox&quot;
(Text) , 3 (Left) , 3 (Top) , 100 (Width) , 50 (Height) , (Optional Style) )</font></b></p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Papyrus">That´s it! No difficult API Calls, not
much code, just <font color="#FF0000">ONE FUNCTION!</font></font></p>
<p align="center"><font face="Papyrus" color="#000000">Very easy to use, even
for beginners.</font></p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Papyrus" size="4" color="#0000FF">2. How to
interact with the controls</font></p>
<p align="center"><font face="Papyrus" color="#000000">If you want to use the
controls like normal controls, with Events and Properties it is a bit more
difficult.</font></p>
<p align="center"><font face="Papyrus" color="#000000">For&nbsp; every Property
and Event you firs need the WindowHandle of the control.</font></p>
<p align="center"><font face="Papyrus" color="#000000">You get it from the
CreateControl function (look above).</font></p>
<p align="center"><font face="Papyrus" color="#000000">If you want to get e.g.
the Text of a created TextBox control you can do it like this:</font></p>
<p align="center">&nbsp;</p>
<p align="center"><b><font color="#66FF66"><font face="Tahoma" size="2">'</font><font face="Tahoma" size="2">Declare
Variable to save the WindowHandle</font></font></b></p>
<p align="center"><b><font face="Tahoma" size="2">Public TextBoxHwnd as Long</font></b></p>
<p align="center"><b><font face="Tahoma" size="2" color="#66FF66">'Create the
TextBox</font></b></p>
<p align="center"><font color="#000000" face="Tahoma" size="2"><b>TextBoxHwnd =
CreateControl( &quot;Edit&quot; , &quot;Text to get&quot; ,&nbsp; 3 , 3 , 100 ,
40 )</b></font></p>
<p align="center">&nbsp;</p>
<p align="center"><font color="#000000" face="Papyrus">Using this Function you
can get the actual Text of the TextBox</font></p>
<p align="center"><font face="Tahoma" size="2" color="#000000"><b>Function
Get_Text_Of_Control(ByVal cHwnd as Long) as String</b></font></p>
<p align="center"><b><font face="Tahoma" size="2" color="#000000">Dim
ControlText As String</font></b></p>
<p align="center"><b><font face="Tahoma" size="2" color="#000000">ControlText = Space(254)</font></b></p>
<p align="center"><b><font face="Tahoma" size="2" color="#66FF66">'Use the GetWindowText API to get the actual
text of the TextBox control</font></b></p>
<p align="center"><b><font face="Tahoma" size="2" color="#000000">    GetWindowText
c</font></b><font face="Tahoma" size="2" color="#000000"><b>Hwnd </b></font><b><font face="Tahoma" size="2" color="#000000">,
ControlText , 254</font></b></p>
<p align="center"><font face="Tahoma" size="2" color="#000000"><b>Get_Text_Of_Control</b></font><b><font face="Tahoma" size="2" color="#000000">
= Trim(</font></b><b><font face="Tahoma" size="2" color="#000000">ControlText</font></b><b><font face="Tahoma" size="2" color="#000000">)</font></b></p>
<p align="center"><b><font face="Tahoma" size="2" color="#000000">End Function</font></b></p>
<p align="center">&nbsp;</p>
<p align="center"><font color="#000000" face="Papyrus">Use this function like
this:</font></p>
<p align="center"><b><font face="Tahoma" size="2" color="#000000">TextBoxText =
Get_Text_Of_Control(TextBoxHwnd)</font></b></p>
<p align="center"><b><font face="Tahoma" size="2" color="#000000">MsgBox
TextBoxText</font></b></p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Papyrus">To use an Event, e.g. the click Event of
a Button you can do it like this:</font></p>
<p align="center">&nbsp;</p>
<p align="center"><b><font face="Tahoma" size="2" color="#66FF66">'Declare
Variable to save the WindowHandle</font></b></p>
<p align="center"><b><font face="Tahoma" size="2">Public ButtonHwnd as Long</font></b></p>
<p align="center"><font color="#66FF66" face="Tahoma" size="2"><b> 'To save the old WindowProcedure for the button</b></font></p>
<p align="center"><b><font face="Tahoma" size="2">Public gButOldProc as Long</font></b></p>
<p align="center">&nbsp;</p>
<p align="center"><b><font face="Tahoma" size="2" color="#66FF66">'Create the
Button</font></b></p>
<p align="center"><b><font face="Tahoma" size="2">ButtonHwnd </font></b><font color="#000000" face="Tahoma" size="2"><b>=
CreateControl( &quot;Button&quot; , &quot;Click this button&quot; ,&nbsp; 3 , 3
, 100 , 40 )</b></font></p>
<p align="center"><b><font face="Tahoma" size="2" color="#66FF66">  'Get the address of the standard button procedure and save it in
&quot;gButOldProc&quot;</font></b></p>
<p align="center"><b><font face="Tahoma" size="2">gButOldProc&amp; = GetWindowLong(ButtonHwnd&amp;,
GWL_WNDPROC)</font></b></p>
<p align="center"><b><font face="Tahoma" size="2"><font color="#66FF66">'Use GWL_WNDPROC to save the adress of the procedure for the
button</font></font></b></p>
<p align="center"><b><font color="#66FF66" face="Tahoma" size="2">'You have to do this for every control you want to have a procedure</font></b></p>
<p align="center"><b><font face="Tahoma" size="2">  Call SetWindowLong(ButtonHwnd&amp;, GWL_WNDPROC, GetAddress(AddressOf ButtonWndProc))</font></b></p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Tahoma" size="2"><b><font color="#66FF66">'This is the procedure that is called when you click the button</font></b></font></p>
<p align="center"><font face="Tahoma" size="2"><b>Public Function ButtonWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long</b></font></p>
<p align="center"><font face="Tahoma" size="2"><b>  Select Case uMsg&amp;</b></font></p>
<p align="center"><font face="Tahoma" size="2"><b>  Case WM_LBUTTONUP:</b></font></p>
<p align="center"><font face="Tahoma" size="2"><b><font color="#66FF66">'Left button is up (user clicked the Button)</font></b></font></p>
<p align="center"><font face="Tahoma" size="2"><b><font color="#66FF66">'Use
&quot;WM_LBUTTONDOWN&quot;</font></b></font></p>
<p align="center"><font face="Tahoma" size="2"><b><font color="#000000">MsgBox
&quot;The button was clicked&quot;</font></b></font></p>
<p align="center"><font face="Tahoma" size="2"><b><font color="#66FF66">'Call the standard window proc</font></b></font></p>
<p align="center"><font face="Tahoma" size="2"><b>  ButtonWndProc = CallWindowProc(gButOldProc&amp;, hwnd&amp;, uMsg&amp;, wParam&amp;, lParam&amp;)</b></font></p>
<p align="center"><font face="Tahoma" size="2"><b>End Function</b></font></p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Papyrus" size="4" color="#0000FF">3. Final
Explanations</font></p>
<p align="center"><font face="Papyrus" color="#000000">The special thing on this
code is that you can use every registered Windows class name for a control.</font></p>
<p align="center"><font face="Papyrus" color="#000000">You can also create an
own class name using the API &quot;RegisterWindowClass&quot;.</font></p>
<p align="center"><font face="Papyrus" color="#000000">That´s all for today,
bye.</font></p>
<p align="center"><font face="Papyrus" color="#000000">Maybe I update this code
once more.</font></p>
<p align="center"><font face="Papyrus" color="#000000"><b>PS: Please excuse me
for my bad English, I´m German.</b></font></p>
<p align="center"><b><font face="Papyrus" size="4" color="#FF0000">And PLEASE,
PLEASE, PLEASE VOTE FOR ME!!!</font></b></p>
</body>
</html>

