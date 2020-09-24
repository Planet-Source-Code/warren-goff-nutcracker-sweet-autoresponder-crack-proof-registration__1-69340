<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY Background="ame.jpg">
<center>
<font size="6"><font color="Blue">Anon Email</font></font>
</center><br><br>
<form method="post" name= "thisform" action="mailall.asp">

<P>&nbsp;</P>
<P>Name:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<INPUT id=text1 name=txtName></P>
<P>Subject:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<INPUT id=text2 name=txtSubject></P>
<P>From 
Email:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<INPUT id=text3 name=txtEmail></P>
<P>To 
Email:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<INPUT id=text4 name=txtTo></P>
<P>Message:</P>
<P>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<TEXTAREA id=TEXTAREA1 name=txtMessage style="HEIGHT: 120px; WIDTH: 600px"></TEXTAREA></P>
<P><INPUT id=submit1 name=submit1 type=submit value=Send>&nbsp; <INPUT id=reset1 name=reset1 type=reset value=Clear></P>
</form>
<hr>
Please do not remove this line or the following:<br>
Code Revision by Michael Heath.  Original concept and code by Tim Butler
<hr>
<!--#include file="include.htm"-->
<hr>
<!--#include file="Readme.txt"-->
</BODY>
</HTML>
