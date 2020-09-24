<%@ Language=VBScript %>
<%
	Dim objMail, strBody, strSalutation, strSubject, strSend
	
	Set objMail = Server.CreateObject("CDONTS.NEWMAIL")
	
	'strSalutation = "Dear Administrator"
	strSubject = Request.Form("txtSubject")
	strBody = Request.Form("txtMessage")
	strTo= Request.form("txtTo")
	objMail.To = strTo 
	objMail.From = Request.Form("txtEmail")
	objMail.Subject = strSubject
	objMail.Body = strBody
	
		
	objMail.Send
	set objMail = Nothing


%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY  background="ame.jpg">
Your mail has been sent.<br>
Name:  <% = Request.Form("txtName") %> <br>
Subject:  <% = Request.Form("txtSubject") %> <br>
From:  <% = Request.Form("txtEmail") %> <br>
To:  <% = Request.Form("txtTo") %> <br>
Message:  <% =Request.Form("txtMessage") %><br>

<P>&nbsp;</P>
<hr>

<P>
Please do not remove this line or the following:<br>
Code 
Revision by Michael Heath. Original concept and code by Tim Butler </P>
<P>
            
<hr>

<P></P>
<P><!--#include file="Readme.txt"--></P>
</BODY>
</HTML>
