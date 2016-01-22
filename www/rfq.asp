<%@ Language=VBScript %>


<%	   
	Dim lobjMl 
	dim strbody
	
	if(Request.form("sendflg") = "1")	then
		
	strFrom = "InfraredOptics.com"
	strSubject = "RFQ by " & Request.form("txtFname") 
	strTo = "iroptical@aol.com"
	
	strbody =strbody & vbcrlf
	
	strBody = "The following is the request for Quote "  & vbCrLf & vbCrLf
	strBody = strBody & "Pricing Detail for Part No. " & vbTab & Request.form("txtPartNo") & "."  & vbCrLf 
	strBody = strBody & "First Name "& vbTab & vbTab & vbTab & vbTab & Request.form("txtFname") & "."  & vbCrLf 
	strBody = strBody & "Last Name " & vbTab & vbTab & vbTab & vbTab & Request.form("txtLname") & "." & vbCrLf 
	strBody = strBody & "Mode of sending the Quote " & vbTab  & Request.form("rdChoice") & "." & vbCrLf 
	strBody = strBody & "Street Address " & vbTab & vbTab  & vbTab & Request.form("txtStreet") & "." & vbCrLf 
	strBody = strBody & "City " & vbTab & vbTab & vbTab & vbTab & vbTab & Request.form("txtCity") & "." & vbCrLf 
	strBody = strBody & "State " & vbTab & vbTab & vbTab & vbTab & Request.form("txtState") & "." & vbCrLf 
	strBody = strBody & "Zip Code " & vbTab & vbTab & vbTab & vbTab & Request.form("txtZip") & "."  & vbCrLf 
	strBody = strBody & "Country " & vbTab & vbTab & vbTab & vbTab & Request.form("txtCountry") & "." & vbCrLf 
	strBody = strBody & "E-mail " & vbTab & vbTab & vbTab & vbTab & Request.form("txtEmail") & "." & vbCrLf 
	strBody = strBody & "Phone Number " & vbTab & vbTab & vbTab & Request.form("txtPhone") & "." & vbCrLf
	strBody = strBody & "Fax Number "& vbTab & vbTab & vbTab & vbTab & Request.form("TxtFax") & "." & vbCrLf 
	strBody = strBody & "Comments " & vbTab & vbTab & vbTab & vbTab & Request.form("taComments") & "." 
	


	
	
	Set lobjMl = server.createObject("CDONTS.newmail")
	if IsObject(lobjMl) = False then			
	end if
	
	lobjMl.From = strFrom
	lobjMl.To=strTo
	lobjMl.Subject = strSubject
	lobjMl.Body=strbody	
	lobjMl.send
	
	Set lobjMl = Nothing
	if Err.number = 0 then
				
	else 
		
	end if
	Response.Redirect("index3.htm")
end if
%>

<html>

<head>
<title>RFQ</title>

<script language= "javascript">
function send()
{
	if(validate())
	{ 
		return false
	}
	document.form1.sendflg.value = "1";
	document.form1.method="post";
	document.form1.action = "rfq.asp";	
	document.form1.submit();
}

function isNull(str)
{
	 bool = false	 
	while(str.charAt(0)==" ")
	{
		str=str.substring(1,str.length);
	}
	if(str.length==0)
	{
		bool = true;
	}
	return bool ;
}

function validate()
{
	val = false;
	if(isNull(document.form1.txtPartNo.value))
	{
		alert("Please enter the part no.");
		document.form1.txtPartNo.focus()
		val = true
	}
	else if(isNull(document.form1.txtFname.value))
	{
		alert("Please enter the First Name");
		document.form1.txtFname.focus()
		val = true
	}	
	else if(isNull(document.form1.txtLname.value))
	{
		alert("Please enter your Last name");
		document.form1.txtLname.focus()
		val = true
	}else if(isNull(document.form1.txtStreet.value))
	{
		alert("Please enter your Street Address");
		document.form1.txtStreet.focus()
		val = true
	}else if(isNull(document.form1.txtCity.value))
	{
		alert("Please enter your City");
		document.form1.txtCity.focus()
		val = true
	}else if(isNull(document.form1.txtState.value))
	{
		alert("Please enter your State");
		document.form1.txtState.focus()
		val = true
	}else if(isNull(document.form1.txtZip.value))
	{
		alert("Please enter your Zip Code");
		document.form1.txtZip.focus()
		val = true
	}else if(isNull(document.form1.txtCountry.value))
	{
		alert("Please enter your Country");
		document.form1.txtCountry.focus()
		val = true
	}else if(isNull(document.form1.txtEmail.value))
	{
		alert("Please enter your Email Id");
		document.form1.txtEmail.focus()
		val = true
	}else if(isNull(document.form1.txtPhone.value))
	{
		alert("Please enter your Phone Number");
		document.form1.txtPhone.focus()
		val = true
	}	else if(! IsEmail("Email Id",document.form1.txtEmail.value))
	{		
		val = true
		document.form1.txtEmail.focus()
	}
	
	return val
}


function IsEmail (msg,str)
{   
	
	
    var i = 1;
    var sLength = str.length;

    // look for @
    while ((i < sLength) && (str.charAt(i) != "@"))
    { i++
    }

    if ((i >= sLength) || (str.charAt(i) != "@")) 
		{
			alert("Please enter valid " + msg);			
			return false;
		}
    else i += 2;

	// look for .
    while ((i < sLength) && (str.charAt(i) != "."))
    { i++
    }

    // there must be at least one character after the .
    if ((i >= sLength - 1) || (str.charAt(i) != "."))
     {  
		alert("Please enter valid " + msg);
		
		return false;
      }
     
     //added on 25/10/2001 for special characters
	for (i=0;i<sLength+1;i++)
	{
	if((str.charAt(i) == "!") || (str.charAt(i) == "#") || (str.charAt(i) == "$")
		|| (str.charAt(i) == "^") || (str.charAt(i) == "&") || (str.charAt(i) == "%")
		|| (str.charAt(i) == "?") || (str.charAt(i) == "*")  || (str.charAt(i) == "<") 
		|| (str.charAt(i) == ">") || (str.charAt(i) == "{") || (str.charAt(i) == "}") 
		|| (str.charAt(i) == "[") || (str.charAt(i) == "]") || (str.charAt(i) == "(") 
		|| (str.charAt(i) == ")") || (str.charAt(i) == "+")|| (str.charAt(i) == "~") 
		|| (str.charAt(i) == "=") || (str.charAt(i) == "/") || (str.charAt(i) == "|")
		|| (str.charAt(i) == "'")|| (str.charAt(i) == '"')|| (str.charAt(i) == ':')
		|| (str.charAt(i) == ';')|| (str.charAt(i) == ',')|| (str.charAt(i) == '`')
		|| (str.charAt(i) == '-'))
		{
			var p = str.charAt(i)
			alert ("Please enter valid " + msg)
		
			return false;
		} 
	}
   //end    
     
    //added on 25/10/2001 for duplicate @ sign
	var count = 0
	for (i=0;i<=sLength+1;i++)
	{
		if(str.charAt(i) == "@")
		count=count+1;
	}

	if(count>1)
	{
		alert ("Please enter valid " + msg)
		
		return false;
	}
	// end
   
    //added
    if(str.charAt(0) == ".")
	{
		alert ("Please enter valid " + msg)
		
		return false;
	}
    //end
     //added on 25/10/2001 for "." check. "." can't come in the end
     if(str.charAt(sLength-1) == ".")
	{
		alert ("Please enter valid " + msg)
		
		return false;
	}
     //end
     //added
     for (i=0;i<sLength+1;i++)
	{
		if((str.charAt(i) == "@" && str.charAt(i+1) == ".")
		||(str.charAt(i+1) == "@" && str.charAt(i) == ".")
		||(str.charAt(i+1) == "." && str.charAt(i) == "."))
			{
				alert ("Please enter valid " + msg)
		
				return false;
			}
	}
     //end
     //added
     if(str.charAt(0) == "@")
	{
		alert ("Please enter valid " + msg)
		
		return false;
	}
     //end
    else return true;
}


</script>

<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body background="images/background.jpg">

<p align="center"><b><font face="Verdana" color="#0000ff"><u>REQUEST FOR QUOTE</u></font></b></p>
<p><font face="Verdana" color="#0000ff"><b>Please Fill in the fields below if
you want to get our catalog or pricing and/or order the parts</b></font></p>
<form method="post" action="rfq.asp" name="form1">
  <table border="0" width="100%">
    <tr>
      <td><font face="Verdana" color="#0000ff">Pricing Detail for Part No.</font></td>
      <td width="523" height="25"><font face="Verdana"><input name="txtPartNo" size="30"></font></td>
    </tr>
    <tr>
      <td width="131" height="25"><font face="Verdana" color="#0000ff">First
        Name</font></td>
      <td width="523" height="25"><font face="Verdana"><input name="txtFname" size="30"></font></td>
    </tr>
    <tr>
      <td width="131" height="25"><font face="Verdana" color="#0000ff">Last Name</font></td>
      <td width="523" height="25"><font face="Verdana"><input name="txtLname" size="30"></font></td>
    </tr>
    <tr>
      <td width="654" height="25" colspan="2"><font face="Verdana" color="#0000ff"><b>How
        would you like to receive the information you requested</b></font></td>
    </tr>
    <tr>
      <td width="654" height="25" colspan="2"><font face="Verdana">&nbsp;<input type="radio" value="Email" checked name="rdChoice">
        <font color="#0000ff">By email<br>
        &nbsp;<input type="radio" value="RegularMail" name="rdChoice"> By
        Regular mail<br>
        &nbsp;<input type="radio" value="Fax" name="rdChoice"> By FAX</font></font></td>
    </tr>
    <tr>
      <td width="131" height="25"><font face="Verdana" color="#0000ff">Street
        Address:</font></td>
      <td width="523" height="25"><font face="Verdana"><input name="txtStreet" size="60"></font></td>
    </tr>
    <tr>
      <td width="131" height="25"><font face="Verdana" color="#0000ff">City:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
      <td width="523" height="25"><font face="Verdana"><input name="txtCity" size="20"></font></td>
    </tr>
    <tr>
      <td width="131" height="24"><font face="Verdana" color="#0000ff">State:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
      <td width="523" height="24"><font face="Verdana"><input name="txtState" size="20"></font></td>
    </tr>
    <tr>
      <td width="131" height="24"><font face="Verdana" color="#0000ff">ZIP
        Code:&nbsp;&nbsp;&nbsp;</font></td>
      <td width="523" height="24"><font face="Verdana"><input name="txtZip" size="20"></font></td>
    </tr>
    <tr>
      <td width="131" height="24"><font face="Verdana" color="#0000ff">Country&nbsp;</font></td>
      <td width="523" height="24"><font face="Verdana"><input name="txtCountry" size="20"></font></td>
    </tr>
    <tr>
      <td width="131" height="24"><font face="Verdana" color="#0000ff">E-mail:</font></td>
      <td width="523" height="24"><font face="Verdana"><input name="txtEmail" size="20"></font></td>
    </tr>
    <tr>
      <td width="131" height="24"><font face="Verdana" color="#0000ff">Phone
        Number:</font></td>
      <td width="523" height="24"><font face="Verdana"><input name="txtPhone" size="20"></font></td>
    </tr>
    <tr>
      <td width="131" height="24"><font face="Verdana" color="#0000ff">FAX
        Number</font></td>
      <td width="523" height="24"><font face="Verdana"><input name="TxtFax" size="20"></font></td>
    </tr>
    <tr>
      <td width="276" height="24"><font face="Verdana" color="#0000ff">Comments</font></td>
      <td width="378" height="24"><font face="Verdana"><textarea name="taComments" rows="3" cols="56"></textarea></font></td>
    </tr>
    <tr>
      <td width="649" height="41" colspan="2">
        <p align="center"><font face="Verdana"><input type="button" value="Send To IOP" style="COLOR: #0000ff" onclick="return send()">&nbsp;
        <input type="reset" value="Reset Form" name="reset" style="COLOR: #0000ff"></font></p>
      </td>
      <input type="hidden" name="sendflg">
    </tr>
  </table>
</form>

</body>

</html>
