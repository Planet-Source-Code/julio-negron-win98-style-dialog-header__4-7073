<div align="center">

## Win98\-style Dialog Header


</div>

### Description

Create a Windows 98-Style dialog box header. Ever wanted to recreate that cool gradient bar that starts at one color and ends up another-without images? Now You Can! Easily Modifed to include icons or any type of images. BONUS: The "CreateColorTable()" function used in the script creates a 216 color table without using images! This script is totally independant of any other files or images, simply cut, paste and modifiy at your own will.
 
### More Info
 
Inputs are controlled by the form and explained within the script.

Some colors will not display properly on 256 color environments. The script does not compensate for this variable, it will display any and all colors it needs to accomplish the effect.

The code returns a basic table with cells that are opened just wide enough so you can see the background color, optional text is inserted in the first cell. Easily modified to include icons or images.

This code can output many cells depending on what colors are chosen for the effect, it may slightly hinder web site performance.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Julio Negron](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/julio-negron.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/julio-negron-win98-style-dialog-header__4-7073/archive/master.zip)

### API Declarations

This code may be used freely either privately or commercially. It may NOT be sold through any medium or market for profit. Please notify me of any positive modifications: jnegron@volt.com.


### Source Code

```
<%
if Request.Form > "" then
dim debug
if Request.Form("debugOn") = "on" then debug = CBool("true")
'****************************************************************************
'This function creates a specified number of nonbreaking spaces
'****************************************************************************
Private Function nbsp(iNumber)
	mySpace = "&nbsp;"
	do until count = iNumber
		mySpace = mySpace & mySpace
		count = count + 1
	loop
	nbsp = mySpace
end function
'****************************************************************************
'This function converts a three digit number into a two digit hex number and
'ensures that the result is always a two digit hex number.
'****************************************************************************
Private Function smartHex(strRGB)
if debug then
	Response.write "<font face=Verdana size='+1'><table border>"
	Response.Write "<tr><td>smartHex Function (begin):</td><td><b>" & strRGB & " </b></td><td>TYPE:</td><td><b>" & TypeName(strRGB) & "</b></td></tr>"
end if
	if typeName(strRGB) = "Double" or Len(strRGB) = 3 then
		strRGB = hex(strRGB)
	end if
if debug then
	Response.Write "<tr><td>strRGB after Hex Conversion:</td><td><b>" & strRGB & " </b></td><td>TYPE:</td><td><b>" & TypeName(strRGB) & "</b></td></tr>"
End If
	if len(strRGB) = 1 then
		strRGB = 0 & strRGB
	end if
If debug Then
	Response.Write "<td>strRGB after Conditional Statement:</td><td><b>" & strRGB & " </b></td><td>TYPE:</td><td><b>" & TypeName(strRGB) & "</b></td></tr>"
End If
	smartHex = strRGB
If debug Then
	Response.Write "<tr><td>smartHex Function (end):</td><td><b>" & smartHex & " </b></td><td>TYPE:</td><td><b>" & TypeName(smartHex) & "</b></td></tr>"
	Response.Write "</table></font><br>"
End IF
End function
'****************************************************************************
'This function takes a three digit integer and increments it
'or decreases it by 1 depending on the colorEnd variable. If the
'colorEnd and the colorStart variables are equal, it will not change
'the resulting number.
'****************************************************************************
Private Function CheckNum(intNumber,colorStart,colorEnd)
If debug Then
	Response.write "<font face=Georgia size='+1'><table border>"
	Response.Write "<tr><td>CheckNum Function(begin):</td><td><b>" & intNumber & "</b></td><td>TYPE:</td><td><b>" & TypeName(intNumber) & "</b></td></tr>"
End If
if intNumber <> colorEnd then
	if colorEnd > colorStart then
		if intNumber = "255" then
			intNumber = intNumber
		else
			intNumber = intNumber + 1
		end if
	else if colorEnd < colorStart then
		if intNumber = "0" then
			intNumber = intNumber
		else
			intNumber = intNumber - 1
		end if
	else if colorEnd = colorStart then
		intNumber = intNumber
	end if
	end if
	end if
If debug Then
	Response.Write "<tr><td>CheckNum Function(After 'if...then'):</td><td><b>" & intNumber & "</b></td><td>TYPE:</td><td><b>" & TypeName(intNumber) & "</b></td></tr>"
End If
else
	intNumber = colorEnd
end if
Select Case Len(intNumber)
	case 1
	intNumber = "00" & intNumber
	case 2
	intNumber = "0" & intNumber
End Select
If debug Then
	Response.Write "<tr><td>CheckNum Function(end):</td><td><b>" & intNumber & "</b></td><td>TYPE:</td><td><b>" & TypeName(intNumber) & "</b></td></tr>"
	Response.Write "</table></font><br>"
End If
CheckNum = intNumber
end function
'****************************************************************************
'This function creates a banner with a specified height and width who's
'background color starts as one color and ends up as another at the end. Optional
'text can be added for use as a header.
'****************************************************************************
Public Function StartColorArray(strColor1,strColor2,iHeight,iWidth,strText,strTextColor,iTextSize)
	dim strColor, count
	dim intNum1a, intNum1b, intNum1c, intNum2a, intNum2b, intNum2c
	dim color1a, color1b, color1c, color2a, color2b, color2c
	count	 = 0
	If debug Then Response.Write "count: " & count & "<br>"
	intNum1a = Int("&H" & Left(strColor1,2))
	If debug Then Response.Write "<b>intNum1a: </b>" & intNum1a & nbsp(2)
	intNum1b = Int("&H" & Mid(strColor1,3,2))
	If debug Then Response.Write "<b>intNum1b: </b>" & intNum1b & nbsp(2)
	intNum1c = Int("&H" & Right(strColor1,2))
	If debug Then Response.Write "<b>intNum1c: </b>" & intNum1c & "<br>"
	intNum2a = Int("&H" & Left(strColor2,2))
	If debug Then Response.Write "<b>intNum2a: </b>" & intNum2a & nbsp(2)
	intNum2b = Int("&H" & Mid(strColor2,3,2))
	If debug Then Response.Write "<b>intNum2b: </b>" & intNum2b & nbsp(2)
	intNum2c = Int("&H" & Right(strColor2,2))
	If debug Then Response.Write "<b>intNum2c: </b>" & intNum2c & "<br>"
	color1A	 = intNum1a
	color1B	 = intNum1b
	color1C	 = intNum1c
	color2A	 = intNum2a
	color2B	 = intNum2b
	color2C	 = intNum2c
	iTextWidth = len(strText) * 12
if not debug then
	Response.Write "<table BORDER='0' CELLSPACING='0' CELLPADDING='0' "
		if strText = "" then
			Response.Write "height='" & iHeight & "' "
		else
			iHeight = iTextSize * 10 - 10
			Response.Write "height='" & iHeight & "' "
		end if
	Response.Write "width='" & iWidth & "'><tr>" _
	       & "<td bgColor='#" & strColor1 & "' width='" & iTextWidth & "'>"
		if strText > "" then
			 Response.Write "<font color='" & strTextColor & "' size='" & iTextSize & "'>" _
	     & "&nbsp;<STRONG>" & strText & "</STRONG></font>"
		end if
	Response.Write "</td>"
End If
Do until strColor = strColor2
	If debug Then Response.Write "<font color=blue size=+1>Start Loop " & count + 1 & "</font><br>"
	count = count + 1
	If debug Then
		Response.Write "color2A: " & color2A & "&nbsp;&nbsp;Type: " & TypeName(color2A) & "&nbsp;&nbsp;"
		Response.Write "color2B: " & color2B & "&nbsp;&nbsp;Type: " & TypeName(color2B) & "&nbsp;&nbsp;"
		Response.Write "color2C: " & color2C & "&nbsp;&nbsp;Type: " & TypeName(color2C) & "&nbsp;&nbsp;"
	End If
	intNum1a = CheckNum(intNum1a,color1A,color2A)
	intNum1b = CheckNum(intNum1b,color1B,color2B)
	intNum1c = CheckNum(intNum1c,color1C,color2C)
	If debug Then
		Response.Write "<b>intNum1a: </b>" & intNum1a & nbsp(2)
		Response.Write "<b>intNum1b: </b>" & intNum1b & nbsp(2)
		Response.Write "<b>intNum1c: </b>" & intNum1c & "<br>"
	End If
	strColor = smartHex(intNum1a) & smartHex(intNum1b) & smartHex(intNum1c)
	intNum1a = int("&H" & intNum1a)
	intNum1b = int("&H" & intNum1b)
	intNum1c = int("&H" & intNum1c)
	If debug Then
		Response.Write "<b>strColor: </b>" & strColor & "<br>"
	Else
		Response.Write "<td width='1' bgcolor='#" & strColor & "'><br></td>"
	End IF
	If len(strColor) > 6 Then
		Response.Write "<font color=red size=+1><b>Error:</b> Hex Number has <b>surpassed</b> the 6 digit limit.</font>"
		exit Do
	else if len(strColor) < 6 Then
		Response.Write "<font color=red size=+1><b>Error:</b> Hex Number is <b>less</b> than the 6 digits.</font>"
		exit Do
	End If
	End If
	If debug Then
		Response.Write "strColor1: <font color=red><b>" & strColor1 & "</b></font>" & nbsp(4) _
				 & "strColor2: <font color=red><b>" & strColor2 & "</b></font>" & "<br>"
		Response.Write "<font color=blue size=+1>End Loop</font><br><br>"
	End If
loop
if not debug then
	if strText > "" then
		Response.Write "<td width='20%' bgcolor='#" & strColor & "'><br></td>"
	end if
	Response.Write "</tr></table>"
end if
end function
end if
'*************************************************************************
'This function creates a color table with 216 colors, no images needed. :0)
'NOTE: this function has only been tested in a 16 bit color display.
'*************************************************************************
Function CreateColorTable()
dim arColor(216)
arColor(0) = "00FF00"
arColor(1) = "00FF33"
arColor(2) = "00FF66"
arColor(3) = "00FF99"
arColor(4) = "00FFCC"
arColor(5) = "00FFFF"
arColor(6) = "33FF00"
arColor(7) = "33FF33"
arColor(8) = "33FF66"
arColor(9) = "33FF99"
arColor(10) = "33FFCC"
arColor(11) = "33FFFF"
arColor(12) = "66FF00"
arColor(13) = "66FF33"
arColor(14) = "66FF66"
arColor(15) = "66FF99"
arColor(16) = "66FFCC"
arColor(17) = "66FFFF"
arColor(18) = "99FF00"
arColor(19) = "99FF33"
arColor(20) = "99FF66"
arColor(21) = "99FF99"
arColor(22) = "99FFCC"
arColor(23) = "99FFFF"
arColor(24) = "CCFF00"
arColor(25) = "CCFF33"
arColor(26) = "CCFF66"
arColor(27) = "CCFF99"
arColor(28) = "CCFFCC"
arColor(29) = "CCFFFF"
arColor(30) = "FFFF00"
arColor(31) = "FFFF33"
arColor(32) = "FFFF66"
arColor(33) = "FFFF99"
arColor(34) = "FFFFCC"
arColor(35) = "FFFFFF"
arColor(36) = "00CC00"
arColor(37) = "00CC33"
arColor(38) = "00CC66"
arColor(39) = "00CC99"
arColor(40) = "00CCCC"
arColor(41) = "00CCFF"
arColor(42) = "33CC00"
arColor(43) = "33CC33"
arColor(44) = "33CC66"
arColor(45) = "33CC99"
arColor(46) = "33CCCC"
arColor(47) = "33CCFF"
arColor(48) = "66CC00"
arColor(49) = "66CC33"
arColor(50) = "66CC66"
arColor(51) = "66CC99"
arColor(52) = "66CCCC"
arColor(53) = "66CCFF"
arColor(54) = "99CC00"
arColor(55) = "99CC33"
arColor(56) = "99CC66"
arColor(57) = "99CC99"
arColor(58) = "99CCCC"
arColor(59) = "99CCFF"
arColor(60) = "CCCC00"
arColor(61) = "CCCC33"
arColor(62) = "CCCC66"
arColor(63) = "CCCC99"
arColor(64) = "CCCCCC"
arColor(65) = "CCCCFF"
arColor(66) = "FFCC00"
arColor(67) = "FFCC33"
arColor(68) = "FFCC66"
arColor(69) = "FFCC99"
arColor(70) = "FFCCCC"
arColor(71) = "FFCCFF"
arColor(72) = "009900"
arColor(73) = "009933"
arColor(74) = "009966"
arColor(75) = "009999"
arColor(76) = "0099CC"
arColor(77) = "0099FF"
arColor(78) = "339900"
arColor(79) = "339933"
arColor(80) = "339966"
arColor(81) = "339999"
arColor(82) = "3399CC"
arColor(83) = "3399FF"
arColor(84) = "669900"
arColor(85) = "669933"
arColor(86) = "669966"
arColor(87) = "669999"
arColor(88) = "6699CC"
arColor(89) = "6699FF"
arColor(90) = "999900"
arColor(91) = "999933"
arColor(92) = "999966"
arColor(93) = "999999"
arColor(94) = "9999CC"
arColor(95) = "9999FF"
arColor(96) = "CC9900"
arColor(97) = "CC9933"
arColor(98) = "CC9966"
arColor(99) = "CC9999"
arColor(100) = "CC99CC"
arColor(101) = "CC99FF"
arColor(102) = "FF9900"
arColor(103) = "FF9933"
arColor(104) = "FF9966"
arColor(105) = "FF9999"
arColor(106) = "FF99CC"
arColor(107) = "FF99FF"
arColor(108) = "006600"
arColor(109) = "006633"
arColor(110) = "006666"
arColor(111) = "006699"
arColor(112) = "0066CC"
arColor(113) = "0066FF"
arColor(114) = "336600"
arColor(115) = "336633"
arColor(116) = "336666"
arColor(117) = "336699"
arColor(118) = "3366CC"
arColor(119) = "3366FF"
arColor(120) = "666600"
arColor(121) = "666633"
arColor(122) = "666666"
arColor(123) = "666699"
arColor(124) = "6666CC"
arColor(125) = "6666FF"
arColor(126) = "996600"
arColor(127) = "996633"
arColor(128) = "996666"
arColor(129) = "996699"
arColor(130) = "9966CC"
arColor(131) = "9966FF"
arColor(132) = "CC6600"
arColor(133) = "CC6633"
arColor(134) = "CC6666"
arColor(135) = "CC6699"
arColor(136) = "CC66CC"
arColor(137) = "CC66FF"
arColor(138) = "FF6600"
arColor(139) = "FF6633"
arColor(140) = "FF6666"
arColor(141) = "FF6699"
arColor(142) = "FF66CC"
arColor(143) = "FF66FF"
arColor(144) = "003300"
arColor(145) = "003333"
arColor(146) = "003366"
arColor(147) = "003399"
arColor(148) = "0033CC"
arColor(149) = "0033FF"
arColor(150) = "333300"
arColor(151) = "333333"
arColor(152) = "333366"
arColor(153) = "333399"
arColor(154) = "3333CC"
arColor(155) = "3333FF"
arColor(156) = "663300"
arColor(157) = "663333"
arColor(158) = "663366"
arColor(159) = "663399"
arColor(160) = "6633CC"
arColor(161) = "6633FF"
arColor(162) = "993300"
arColor(163) = "993333"
arColor(164) = "993366"
arColor(165) = "993399"
arColor(166) = "9933CC"
arColor(167) = "9933FF"
arColor(168) = "CC3300"
arColor(169) = "CC3333"
arColor(170) = "CC3366"
arColor(171) = "CC3399"
arColor(172) = "CC33CC"
arColor(173) = "CC33FF"
arColor(174) = "FF3300"
arColor(175) = "FF3333"
arColor(176) = "FF3366"
arColor(177) = "FF3399"
arColor(178) = "FF33CC"
arColor(179) = "FF33FF"
arColor(180) = "000000"
arColor(181) = "000033"
arColor(182) = "000066"
arColor(183) = "000099"
arColor(184) = "0000CC"
arColor(185) = "0000FF"
arColor(186) = "330000"
arColor(187) = "330033"
arColor(188) = "330066"
arColor(189) = "330099"
arColor(190) = "3300CC"
arColor(191) = "3300FF"
arColor(192) = "660000"
arColor(193) = "660033"
arColor(194) = "660066"
arColor(195) = "660099"
arColor(196) = "6600CC"
arColor(197) = "6600FF"
arColor(198) = "990000"
arColor(199) = "990033"
arColor(200) = "990066"
arColor(201) = "990099"
arColor(202) = "9900CC"
arColor(203) = "9900FF"
arColor(204) = "CC0000"
arColor(205) = "CC0033"
arColor(206) = "CC0066"
arColor(207) = "CC0099"
arColor(208) = "CC00CC"
arColor(209) = "CC00FF"
arColor(210) = "FF0000"
arColor(211) = "FF0033"
arColor(212) = "FF0066"
arColor(213) = "FF0099"
arColor(214) = "FF00CC"
arColor(215) = "FF00FF"
Response.Write "<table cellpadding=0 cellspacing=1><tr>"
count = 0
Do Until count = 216
	Response.Write "<td bgcolor='#" & arColor(count) & "'>" _
				 & "<a href=javascript:showColor('" & arColor(count) & "') " _
				 & "onMouseOver=javascript:showColorView('" & arColor(count) & "')>" _
				 & "&nbsp;&nbsp;&nbsp;</a></td>"
	count = count + 1
	select case count
		Case 36
			Response.Write "</tr><tr>"
		Case 72
			Response.Write "</tr><tr>"
		Case 108
			Response.Write "</tr><tr>"
		Case 144
			Response.Write "</tr><tr>"
		Case 180
			Response.Write "</tr><tr>"
		Case 216
			Response.Write "</tr></table>"
	End Select
Loop
End Function
%>
<html><head>
<title>Gradient Bar Function <% if debug then response.write "- DEBUG Mode" %></title>
<style>
a {text-decoration:none}
body,td,input {font-family: verdana; font-size:8pt}
</style>
<SCRIPT LANGUAGE="JavaScript">
function showColorView(val) {
document.colorform.frmColorView.value = "Selected Color: #" + val;
document.colorform.frmColorShow.style.backgroundColor = val;
}
function enableText() {
isTextColorEnabled = document.colorform.frmTextColor.disabled
isTextSizeEnabled = document.colorform.frmTextSize.disabled
isTextColorView = document.colorform.colorView3.disabled
isTextChBoxEnabled = document.colorform.checkbox3.disabled
isTextEnabled = document.colorform.frmText.disabled
if (!isTextColorEnabled) {
	document.colorform.frmTextColor.disabled = true;
	document.colorform.frmTextColor.style.backgroundColor = "DDDDDD"
	}else{
	document.colorform.frmTextColor.disabled = false;
	document.colorform.frmTextColor.style.backgroundColor = "FFFFFF"
	}
if (!isTextEnabled) {
	document.colorform.frmText.disabled = true;
	document.colorform.frmText.style.backgroundColor = "DDDDDD"
	}else{
	document.colorform.frmText.disabled = false;
	document.colorform.frmText.style.backgroundColor = "FFFFFF"
	}
if (!isTextSizeEnabled) {
	document.colorform.frmTextSize.disabled = true;
	document.colorform.frmTextSize.style.backgroundColor = "DDDDDD"
	}else{
	document.colorform.frmTextSize.disabled = false;
	document.colorform.frmTextSize.style.backgroundColor = "FFFFFF"
	}
if (!isTextColorView) {
	colorViewVal = document.colorform.colorView3.style.backgroundColor
	document.colorform.colorView3.disabled = true;
	document.colorform.colorView3.style.backgroundColor = "DDDDDD"
	}else{
	document.colorform.colorView3.disabled = false;
	document.colorform.colorView3.style.backgroundColor = colorViewVal
	}
if (!isTextChBoxEnabled) {
	document.colorform.checkbox3.disabled = true;
	document.colorform.checkbox3.style.backgroundColor = "DDDDDD"
	}else{
	document.colorform.checkbox3.disabled = false;
	document.colorform.checkbox3.style.backgroundColor = "FFFFFF"
	}
}
function setColor(){
	document.colorform.colorView1.style.backgroundColor = document.colorform.frmColor1.value;
	document.colorform.colorView2.style.backgroundColor = document.colorform.frmColor2.value;
	document.colorform.colorView3.style.backgroundColor = document.colorform.frmTextColor.value;
}
function showColor(val) {
	var check1,check2
	check1 = document.colorform.checkbox1;
	check2 = document.colorform.checkbox2;
	if (document.colorform.checkbox3.checked){
		document.colorform.frmTextColor.value = val;
		document.colorform.colorView3.style.backgroundColor = val;
	}else {
		if (!check1.checked && !check2.checked) {
			window.alert("You must check either the 'Start Color' or 'End Color'.");
		}else {
			if (document.colorform.checkbox1.checked) {
				document.colorform.frmColor1.value = val;
				document.colorform.colorView1.style.backgroundColor = val;
			}else {
				document.colorform.frmColor2.value = val;
				document.colorform.colorView2.style.backgroundColor = val;
				}
			}
		}
	}
function check1() {
	document.colorform.checkbox3.checked = false;
	if (document.colorform.checkbox1.checked) {
		document.colorform.checkbox2.checked = false;
	}else{
		document.colorform.checkbox1.checked = false;
		document.colorform.checkbox2.checked = true;
	}
}
function check2() {
	document.colorform.checkbox3.checked = false;
	if (document.colorform.checkbox2.checked) {
		document.colorform.checkbox1.checked = false;
	}else{
		document.colorform.checkbox2.checked = false;
		document.colorform.checkbox1.checked = true;
	}
}
function check3() {
	document.colorform.checkbox1.checked = false;
	document.colorform.checkbox2.checked = false;
	}
</script>
</head>
<body onload="setColor()">
<center>
<form method=post id=colorform name=colorform>
<table bgcolor="black" cellpadding=1 cellspacing=0>
	<tr>
		<td colspan=2>
			<%= CreateColorTable() %>
		</td>
	</tr>
	<tr>
		<td>
			<input type=text name=frmColorView readonly style="border-width:0px; width:100%; font-weight:bold; color:#FFFF00; background-color:navy; text-align:center">
		</td>
		<td>
			<input type=text name=frmColorShow readonly style="border-width:0px; width:100%; font-weight:bold; color:#FFFF00; background-color:navy; text-align:center">
		</td>
	</tr>
</table>
Click on a cell to set color for selected item.
<table cellpadding=6>
<tr>
<td align=left>
	<INPUT type="text" id=colorView1 name=colorView1 readonly style="background-color:<%=Request.Form("color1ViewVal")%>">
</td><td>
	<INPUT type="checkbox" id=checkbox1 name=checkbox1 onClick="check1()">Start Color:
</td><td>
	<input type=text name=frmColor1 value="<%=Request.Form("frmColor1")%>">
</td>
</tr>
<tr>
<td align=left>
	<INPUT type="text" id=colorView2 name=colorView2 readonly>
</td><td>
	<INPUT type="checkbox" id=checkbox2 name=checkbox2 onClick="check2()">End Color:&nbsp;
</td><td>
	<input type=text name=frmColor2 value="<%=Request.Form("frmColor2")%>">
</td>
</tr>
<tr>
<td colspan=3>
	Height: <input type=text name=frmHeight size=5 maxlength=5 value="<%=Request.Form("frmHeight")%>">&nbsp;
	Width: <input type=text name=frmWidth size=5 maxlength=5 value="<%=Request.Form("frmWidth")%>">&nbsp;
</td>
</tr>
<tr>
<td colspan=3>
<hr>
<table width=100%><tr>
<td><INPUT Checked type="checkbox" id=frmUseText name=frmUseText onClick="enableText()">
Enable Text</td>
<td align=left>
	<INPUT type="checkbox" id=checkbox3 name=checkbox3 onClick="check3()">
	Text Color:&nbsp;
	<input type=text name=frmTextColor size=7 value="<%=Request.Form("frmTextColor")%>">
	<INPUT type="text" id=colorView3 name=colorView3 readonly>
</td>
</tr>
<tr>
<td>
<INPUT type="checkbox" id="debugOn" name="debugOn" <%if debug then Response.Write "checked"%>>Debug
</td>
<td>
	Text: <input type=text name=frmText value="<%=Request.Form("frmText")%>">&nbsp;
	Text Size(1-9): <input type=text name=frmTextSize size=4 maxlength=2 value="<%=Request.Form("frmTextSize")%>">
</td>
</tr></table>
</td>
</tr></table>
<input type=submit value=submit id=submit1 name=submit1>
</form>
<%
if Request.Form > "" then
myColor1 = UCase(Request.Form("frmColor1"))
myColor2 = UCase(Request.Form("frmColor2"))
myHeight = Request.Form("frmHeight")
myWidth = Request.Form("frmWidth")
myText = Request.Form("frmText")
myTextColor = Request.Form("frmTextColor")
myTextSize = Request.Form("frmTextSize")
Response.Write "<font size='+1'><STRONG>Output"
if debug then Response.Write " - Debug Mode"
Response.Write ":</STRONG></font><br>"
Response.Write StartColorArray(myColor1,myColor2,myHeight,myWidth,myText,myTextColor,myTextSize)
set myColor1 = nothing
set myColor2 = nothing
set myHeight = nothing
set myWidth = nothing
set myText = nothing
set myTextColor = nothing
set myTextSize = nothing
end if
%>
</center>
</body></html>
```

