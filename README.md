<div align="center">

## Data Encryption \- For Beginners


</div>

### Description

Explains the basic technique used to achieve data encryption and get you on your way to understanding Cryptography in general.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jonathan Roach](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jonathan-roach.md)
**Level**          |Beginner
**User Rating**    |4.6 (51 globes from 11 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jonathan-roach-data-encryption-for-beginners__1-29764/archive/master.zip)





### Source Code

```
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<style>
<!--
 /* Font Definitions */
@font-face
	{font-family:Gulim;
	panose-1:0 0 0 0 0 0 0 0 0 0;
	mso-font-alt:\AD74\B9BC;
	mso-font-charset:129;
	mso-generic-font-family:roman;
	mso-font-format:other;
	mso-font-pitch:fixed;
	mso-font-signature:1 151388160 16 0 524288 0;}
@font-face
	{font-family:Tahoma;
	panose-1:2 11 6 4 3 5 4 4 2 4;
	mso-font-charset:0;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:16792199 0 0 0 65791 0;}
@font-face
	{font-family:"\@Gulim";
	panose-1:0 0 0 0 0 0 0 0 0 0;
	mso-font-charset:129;
	mso-generic-font-family:roman;
	mso-font-format:other;
	mso-font-pitch:fixed;
	mso-font-signature:1 151388160 16 0 524288 0;}
 /* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
@page Section1
	{size:8.5in 11.0in;
	margin:1.0in 1.25in 1.0in 1.25in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
-->
</style>
</head>
<body lang=EN-US style='tab-interval:.5in'>
<div class=Section1>
<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
style='font-size:18.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;
mso-bidi-font-family:"Times New Roman"'>Cryptography Primer<o:p></o:p></span></b></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>By Jonathan Roach<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Hello and welcome to
my primer article on cryptography!<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><i style='mso-bidi-font-style:normal'><span
style='font-size:9.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;mso-bidi-font-family:
"Times New Roman"'>cryp-tog-ra-phy (krip-taw-graph-e)<o:p></o:p></span></i></p>
<p class=MsoNormal><i style='mso-bidi-font-style:normal'><span
style='font-size:9.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;mso-bidi-font-family:
"Times New Roman"'>The process or skill of communicating in or deciphering<o:p></o:p></span></i></p>
<p class=MsoNormal><i style='mso-bidi-font-style:normal'><span
style='font-size:9.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;mso-bidi-font-family:
"Times New Roman"'>secret writings or ciphers.</span></i><span
style='font-size:9.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;mso-bidi-font-family:
"Times New Roman"'><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;
mso-bidi-font-family:"Times New Roman"'>Introduction<o:p></o:p></span></b></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Prying eyes,
espionage, fraud, and theft of personal information.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>These are a few of
the reasons for concealing, masking, shadowing<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>or encrypting
information in order to minimize the chance of that<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>information being
revealed to potentially dangerous or mischievous<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>individuals/organizations.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Okay, okay... maybe
it's not that big of a deal, maybe you just want<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>to feel covert when
you send email to your friends or something.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>The method used to
achieve the above is referred to as &quot;Cryptography&quot;,<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>and this article is
aimed at giving you a basic look into the world of<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>data encryption and
providing you with Visual Basic source code to get<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>you started on your
way.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;
mso-bidi-font-family:"Times New Roman"'>What does Cryptography do?<o:p></o:p></span></b></p>
<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;
mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></b></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Cryptography
conceals or hides data in order to make it un-readable to<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>the average person,
it is used to secure documents and data by mixing<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>or scrambling the
original data into mumbo jumbo basically.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Take this generic
example of encryption, lets say you want to send an<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>email to your
friend, and you don't want anyone else to see the true<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>message you are
going to send, because... it's top secret of course.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Your original
message would look something like this:<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style='mso-tab-count:1'>            </span>&quot;Hey Frank... I got that new
encryption handbook.&quot;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>You would then
perform an encryption routine on the message before<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>sending your email
and the result would look something like this:<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style='mso-tab-count:1'>            </span>&quot;Ì&#8217;&#710;/&#382;&#376;¨&#8220;&#710;áÝ ??|&#8220;&#339;»&#710;³ßÝ</span><span
style='font-size:9.0pt;mso-bidi-font-size:12.0pt;font-family:Gulim'>?</span><span
style='font-size:9.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;mso-bidi-font-family:
"Times New Roman"'>?©§&#8224; &#710;³ß?|&#8220;&#339;?©§&#8224;&quot;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>You get the idea,
the original text is all scrambled and basically not of any<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>use to anyone, or so
it appears...<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>So now you send your
encrypted mail off to Frank, if Frank is unaware of<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>encryption then he
will probably mail you back and say something along<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>the lines of
&quot;What the heck is this stuff you sent me?&quot;.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>However, because you
sent it off to your good buddy Frank and he is<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>using the same
encryption/decryption software that you are, and he is<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>aware of the
code/key needed to reverse your scrambled message all<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>is well and he can
view your message, while any others who may have<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>intercepted it along
the way could not, or at least had a heck of a time<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>in doing so...<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>So there you have
it, Cryptography scrambles/transforms data.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;
mso-bidi-font-family:"Times New Roman"'>Cryptography Overview<o:p></o:p></span></b></p>
<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;
mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></b></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Cryptography
requires an encryption algorithm and a key; in it's basic form<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>that is.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Due to the nature of
this article I will not go into great detail on the many<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>methods of
encryption and key methods in use today but I will provide<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>you with the base
foundation for encryption/decryption.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>An encryption
algorithm is simply the engine or code that handles all of<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>the processes that
transform the original text (plaintext) into encoded text<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>(ciphertext).<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>This engine simply
performs mathematical and/or logical operations on the<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>plaintext to
transform it into the ciphertext and vice versa.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>The key as it's name
implies is just that, it is the key (code) that allows the<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>algorithm to
encrypt/decrypt data.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;
mso-bidi-font-family:"Times New Roman"'>Common Cryptography Algorithms<o:p></o:p></span></b></p>
<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;
mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></b></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>There are many
different algorithms for encrypting/decrypting data in use<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>around the world
today, some of them are very complex and others are<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>more simplistic,
however they all serve the same purpose.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Below is a short
listing on some of the different cryptography algorithms;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>DES - United States
Data Encryption Standard<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>3DES - The above,
encoded 3 times<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>RSA - Rivest, Shamir
and Adleman<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>GOST - Developed by
scientists of the former Soviet Union<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>IDEA - A component
of PGP (Pretty Good Privacy)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>There are many more,
but the above should be a starting point for you<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>to seek out more
info on the net.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;
mso-bidi-font-family:"Times New Roman"'>The One-Time Pad<o:p></o:p></span></b></p>
<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;
mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></b></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>The one-time pad is
one of the simplest encryption algorithms, it involves<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>a key being used
which is the same length as the plaintext and then using<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>simple math on the
plaintext via the key, the math could be multiplication<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>or exclusive-or
(XOR) for example;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Dim plainText As
String<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Dim cipherKey As
String<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Dim Counter As
Integer<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Dim Char As String<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Dim keyChar As
String<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Dim cipherText As
String<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Private Sub Crypt()<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>plainText =
&quot;CovertText&quot;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>cipherKey =
&quot;password42&quot;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>MsgBox &#8220;Before: &#8220;
&amp; plainText<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>'Encrypt it<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>For Counter = 1 To
Len(plainText)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style="mso-spacerun: yes">    </span>Char = Asc(Mid(plainText, Counter, 1))<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style="mso-spacerun: yes">    </span>keyChar = Asc(Mid(cipherKey, Counter, 1))<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style="mso-spacerun: yes">    </span>cipherText = cipherText &amp; Chr(Char Xor
keyChar)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Next Counter<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>MsgBox &#8220;After: &#8220;
&amp; cipherText<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>plainText =
&quot;&quot;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>'Decrypt it<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>For Counter = 1 To
Len(cipherText)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style="mso-spacerun: yes">    </span>Char = Asc(Mid(cipherText, Counter, 1))<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style="mso-spacerun: yes">    </span>keyChar = Asc(Mid(cipherKey, Counter, 1))<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style="mso-spacerun: yes">    </span>plainText = plainText &amp; Chr(Char Xor
keyChar)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Next Counter<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>MsgBox &#8220;Back to
original: &#8220; &amp; plainText<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>cipherText =
&quot;&quot;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>End Sub<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Just copy the above
code and paste it into a new project, then add a<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>command button and
in its click event put a call to the Crypt() sub.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>As follows,<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Private Sub
Command1_Click()<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Crypt<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>End Sub<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Run the project and
click the button to see it in action.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>There is a problem
with the above encryption algorithm though, first<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>off if you want to
encrypt something that is large in size the key size<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>would also be very
large.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>For example if you
wanted to encrypt a string that is 50 characters in<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>length then your key
would also have to be 50 characters in length,<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>kind of a pain for
our good friend Frank to have to enter a 50 character<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>decoding key for a
simple message.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>This problem can be
overcome with our next topic, which deals with<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>the key length
problem by using a repeating key.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;
mso-bidi-font-family:"Times New Roman"'>Repeating Key Algorithm - Viginere
Cipher<o:p></o:p></span></b></p>
<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma;
mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></b></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>This type of
encryption algorithm deals with a key that repeats during<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>the
encryption/decryption process, for example the algorithm above<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>uses a character by
character algorithm, it performs math operations<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>on each character in
the plaintext and key until the length of the key<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>and plaintext is
reached - because the key and plaintext are the same<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>length.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>However with a
repeating key, our key can be any length we choose<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>as when the end of
the key is reached in our algorithm we simple start<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>again at the begin
of the key until our plaintext encryption is completed.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Take this generic
example, if our plaintext is &quot;I have top secret codes&quot;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>and our key is
&quot;Pass&quot;; obviously the key is shorter than our plaintext.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Thus we are in our
loop to encrypt our plaintext and this is how it looks:<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style='mso-tab-count:1'>            </span>plaintext char = I<span
style='mso-tab-count:1'>            </span><span style='mso-tab-count:1'>            </span>key
char = P<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style='mso-tab-count:1'>            </span>plaintext char = Space<span
style='mso-tab-count:1'>     </span>key char = a<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style='mso-tab-count:1'>            </span>plaintext char = h<span
style='mso-tab-count:1'>            </span>key char = s<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style='mso-tab-count:1'>            </span>plaintext char = a<span
style='mso-tab-count:1'>            </span>key char = s<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style='mso-tab-count:1'>            </span>plaintext char = v<span
style='mso-tab-count:1'>            </span>key char = P<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style='mso-tab-count:1'>            </span>plaintext char = e<span
style='mso-tab-count:1'>            </span>key char = a<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style='mso-tab-count:1'>            </span>plaintext char = Space<span
style='mso-tab-count:1'>     </span>key char = s<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>You get the idea?
The key just repeats until the length of the plaintext<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>is reached. This
method is much more practical and flexible for key names<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>anyway.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Below is a sample
algorithm that uses a repeating key.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Private Sub Crypt()<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>cipherKey =
&quot;pw201&quot;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>plainText =
&quot;Top-Secret Message from Roach&quot;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>cipherText =
&quot;&quot;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>KeyIndex = 1<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>MsgBox &quot;Before:
&quot; &amp; plainText, 0, &quot;Before Encryption&quot;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>For Counter = 1 To
Len(plainText)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style="mso-spacerun: yes">    </span>Char = Asc(Mid(plainText, Counter, 1))<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style="mso-spacerun: yes">    </span>keyChar = Asc(Mid(cipherKey, KeyIndex, 1))<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style="mso-spacerun: yes">    </span>cipherText = cipherText &amp; Chr(Char Xor
keyChar)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style="mso-spacerun: yes">    </span>KeyIndex = KeyIndex + 1<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><span
style="mso-spacerun: yes">    </span>If KeyIndex &gt; Len(cipherKey) Then
KeyIndex = 1<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Next Counter<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>MsgBox &quot;After
Encryption: &quot; &amp; cipherText, 0, &quot;Original:&quot; &amp; plainText<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>End Sub<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Once again you can
paste this into a new project and call the Crypt()<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>sub from a buttons
click event to try it out. To reverse the encryption<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>just run the
encrypted text back through the counter loop in place of<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>the plaintext.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Thank you for
sticking with me through this brief article on the subject,<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>I hope that you
gained a little knowledge about encryption from this.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Note: I do not claim
that you shall become an encryption expert or<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>that any of the
methods described in this article are bomb proof,<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>crack proof, water
proof... whatever, I merely wanted to share the<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>basic knowledge of
the subject.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Regards,<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:9.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Jonathan Roach</span></p>
</div>
</body>
</html>
```

