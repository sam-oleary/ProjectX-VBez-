
<!DOCTYPE html>

<html>
<head>
<title>Wilde XL Solutions - Excel/VBA solutions for home and business!</title>

<link rel="stylesheet" style="text/css" href="main.css" />
<link href='http://fonts.googleapis.com/css?family=Strait' rel='stylesheet' type='text/css'>

</head>

<body>

<div id="menuBar">
<ul>
	<li><a href="../index.html">Home</a></li>
	<li>|</li>
	<li><a href="../about.html">About</a></li>
	<li>|</li>
	<li><a href="../vba-resources.html">VBA Resources</a></li>
	<li>|</li>
	<li><a href="../contact.html">Contact</a></li>
</ul>
<!--twitter-->
<center>
<a href="https://twitter.com/WXLSinfo" class="twitter-follow-button" data-show-count="false" data-size="large" data-dnt="true">Follow @WXLSinfo</a>
<script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0],p=/^http:/.test(d.location)?'http':'https';if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src=p+'://platform.twitter.com/widgets.js';fjs.parentNode.insertBefore(js,fjs);}}(document, 'script', 'twitter-wjs');</script>
</center>
</div>
<!------------------------------------------------------------->

<div id="containerDiv" style="height: auto;">
<center>
<!--AdSense Banner-->
<script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
<!-- HeaderAd -->
<ins class="adsbygoogle"
     style="display:inline-block;width:728px;height:90px"
     data-ad-client="ca-pub-2293004001425348"
     data-ad-slot="1102359317"></ins>
<script>
(adsbygoogle = window.adsbygoogle || []).push({});
</script>
<!--AdSense Banner-->
<br />
</center>

<?php
$strArgs = $_POST['strArgs'];

// Change spaces to non-breaking space in order to preserve formatting in HTML:
$strArgs = str_replace(' ', '&nbsp;', $strArgs);

$strArr = explode("\n", $strArgs);

// keywords array:
$keywords = Array("Alias", "And", "As", "Boolean", "ByRef", "ByVal", "Call", "Case", "CBool", "CDate", "CDbl", "CInt",
	"Class", "CLng", "CObj", "Const", "CStr", "Date", "Debug", "Declare", "Dim", "Do", "Double", "Each", "Else",
	"ElseIf", "End", "EndIf", "Enum", "Erase", "Error", "Exit", "False", "For", "Friend", "Function", "Get", "Global",
	"GoTo", "If", "Implements", "In", "Inherits", "Integer", "Interface", "Is", "Let", "Lib", "Like", "Long",
	"Loop", "Me", "Mod", "Module", "New", "Next", "Not", "Nothing", "Object", "Of", "On", "Option", "Optional",
	"Or", "OrElse", "Print", "Private", "Property", "Protected", "Public", "ReadOnly", "ReDim", "Resume", "Return",
	"Select", "Set", "Shared", "Static", "Step", "Stop", "String", "Sub", "Then", "To", "True", "TypeOf", "Variant",
	"Wend", "When", "While", "With", "WriteOnly");
	
foreach($strArr as $line){
$newLine = "";
// Change comments
	preg_match_all("/[\'].*/", $line, $matches);
if(strlen($matches[0][0]) != ""){	
	$newLine = str_replace($matches[0][0], "<span style=\"color: green;\">".$matches[0][0]."</span>", $line);
} else {
$newLine = $line;
}

$wordArr = explode('&nbsp;', $newLine);

	foreach($wordArr as $word){
		if(substr($word, 0, 28) == "<span style=\"color: green;\">") break;
		if(in_array(trim($word), $keywords)){
		$newLineTwo = preg_replace("/((?<=&nbsp;)$word|$word(?=&nbsp;))/", "<span style=\"color: navy;\">".$word."</span>", $newLine);
		$newLine = $newLineTwo;
		}
	}
		$strArgs = str_replace($line, $newLine, $strArgs);
}

// Change PHP line break to <br /> to preserve formatting in HTML:
$strArgs = nl2br($strArgs);

// Create output page
?>
<div>

	<div style="background-color: #eee; font-family: 'Courier New', Courier, monospace; font-size: 12px;">
		<?PHP echo "<strong>".$strArgs."</strong>"; ?></div>
		<br /><br />
	<pre><textarea cols=150 rows=40><?PHP echo "<span style=\"font-family: 'Courier New', Courier, monospace; font-size: 12px;\">".$strArgs."</span>"; ?></textarea>
		</pre><br /><br />


</div>

<!------------------------------------------------------------->
<br /><br /><br />
<div id="copyRight">
<center>
<font class="disclaimer">Microsoft&reg; Excel&reg; and Microsoft&reg; Visual Basic&reg; are registered tradmarks of Microsoft&reg;</font><br />
&copy; Wilde XL Solutions (WXLS) <script>document.write(new Date().getFullYear())</script>. All rights reserved.
</center>
</div>

</body>
</html>


