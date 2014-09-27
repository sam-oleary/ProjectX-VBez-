
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

$strArr = explode("\n", $strArgs);
$theme = explode("|", $_POST['optTheme']);

/********************
* $theme Index Vals *
* [0] Background	*
* [1] Text			*
* [2] Keyword		*
* [3] Comment		*
* [4] Number		*
* [5] String		*
* [6] Line Text		*
* [7] Line BGround	*
*********************/
$lineCountLen = sizeof($strArr);
$countOutputText = "<span style=\"background-color: ".$theme[7]."; color: ".$theme[6].";\">".str_pad("", 6 * (strlen($lineCountLen) + 1), "&nbsp;", STR_PAD_LEFT)."</span>";

$keywords = Array("Alias", "And", "As", "Boolean", "ByRef", "ByVal", "Call", "Case", "CBool", "CDate", "CDbl", "CInt",
	"Class", "CLng", "CObj", "Const", "CStr", "Date", "Debug", "Declare", "Dim", "Do", "Double", "Each", "Else",
	"ElseIf", "End", "EndIf", "Enum", "Erase", "Error", "Exit", "False", "For", "Friend", "Function", "Get", "Global",
	"GoTo", "If", "Implements", "In", "Inherits", "Integer", "Interface", "Is", "Let", "Lib", "Like", "Long",
	"Loop", "Me", "Mod", "Module", "New", "Next", "Not", "Nothing", "Object", "Of", "On", "Option", "Optional",
	"Or", "OrElse", "Print", "Private", "Property", "Protected", "Public", "ReadOnly", "ReDim", "Resume", "Return",
	"Select", "Set", "Shared", "Static", "Step", "Stop", "String", "Sub", "Then", "To", "True", "TypeOf", "Variant",
	"Wend", "When", "While", "With", "WriteOnly");

	$i =0;
	
foreach($strArr as $line){

	$i++;
	
	preg_match_all("'\'.*'", $line, $matches);
	
	$newLine = str_replace($matches[0][0], "<span style=\"color: ".$theme[3].";\">".$matches[0][0]."</span>",$line);

		If (strpos($newLine, '\'') <> FALSE) {
			preg_match_all("'.*\''", $newLine, $matches);
			$tmpComment = str_replace($matches[0][0], "",$newLine);
			$newLine = $matches[0][0];
		}
	
		foreach($keywords as $word){
			$newLine = preg_replace("/(?<!\w)$word(?!\w)/","<span style=\"color: ".$theme[2]."; font-weight: bold;\">".$word."</span>", $newLine);
		}
		
		preg_match_all("/(?<=( |;|>))\d+/", $newLine, $matchesI);
			foreach($matchesI[0] as $match){
			$newLine = str_replace($match, "<span style=\"color: " . $theme[4] .";\">".$match."</span>", $newLine);
			}

		If (strpos($newLine, '\'') <> FALSE) {
			$newLine = $newLine . $tmpComment;
		}
	
	$strpad = (strlen($newLine) - strlen(ltrim($newLine))); // no of characters to pad by
	$pad_string = "&nbsp;";
	
	// remove formatting from quoted text
	preg_match_all('/"(?!color:)(?=[^>])(.*?)"(?=[^>]*(<|$))/', $newLine, $matchesQ);
		foreach($matchesQ[0] as $match){
			$newLine = str_replace($match, "<span style=\"color: " . $theme[5] .";\">".strip_tags($match)."</span>", $newLine); 
			}
			
	$HTMLtext = str_pad(trim($newLine), strlen(trim($newLine))+($strpad*strlen($pad_string)), $pad_string, STR_PAD_LEFT);
	$lineCount = "<span style=\"background-color: ".$theme[7]."; color: ".$theme[6].";\">".str_pad($i, (strlen($i)+(strlen($pad_string)*(strlen($lineCountLen)-strlen($i)))), $pad_string, STR_PAD_LEFT).":</span> ";
	
	$outputText = $outputText . "<br />" . $HTMLtext;
	$countOutputText = $countOutputText."<br />".$lineCount.$HTMLtext;
	$newLine = "";
}

if($_POST['blnNumber'] ==1){
	$txtDisp = $countOutputText;
	} else {
	$txtDisp = $outputText;
	}
?>


<div>

	<div style="width: 1117px; overflow-x: auto; background-color: <?PHP echo $theme[0]; ?>; color: <?PHP echo $theme[1]; ?>; font-family: 'Courier New', Courier, monospace; font-size: 12px;">
		<?PHP echo $txtDisp; ?></div>
		<br /><br />
<?PHP
// Create HTML page

$start = '<!DOCTYPE html>
<html>
<head>
<title>VBA web formatter by Wilde XL Solutions &copy; </title>
</head>
<body>
<center>
<br /><br />
<span style="color: #333; font-family: Arial, sans-serif; font-size: 12px; border: 1px solid red;">
This page was created using <a href="http://www.wxls.co.uk/formatmyvba" target="_blank">Wilde XL Solutions VBA Web Formatter</a>
</span>
<br /><br />
</center>';

$finish = "<br /><br /><br />
</body>
</html>";



$content = $start.'<div style="font-family: Courier New, Courier, monospace; font-size: 12px; background-color: '.$theme[0].'; color: '.$theme[1].';">'.$txtDisp.'</div>'.$finish;
session_start();
$_SESSION['txtOP'] = $content;

echo '<center><form action="..\OP\DL.php" method="post"><input type="submit" value="Download HTML file" /></form></center>';

?>
<br /><br />
<center><span style="font-family: 'Arial', sans-serif;">-Or click the text below and copy & paste into your own HTML file!-</span></center><br /><br />
	<center><pre><textarea cols=120 rows=40 onclick="this.focus(); this.select()" readonly="readonly"><?PHP echo "<span style=\"font-family: 'Courier New', Courier, monospace; font-size: 12px; background-color: ".$theme[0]."; color: ".$theme[1].";\">".str_replace("&", "&amp;", $txtDisp)."</span>"; ?></textarea>
		</pre></center><br /><br />


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


