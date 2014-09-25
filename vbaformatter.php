<?php
$strArgs = "Global Const debugMode As Boolean = False '//Debug switch

'// FileType Enum
Public Enum ezFileType
    ezXLS = -4143
    ezCSV = 6
    ezXLA = 18
    ezXLSX = 51
    ezXLSM = 52
End Enum

'// Msgbox Enums
    '// Button Style
Public Enum ezMsgButtons
    ezOK = 0
    ezOKCancel = 1
    ezAbortRetryIgnore = 2
    ezYesNoCancel = 3
    ezYesNo = 4
    ezRetryCancel = 5
End Enum

    '// Icon Style
Public Enum ezMsgIcons
    ezCritical = 16
    ezQuestion = 32
    ezExclamation = 48
    ezInformation = 64
End Enum

    '// Default Button
Public Enum ezMsgDefaultButton
    ezButton1 = 0
    ezButton2 = 256
    ezButton3 = 512
End Enum

    '// Misc Settings
Public Enum ezMsgModalState
    ezPauseApplication = 0
    ezPauseAllApplications = 4096
    ezForceToForeground = 65536
    ezRightAlignText = 524288
    ezRightToLeftReading = 1048576
End Enum";

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
		if(in_array($word, $keywords)){
		    echo $word.' is in keywords list<br />'; //DEBUG LINE
		    echo 'line currently reads: '.$newLine.'<br />'; //DEBUG LINE
		    preg_match_all("/$word(?=&nbsp;)|(?<=&nbsp;)$word/", $newLine, $kwMatch); //DEBUG LINE
		    echo $kwMatch[0][0]." (match pattern successful) <br />"; //DEBUG LINE
		$newLineTwo = preg_replace("/$word(?=&nbsp;)|(?<=&nbsp;)$word/", "<span style=\"color: navy;\">".$word."</span>", $newLine);
		    echo 'line now reads: '.$newLineTwo.'<br />'; //DEBUG LINE
		    if($newLine = $newLineTwo) echo "***REPLACE FAILED***<br />"; //DEBUG LINE
		}
	}
		$strArgs = str_replace($line, $newLine, $strArgs);
}


// Change PHP line break to <br /> to preserve formatting in HTML:
$strArgs = nl2br($strArgs);

// Create output page

echo "<strong>".$strArgs."</strong>"
?>
