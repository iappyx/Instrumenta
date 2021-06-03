--Install this file in ~/Library/Application Scripts/com.microsoft.Powerpoint/
--Partly based on the work of Ron de Bruin, https://macexcel.com/examples/mailpdf/macoutlook/

on SendFileWithOutlook(paramString)
	set {subjectValue, attachmentPathValue} to SplitString(paramString, ";")
	tell application "Microsoft Outlook"
		set NewMail to (make new outgoing message with properties {subject:subjectValue})
		tell NewMail
			set AttachmentPath to POSIX file attachmentPathValue
			make new attachment with properties {file:AttachmentPath as alias}
			delay 0.5
		end tell
		open NewMail
		activate NewMail
	end tell
	return "Done"
end SendFileWithOutlook

on SplitString(TheBigString, fieldSeparator)
	tell AppleScript
		set oldTID to text item delimiters
		set text item delimiters to fieldSeparator
		set theItems to text items of TheBigString
		set text item delimiters to oldTID
	end tell
	return theItems
end SplitString