--MIT License

--Copyright (c) 2021 iappyx

--Permission is hereby granted, free of charge, to any person obtaining a copy
--of this software and associated documentation files (the "Software"), to deal
--in the Software without restriction, including without limitation the rights
--to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
--copies of the Software, and to permit persons to whom the Software is
--furnished to do so, subject to the following conditions:

--The above copyright notice and this permission notice shall be included in all
--copies or substantial portions of the Software.

--THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
--IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
--FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
--AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
--LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
--OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
--SOFTWARE.

--INSTALLATION INSTRUCTIONS:
--Install this file in ~/Library/Application Scripts/com.microsoft.Powerpoint/

on SendFileWithOutlook(parameters)
	set {emailSubject, emailAttachment} to SplitParameterString(parameters, ";")
	tell application "Microsoft Outlook"
		set newEmail to (make new outgoing message with properties {subject:emailSubject})
		tell newEmail
			set emailAttachmentPOSIX to POSIX file emailAttachment
			make new attachment with properties {file:emailAttachmentPOSIX as alias}
			delay 0.5
		end tell
		open newEmail
		activate newEmail
	end tell
	return "Done"
end SendFileWithOutlook

on SendFileWithMail(parameters)
	set {emailSubject, emailAttachment} to SplitParameterString(parameters, ";")
	tell application "Mail"
		set newEmail to (make new outgoing message with properties {subject:emailSubject, visible:true})
		tell newEmail
			set emailAttachmentPOSIX to POSIX file emailAttachment
			make new attachment with properties {file name:emailAttachmentPOSIX as alias}
			delay 0.5
		end tell
		activate newEmail
	end tell
	return "Done"
end SendFileWithMail

on PasteTextIntoWord()
	tell application "Microsoft Word"
		activate
		set NewDocument to create new document
		paste and format (text object of selection) type format surrounding formatting with emphasis
	end tell
	return "Done"
end PasteTextIntoWord

on SplitParameterString(parameters, fieldSeparator)
	tell AppleScript
		set textItemDelimiters to text item delimiters
		set text item delimiters to fieldSeparator
		set parameterItems to text items of parameters
		set text item delimiters to textItemDelimiters
	end tell
	return parameterItems
end SplitParameterString

on CheckIfAppleScriptPluginIsInstalled()
	return "1.0"
end CheckIfAppleScriptPluginIsInstalled
