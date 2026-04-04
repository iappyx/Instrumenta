--MIT License

--Copyright (c) 2025 iappyx

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

-- Run with: osascript InstrumentaKeys.applescript "ResizeShapes"

on run argv
	
	-- SETTINGS
	
	set InstrumentaKeysDefaultFunction to "ShowAboutDialog"
	
	-- PREPARE TMP FILE
	
	if not application "Microsoft PowerPoint" is running then
		display dialog "PowerPoint is not running. Exiting script."
		return
	end if
	
	if (count of argv) < 1 then
		set functionName to InstrumentaKeysDefaultFunction
		
	else if ((count of argv) > 1) then
		display dialog "Too many arguments, script can accept only one function. Example: osascript InstrumentaKeys.applescript ShowAboutDialog"
		return
	else
		set functionName to item 1 of argv
	end if
	
	tell application "Microsoft PowerPoint"
		set tempAlias to path to temporary items
	end tell
	
	set tempPath to POSIX path of tempAlias
	set filePath to tempPath & "tmpInstrumentaKeys.txt"
	
	try
		do shell script "rm -f " & quoted form of filePath
	on error errMsg
		
	end try
	
	try
		set fileRef to open for access (POSIX file filePath) with write permission
		write functionName to fileRef
		close access fileRef
	on error errMsg
		display dialog "Error writing file: " & errMsg
		return -- Stop script execution if there's an error
	end try
	
	-- Continue executing only if a function name was successfully written
	
	tell application "System Events"
		
		tell process "Microsoft PowerPoint"
			-- Find the active PowerPoint window
			set appWins to (every window)
			
			if (count of appWins) = 0 then
				
				return
				
			else if ((count of appWins) = 1) then
				
				try
					set mainWin to (first window whose value of attribute "AXMain" is true)
					
				on error
					set mainWin to (first window)
					if value of attribute "AXMinimized" of mainWin is true then
						set value of attribute "AXMinimized" of mainWin to false
						set value of attribute "AXMain" of mainWin to true
					else
						return
					end if
				end try
				
			else
				try
					-- Find the active PowerPoint window
					set mainWin to (first window whose value of attribute "AXMain" is true)
					
				on error
					-- Stop script if there is no active window found (e.g. minimized)
					return
				end try
			end if
			
			-- Store the initially selected ribbon tab
			set initialTab to ""
			set mainTabSelected to false
			repeat with rb in every radio button of tab group 1 of mainWin
				if value of rb is 1 then
					set initialTab to name of rb
					set mainTabSelected to true
					exit repeat
				end if
			end repeat
		end tell
	end tell
	
	-- EXECUTE IN POWERPOINT
	
	-- Activate the [Keys] tab if necessary
	set tabChanged to false
	set keysTabFound to false
	tell application "System Events"
		tell process "Microsoft PowerPoint"
			repeat with rb in every radio button of tab group 1 of mainWin
				if name of rb is "[Keys]" then
					set keysTabFound to true
					if value of rb is 0 then
						click rb
						set tabChanged to true
						delay 0.5
					end if
					exit repeat
				end if
			end repeat
		end tell
	end tell
	
	-- If [Keys] was not found, attempt to find it in the overflow menu
	if not keysTabFound then
		tell application "System Events"
			tell process "Microsoft PowerPoint"
				set mainWin to (first window whose value of attribute "AXMain" is true)
				
				-- Check if the overflow button exists before accessing it
				set overflowButtons to every menu button of tab group 1 of mainWin
				if (count of overflowButtons) > 0 then
					set overflowButton to item 1 of overflowButtons
					click overflowButton
					
					-- Fetch all menu items inside the overflow menu
					set uiElements to entire contents of overflowButton
					
					-- Locate and click [Keys]
					repeat with uiElement in uiElements
						try
							if name of uiElement is "[Keys]" then
								click uiElement
								set tabChanged to true
								set keysTabFound to true
								delay 0.5
								exit repeat
							end if
						end try
					end repeat
				end if
			end tell
		end tell
	end if
	
	-- Stop execution if [Keys] was not found
	if not keysTabFound then
		if not mainTabSelected then
			-- No main tabs selected and [Keys] was not found, then  revert to the first ribbon tab
			tell application "System Events"
				tell process "Microsoft PowerPoint"
					repeat with rb in every radio button of tab group 1 of mainWin
						click rb -- Click first tab found
						exit repeat
					end repeat
				end tell
			end tell
			-- Removed: display dialog "No main tab was selected. Reverted to first visible tab."
		else
			display dialog "Instrumenta Keys helper was not found. Please install it."
		end if
		return
	end if
	
	-- If [Keys] was found in the overflow menu but no main tab was selected, revert to the first tab
	if not mainTabSelected then
		tell application "System Events"
			tell process "Microsoft PowerPoint"
				repeat with rb in every radio button of tab group 1 of mainWin
					click rb -- Click first tab found
					exit repeat
				end repeat
			end tell
		end tell
		--display dialog "No main tab was selected initially. Reverted to first visible tab after selecting [Keys]."
	end if
	
	-- Check if the ribbon was hidden due to double-clicking in the overflow menu
	tell application "System Events"
		tell process "Microsoft PowerPoint"
			
			try
				set scrollExists to exists scroll area 1 of tab group 1 of mainWin
			on error
				set scrollExists to false
			end try
			
			if scrollExists = false then
				
				set overflowButton to first menu button of tab group 1 of mainWin
				if exists overflowButton then
					click overflowButton
					delay 0.5
					
					repeat with uiElement in entire contents of overflowButton
						try
							if name of uiElement is "[Keys]" then
								click uiElement
								delay 0.5
								exit repeat
							end if
						end try
					end repeat
				end if
			end if
			
		end tell
	end tell
	
	-- Click the "Run Function" button
	tell application "System Events"
		tell process "Microsoft PowerPoint"
			repeat with btn in every button of group 1 of scroll area 1 of tab group 1 of mainWin
				if name of btn is "Run Function" then
					click btn
					exit repeat
				end if
			end repeat
		end tell
	end tell
	
	-- Restore the UI to the initially selected ribbon tab, only if it was changed
	if tabChanged then
		tell application "System Events"
			tell process "Microsoft PowerPoint"
				repeat with rb in every radio button of tab group 1 of mainWin
					if name of rb is initialTab then
						click rb
						exit repeat
					end if
				end repeat
			end tell
		end tell
	end if
end run