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

set scriptFolder to POSIX path of (do shell script "dirname " & quoted form of POSIX path of (path to me))
set instrumentaFolder to POSIX path of (path to home folder) & "/Library/Application Scripts/com.microsoft.Powerpoint/"
do shell script "mkdir -p " & quoted form of instrumentaFolder
set installInstrumenta to button returned of (display dialog "Do you want to install Instrumenta?" buttons {"No", "Yes"} default button "Yes")

if installInstrumenta is "Yes" then
    set instrumentaFiles to {"InstrumentaAppleScriptPlugin.applescript", "InstrumentaPowerpointToolbar.ppam"}
    
    repeat with fileName in instrumentaFiles
        set sourceFile to scriptFolder & "/" & fileName
        set destinationFile to instrumentaFolder & "/" & fileName
        do shell script "cp " & quoted form of sourceFile & " " & quoted form of destinationFile
    end repeat
    
    set installKeys to button returned of (display dialog "Do you also want to install Instrumenta Keys?" buttons {"No", "Yes"} default button "Yes")
    
    if installKeys is "Yes" then
        set keysFiles to {"InstrumentaKeys.applescript", "InstrumentaKeysHelper.ppam"}
        
        repeat with fileName in keysFiles
            set sourceFile to scriptFolder & "/" & fileName
            set destinationFile to instrumentaFolder & "/" & fileName
            do shell script "cp " & quoted form of sourceFile & " " & quoted form of destinationFile
        end repeat
    end if
    
    tell application "Microsoft PowerPoint"
        try
            set instrumentaToolbarAddIn to register add in (instrumentaFolder & "/InstrumentaPowerpointToolbar.ppam")
            set auto load of instrumentaToolbarAddIn to true
            set loaded of instrumentaToolbarAddIn to true
            
            if installKeys is "Yes" then
                set instrumentaKeysAddIn to register add in (instrumentaFolder & "/InstrumentaKeysHelper.ppam")
                set auto load of instrumentaKeysAddIn to true
                set loaded of instrumentaKeysAddIn to true
            end if
        on error
            display alert "Plugin activation failed" message "Failed to register one or more plugins" as critical buttons {"Cancel"}
            return
        end try
    end tell
end if

display dialog "Manual installation complete! Installed components are now active." buttons {"OK"} default button "OK"
