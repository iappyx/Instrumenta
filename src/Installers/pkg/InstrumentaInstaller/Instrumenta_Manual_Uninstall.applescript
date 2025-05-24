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

set instrumentaFolder to POSIX path of (path to home folder) & "/Library/Application Scripts/com.microsoft.Powerpoint/"

set uninstallInstrumenta to button returned of (display dialog "Do you want to uninstall Instrumenta?" buttons {"No", "Yes"} default button "Yes")

if uninstallInstrumenta is "Yes" then
    set instrumentaFiles to {"InstrumentaAppleScriptPlugin.applescript", "InstrumentaPowerpointToolbar.ppam"}
    
    repeat with fileName in instrumentaFiles
        do shell script "rm -f " & quoted form of (instrumentaFolder & "/" & fileName)
    end repeat
    
    set uninstallKeys to button returned of (display dialog "Do you also want to uninstall Instrumenta Keys?" buttons {"No", "Yes"} default button "Yes")
    
    if uninstallKeys is "Yes" then
        set keysFiles to {"InstrumentaKeys.applescript", "InstrumentaKeysHelper.ppam"}
        
        repeat with fileName in keysFiles
            do shell script "rm -f " & quoted form of (instrumentaFolder & "/" & fileName)
        end repeat
    end if
end if

display dialog "Manual uninstallation complete! The selected components have been removed." buttons {"OK"} default button "OK"

