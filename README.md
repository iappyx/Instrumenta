
![Alt text](img/logo-instrumenta-small.png?raw=true "Instrumenta Powerpoint Toolbar")
# Instrumenta Powerpoint Toolbar

Many strategy consultancy firms have proprietary Powerpoint add-ins that provide access to often used tools and features that help to quickly fine tune a powerpoint presentation. After spending 10 years in strategy consulting and joining 'the industry' myself, I was looking for an alternative for the add-ins I was used to. Although lots of commercial options are available, I could not find a free and open source alternative. 

As a spare time project in times of COVID-19, I decided to create Instrumenta as a free and open source consulting powerpoint toolbar. The ultimate goal is to create a feature rich toolbar that is compatible with both Windows and Mac versions of Microsoft Office. MIT-licensed, and use at your own risk. If you use the code to create your own toolbar —whether it’s for free or commercial purposes— it would be appreciated if you let me know and provide proper attribution in accordance with the MIT license requirements.

[@iappyx]( https://github.com/iappyx )

![Alt text](img/instrumenta-win-1.30.png?raw=true "Instrumenta Powerpoint Toolbar (Windows)")


# Features
Current features include:
| Group | Feature |
|-|-|
| Generic | - Basic formatting and shortcuts to different frequently used powerpoint functions |
| Text | - Increase/decrease line spacing<br>- Remove text from shape<br>- Remove hyperlinks from shape<br>- Set or toggle autofit <br>- Remove formatting<br>- Swap text<br>- Remove strikethrough text<br>- Insert special characters<br>- Ticks and crosses<br>- Replace fonts<br>- Color bold text<br>- Set proofing language for all slides <br>- Split text into multiple shapes <br>- Merge text from multiple shapes into one shape |
| Shapes | - Group shapes by row/column<br>- Select shapes by fill and/or line color<br>- Select shapes by width and/or height<br>- Select shapes by type of shape<br>- Swap position of two shapes<br>- Copy rounded corners of shapes to selected shapes<br>- Copy shapetype and all adjustments of shapes to selected shapes<br>- Rectify lines<br>- Clone shapes to right/down<br>- Copy/paste position and dimensions of shapes (across slides)<br>- Copy shape to multiple slides (multislide shape)<br>- Update position and dimensions of selected multislide shape on all slides<br>- Delete selected multislide shape on all slides<br>- Crop shape to slide<br>- Connect sides of two rectangles <br>- Increase shape transparency <br>- Toggle lock aspect ratio of shapes <br>- Resize and space elements evenly (horizontally and vertically)|
| Pictures | - Apply same crop to selected pictures <br>- Crop picture to slide |
| Align, distribute and size | - Align, distribute and size shapes<br>- Align objects over table cells, rows or columns <br>- Arrange shapes <br>- Set same height and/or width for shapes<br>- Size shapes to tallest, shortest, widest or narrowest<br>- Remove, increase or decrease horizontal/vertical gap between shapes<br>- Remove, increase or decrease margins for shapes<br>- Remove, increase or decrease margins for tables or selected cells <br>- Stretch objects to top, left, right or bottom<br>- Stretch objects to top edge, left edge, right edge or bottom edge  |
| Table | - Format table<br>- Quick format table (preset)<br>- Move rows and columns within a table<br>- Add, delete, increase or decrease column/row gaps <br>- Distribute columns/rows while ignoring column/row gaps <br>- Convert table to shapes<br>- Convert shapes to Table<br>- Transpose table<br>- Split table by row / column<br>- Sum columns in table (all values above selected cells)<br>- Sum rows in table (all values left from selected cells) |
| Export | - Save selected slides as new file<br>- E-mail selected slides (as PDF or PPT)<br>- Copy storyline to clipboard<br>- Export storyline to Word<br>- Paste storyline in shape<br>- Copy slide notes to clipboard<br>- Export slide notes to Word|
| Paste and insert | - Insert slide from slide library<br>- Copy selected slides to slide library<br>- Harvey Balls<br>- Traffic lights (RAG status)<br>- Star rating (0-5)<br>- Average Harvey Balls, Traffic lights and star ratings based on selected<br>- Numbered captions to shapes (including tables and images)<br>- Renumber captions across slides<br>- Sticky notes<br>- Move sticky notes on and off this slide/all slides<br>- Remove sticky notes from this slide/all slides<br>- Convert comments to sticky notes<br>- Steps counter (per slide and cross-slides)<br>- Agenda pages<br>- Stamps<br>- Move stamps on and off this slide/all slides<br>- Remove stamps from this slide/all slides<br>- Insert process (SmartArt) <br>- Insert Emoji <br>- Insert QR-code|
| Advanced | - Mail merge a specific slide based on Excel-file<br>- Mail merge full presentation based on Excel-file (creating seperate presentations)<br>- Manually replace all merge fields on all slides (can be used for templates)<br>- Move selected slides to end and hide<br>- Remove all hidden slides<br>- Remove animations from all/selected slides<br>- Remove slide entry transitions from all/selected slides<br>- Remove speaker notes from all/selected slides<br>- Remove comments from all/selected slides<br>- Remove all unused master slides<br>- Convert all/selected slides to pictures (readonly)<br>- Watermark and convert all/selected slides to pictures (readonly)<br>- Anonymize all/selected slides with Lorem Ipsum<br>- Add (hidden) tags to slides and shapes<br>- Manage (hidden) tags of slides and shapes<br>- Select sliderange based on tags<br>- Select sliderange based on specific stamps on those slide<br>- Lock and unlock position of objects on slide<br>- Check for new versions of Instrumenta in the About-dialog <br>- Change Instrumenta settings |

# Platform support
All functions tested in Windows on the latest Office at that moment in time.

The add in will work in OS X, with some minor issues:
* Some icons are not the same as in the Windows-version. Microsoft Office does not support all icons from Windows on the Mac platform.
* *Lock and unlock position of objects on slide* is not supported. This method is not (yet) implemented in VBA for Powerpoint on Mac. However, shapes that have been 'locked' in Windows will be shown as 'locked' on Mac as well.
* *Export to E-mail (as PPT or PDF)*, *Export storyline to Word* and *Export slide notes to Word* are supported but require installation of an AppleScript-file due to OS X sandbox. See installation instructions below.

As stated in the license: THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

# Feature requests and contributions
I am happy to receive feature requests and code contributions! Let's make the best toolbar together. For feature requests please create new issue and label it as an enhancement (https://github.com/iappyx/Instrumenta/issues/new/choose). 

If you want to contribute, please make sure that the code can be freely used as open source code. 
Please only update the files in /src/Modules, /src/Forms, /src/CustomUi and /src/Classes. For security reasons I will not accept updated .pptm or .ppam files.

If you like this plugin, please let me and the community know how you are using this in your daily work: https://github.com/iappyx/Instrumenta/discussions/5

# How to install 
Instrumenta is a Visual Basic for Applications (VBA) add-in that can be installed within Powerpoint, requiring no administrative rights on most enterprise systems.

## Windows
You can save the add-in to your PC and then install the add-in by adding it to the Available Add-Ins list:
- Download the add-in file in the latest release (https://github.com/iappyx/Instrumenta/releases/download/1.40/InstrumentaPowerpointToolbar.ppam) or the latest beta (https://github.com/iappyx/Instrumenta/raw/main/bin/InstrumentaPowerpointToolbar.ppam) and save it in a fixed location
- Open Powerpoint, click the File tab, and then click Options
- In the Options dialog box, click Add-Ins.
- In the Manage list at the bottom of the dialog box, click PowerPoint Add-ins, and then click Go.
- In the Add-Ins dialog box, click Add New.
- In the Add New PowerPoint Add-In dialog box, browse for the add-in file, and then click OK.
- A security notice appears. Click Enable Macros, and then click Close. 
  - Note: If you cannot enable Macros in this dialog, please refer to [these](https://support.microsoft.com/en-gb/topic/a-potentially-dangerous-macro-has-been-blocked-0952faa0-37e7-4316-b61d-5b5ed6024216) instructions from Microsoft to unblock Instrumenta: (1) Open Windows File Explorer and go to the folder where you saved the file; (2) Right-click the file and choose Properties from the context menu; (3) At the bottom of the General tab, select the Unblock checkbox and select OK.
- There now should be an "Instrumenta" page in the Powerpoint ribbon

(Instructions based on https://support.microsoft.com/en-us/office/add-or-load-a-powerpoint-add-in-3de8bbc2-2481-457a-8841-7334cd5b455f)

## Mac
You can save the add-in to your Mac and then install the add-in by adding it to the Add-Ins list:
- Download the add-in file in the latest release (https://github.com/iappyx/Instrumenta/releases/download/1.40/InstrumentaPowerpointToolbar.ppam) or the latest beta (https://github.com/iappyx/Instrumenta/raw/main/bin/InstrumentaPowerpointToolbar.ppam) and save it in a fixed location
- Open Powerpoint, click Tools in the application menu, and then click Add-ins...
- In the Add-Ins dialog box, click the + button, browse for the add-in file, and then click Open.
- Click Ok to close the Add-ins dialog box
- There now should be an "Instrumenta" page in the Powerpoint ribbon

Optional steps to show group titles:
- By default group titles are hidden in the ribbon on the Mac
- Open PowerPoint, click PowerPoint in the application menu
- Then click Preferences
- Click View
- Check "Show group titles"

Additional optional steps to enable export to Outlook and Word:
- Download the AppleScript file (https://github.com/iappyx/Instrumenta/releases/download/1.40/InstrumentaAppleScriptPlugin.applescript) 
- Copy the AppleScript file to *~/Library/Application Scripts/com.microsoft.Powerpoint/*
- Please note that this is in the library folder of the *current user*. If the folder does not exist, create it.
- In some cases a reboot of your Mac might be required

The AppleScript file will not change that often, only in case of major changes to the export features, or new features that require the use of AppleScript. In that case Instrumenta can detect if the latest version of the file is installed and will inform you when using a feature that requires an updated version.

## How to build from source
Creating a build is very simple, all coding is done in PowerPoint.

- Open "InstrumentaPowerpointToolbar.pptm" from the "src" directory in PowerPoint
- Through PowerPoint settings, enable the "Developer" tab in the PowerPoint ribbon.
- All coding is done in the Visual Basic Editor (VBA IDE) of PowerPoint, the .bas-files in the "src" directory are there for reference only. I export them after every build.
- You can use the pptm-file and update or create your own and copy-paste the code
- To customize the Ribbon you can use [https://github.com/fernandreu/office-ribbonx-editor](https://github.com/fernandreu/office-ribbonx-editor) on the pptm-file.
- In PowerPoint, save the file as a "PowerPoint Add-in (*.ppam)" file to create your own build

## Upgrade
To upgrade make sure Powerpoint is closed and overwrite the add-in file with the new version. 

### Upgrading from version < 1.0 to version 1.0 or higher
Please know that all versions before the first major release (v1.0) have a different filename. If you are upgrading to v1.0 please remove the old add-in first in the Add-ins dialog box in Powerpoint and then install the new version as mentioned above.
