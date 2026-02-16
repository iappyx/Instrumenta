
![Alt text](img/logo-instrumenta-small.png?raw=true "Instrumenta Powerpoint Toolbar")
# Instrumenta Powerpoint Toolbar

Many strategy consultancy firms have proprietary Powerpoint add-ins that provide access to often used tools and features that help to quickly fine tune a powerpoint presentation. After spending 10 years in strategy consulting and joining 'the industry' myself, I was looking for an alternative for the add-ins I was used to. Although lots of commercial options are available, I could not find a free and open source alternative. 

As a spare time project in times of COVID-19, I decided to create Instrumenta as a free and open source consulting powerpoint toolbar. The ultimate goal is to create a feature rich toolbar that is compatible with both Windows and Mac versions of Microsoft Office. MIT-licensed, and use at your own risk. If you use the code to create your own toolbar —whether it’s for free or commercial purposes— it would be appreciated if you let me know and provide proper attribution in accordance with the MIT license requirements.

[@iappyx]( https://github.com/iappyx )


![Alt text](img/instrumenta-win-1.30.png?raw=true "Instrumenta Powerpoint Toolbar (Windows)")

![Alt text](img/instrumenta-win-1.44.png?raw=true "Instrumenta Powerpoint Toolbar (Windows)")


# Features
Instrumenta has 270+ features, current features include:
| Group | Feature |
|-|-|
| Generic | - Basic formatting and shortcuts to different frequently used powerpoint functions |
| Text | - Increase/decrease line spacing<br>- Remove text from shape<br>- Remove hyperlinks from shape<br>- Set or toggle autofit <br>- Remove formatting<br>- Swap text<br>- Remove strikethrough text<br>- Insert special characters<br>- Ticks and crosses<br>- Replace fonts<br>- Color bold text<br>- Set proofing language for all slides <br>- Split text into multiple shapes <br>- Merge text from multiple shapes into one shape <br>- Master stylesheets with Heading 1–3, Paragraph, Quote, and five custom text styles you can apply across your slides |
| Shapes | - Group shapes by row/column<br>- Select shapes by fill and/or line color<br>- Select shapes by width and/or height<br>- Select shapes by type of shape<br>- Swap position of two shapes<br>- Copy rounded corners of shapes to selected shapes<br>- Copy shapetype and all adjustments of shapes to selected shapes<br>- Rectify lines<br>- Clone shapes to right/down<br>- Copy/paste position and dimensions of shapes (across slides)<br>- Copy shape to multiple slides (multislide shape)<br>- Update position and dimensions of selected multislide shape on all slides<br>- Delete selected multislide shape on all slides<br>- Crop shape to slide<br>- Connect sides of two rectangles <br>- Increase shape transparency <br>- Toggle lock aspect ratio of shapes <br>- Resize and space elements evenly (horizontally and vertically)|
| Pictures | - Apply same crop to selected pictures <br>- Crop picture to slide |
| Align, distribute and size | - Align, distribute and size shapes<br>- Align objects over table cells, rows or columns <br>- Arrange shapes <br>- Set same height and/or width for shapes<br>- Size shapes to tallest, shortest, widest or narrowest<br>- Remove, increase or decrease horizontal/vertical gap between shapes<br>- Remove, increase or decrease margins for shapes<br>- Remove, increase or decrease margins for tables or selected cells <br>- Stretch objects to top, left, right or bottom<br>- Stretch objects to top edge, left edge, right edge or bottom edge  |
| Table | - Format table<br>- Quick format table (preset)<br>- Optimize table height while preserving width<br>- Move rows and columns within a table<br>- Add, delete, increase or decrease column/row gaps <br>- Distribute columns/rows while ignoring column/row gaps <br>- Convert table to shapes<br>- Convert shapes to Table<br>- Transpose table<br>- Insert column preserving other column widths<br>- Split table by row / column<br>- Sum columns in table (all values above selected cells)<br>- Sum rows in table (all values left from selected cells) |
| Export | - Save selected slides as new file<br>- E-mail selected slides (as PDF or PPT)<br>- Copy storyline to clipboard<br>- Export storyline to Word<br>- Paste storyline in shape<br>- Copy slide notes to clipboard<br>- Export slide notes to Word|
| Paste and insert | - Insert slide from slide library<br>- Copy selected slides to slide library<br>- Harvey Balls<br>- Traffic lights (RAG status)<br>- Legend<br>- Star rating (0-5)<br>- Average Harvey Balls, Traffic lights and star ratings based on selected<br>- Numbered captions to shapes (including tables and images)<br>- Renumber captions across slides<br>- Sticky notes<br>- Move sticky notes on and off this slide/all slides<br>- Remove sticky notes from this slide/all slides<br>- Convert comments to sticky notes<br>- Steps counter (per slide and cross-slides)<br>- Agenda pages<br>- Stamps<br>- Move stamps on and off this slide/all slides<br>- Remove stamps from this slide/all slides<br>- Insert process (SmartArt) <br>- Insert Emoji <br>- Insert QR-code|
| Advanced | - Mail merge a specific slide based on Excel-file<br>- Mail merge full presentation based on Excel-file (creating seperate presentations)<br>- Manually replace all merge fields on all slides (can be used for templates)<br>- Move selected slides to end and hide<br>- Remove all hidden slides<br>- Remove animations from all/selected slides<br>- Remove slide entry transitions from all/selected slides<br>- Remove speaker notes from all/selected slides<br>- Remove comments from all/selected slides<br>- Remove all unused master slides<br>- Convert all/selected slides to pictures (readonly)<br>- Watermark and convert all/selected slides to pictures (readonly)<br>- Anonymize all/selected slides with Lorem Ipsum<br>- Add (hidden) tags to slides and shapes<br>- Manage (hidden) tags of slides and shapes<br>- Select sliderange based on tags<br>- Select sliderange based on specific stamps on those slide<br>- Lock and unlock position of objects on slide<br>- Replace colors in all/selected slides<br>- Replace colors in selected shapes<br>- Check for new versions of Instrumenta in the About-dialog <br>- Find Instrumenta features<br>- Change Instrumenta settings |

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

# Keyboard shortcuts
Due to limitations in VBA, Instrumenta does not support keyboard shortcuts out-of-the-box. PowerPoint does not provide built-in functionality for assigning shortcuts to macros.
I have been working on a keyboard shortcut companion called [Instrumenta Keys](https://github.com/iappyx/Instrumenta-Keys). 
It works on both Windows and Mac, but is highly experimental.

You can try Instrumenta Keys, or assign functions to the Quick Access Toolbar instead and use pre-defined shortcuts for these, see [#37](https://github.com/iappyx/Instrumenta/issues/37).

# How to install 
See installation instructions [here](INSTALL.md).

# How to build from source
Creating your own build is very simple, all coding is done in PowerPoint.

- Open "InstrumentaPowerpointToolbar.pptm" from the "src" directory in PowerPoint
- Through PowerPoint settings, enable the "Developer" tab in the PowerPoint ribbon.
- All coding is done in the Visual Basic Editor (VBA IDE) of PowerPoint, the .bas-files in the "src" directory are there for reference only. I export them after every build.
- You can use the pptm-file and update or create your own and copy-paste the code
- To customize the Ribbon you can use [https://github.com/fernandreu/office-ribbonx-editor](https://github.com/fernandreu/office-ribbonx-editor) on the pptm-file.
- In PowerPoint, save the file as a "PowerPoint Add-in (*.ppam)" file to create your own build

The code for the installers can be found in `/src/Installers/`. For Windows this is a NSIS-script and for Mac the installer can be built with `build.sh`.
