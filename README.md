
![Alt text](img/logo-instrumenta-small.png?raw=true "Instrumenta Powerpoint Toolbar")
# Instrumenta Powerpoint Toolbar

Many strategy consultancy firms have proprietary Powerpoint add-ins that provide access to often used tools and features that help to quickly fine tune a powerpoint presentation. After spending 10 years in strategy consulting and joining 'the industry' myself, I was looking for an alternative for the add-ins I was used to. Although lots of commercial options are available I could not find a free and open source alternative. 

As a spare time project in times of COVID-19, I decided to create Instrumenta as a free and open source consulting powerpoint toolbar. This is an initial version. The ultimate goal is to create a feature rich toolbar that is compatible with both Windows and Mac versions of Microsoft Office.

![Alt text](img/instrumenta-win-0.7.png?raw=true "Instrumenta Powerpoint Toolbar (Windows)")


# Features
Current features include:
| Group | Feature |
|-|-|
| Generic | - Basic formatting and shortcuts to different frequently used powerpoint functions |
| Text | - Increase/decrease line spacing<br>- Remove text from shape<br>- Remove formatting<br>- Swap text<br>- Insert special characters<br>- Ticks and crosses<br>- Replace fonts<br>- Set proofing language for all slides |
| Shapes | - Select shapes by fill and/or line color<br>- Select shapes by width and/or height<br>- Select shapes by type of shape<br>- Swap position of two shapes<br>- Copy rounded corners of shapes to selected shapes<br>- Copy shapetype and all adjustments of shapes to selected shapes<br>- Rectify lines<br>- Clone shapes to right/down<br>- Copy/paste position and dimensions of shapes (across slides)<br>- Copy shape to multiple slides (multislide shape)<br>- Update position and dimensions of selected multislide shape on all slides<br>- Delete selected multislide shape on all slides<br>- Crop shape to slide<br>- Connect sides of two rectangles  |
| Pictures | - Crop picture to slide |
| Align, distribute and size | - Align, distribute and size shapes<br>- Align objects over table cells, rows or columns <br>- Set same height and/or width for shapes<br>- Size shapes to tallest, shortest, widest or narrowest<br>- Remove, increase or decrease horizontal/vertical gap between shapes<br>- Remove, increase or decrease margins for shapes<br>- Remove, increase or decrease margins for tables or selected cells |
| Table | - Format table<br>- Convert table to shapes<br>- Transpose table<br>- Sum columns in table (all values above selected cells)<br>- Sum rows in table (all values left from selected cells) |
| Export | - E-mail selected slides (as PDF or PPT) |
| Storyline | - Copy storyline to clipboard<br>- Paste storyline in shape |
| Paste and insert | - Harvey Balls<br>- Traffic lights (RAG status)<br>- Sticky notes<br>- Move sticky notes on and off this slide/all slides<br>- Remove sticky notes from this slide/all slides<br>- Convert comments to sticky notes<br>- Steps counter (per slide and cross-slides)<br>- Agenda pages<br>- Stamps<br>- Move stamps on and off this slide/all slides<br>- Remove stamps from this slide/all slides<br>- Insert process (SmartArt) |
| Advanced | - Remove animations from all slides<br>- Remove slide entry transitions from all slides<br>- Remove speaker notes from all slides<br>- Remove comments from all slides<br>- Remove unused master slides<br>- Add (hidden) tags to slides and shapes<br>- Manage (hidden) tags of slides and shapes<br>- Select sliderange based on tags<br>- Select sliderange based on specific stamps on those slide<br>- Check for new versions of Instrumenta in the About-dialog |

# Platform support
All functions tested in Windows on the latest Office at that moment in time.

The add in will work in OS X, with some minor issues:
* Some icons will not show correctly in the ribbon (underlying functionality will work). Custom icons for Instrumenta are on the backlog.
* Export to E-mail (as PPT or PDF) is not supported due to OS X sandbox. There are potential solutions, but those require a lot of manual user configuration and installation of custom scripts. This will not be supported for now.

As stated in the license: THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

# Feature requests and contributions
I am happy to receive feature requests and code contributions! Let's make the best toolbar together. For feature requests please create new issue and label it as an enhancement (https://github.com/iappyx/Instrumenta/issues/new/choose). If you want to contribute, please make sure that the code can be freely used as open source code.

If you like this plugin, please let me and the community know how you are using this in your daily work: https://github.com/iappyx/Instrumenta/discussions/5

# How to install 

You can save the add-in to your computer and then install the add-in by adding it to the Available Add-Ins list:
- Download the add-in file (https://github.com/iappyx/Instrumenta/raw/main/Instrumenta%20Powerpoint%20Toolbar.ppam) and save it in a fixed location 
- Open Powerpoint, click the File tab, and then click Options
- In the Options dialog box, click Add-Ins.
- In the Manage list at the bottom of the dialog box, click PowerPoint Add-ins, and then click Go.
- In the Add-Ins dialog box, click Add New.
- In the Add New PowerPoint Add-In dialog box, browse for the add-in file, and then click OK.
- A security notice appears. Click Enable Macros, and then click Close.
- There now should be an "Instrumenta" page in the Powerpoint ribbon

(Instructions based on https://support.microsoft.com/en-us/office/add-or-load-a-powerpoint-add-in-3de8bbc2-2481-457a-8841-7334cd5b455f)
