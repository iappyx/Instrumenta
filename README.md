# Instrumenta Powerpoint Toolbar

Many strategy consultancy firms have proprietary Powerpoint add-ins that provide access to often used tools and features that help to quickly fine tune a powerpoint presentation. After spending 10 years in strategy consulting and joining 'the industry' myself, I was looking for an alternative for the add-ins I was used to. Although lots of commercial options are available I could not find a free and open source alternative. 

I decided to create Instrumenta as a free and open source consulting powerpoint toolbar. This is an initial version. The ultimate goals is to create a feature rich toolbar that is compatible with both Windows and Mac versions of Microsoft Office.

![Alt text](img/instrumenta-win.png?raw=true "Instrumenta Powerpoint Toolbar (Windows)")


# Features
Current features include:
- Basic formatting and shortcuts to different frequently used powerpoint functions
- Align, distribute and size shapes
- Set same height and/or width for shapes
- Size shapes to tallest, shortest, widest or narrowest
- Remove, increase or decrease horizontal/vertical gap between shapes
- Remove, increase or decrease margins for shapes
- Remove, increase or decrease margins for tables or selected cells
- Select shapes by fill and/or line color
- Select shapes by width and/or height
- Select shapes by type of shape
- Swap position of two shapes
- Remove text from shape
- Remove formatting
- Swap text
- Convert table to shapes
- Copy rounded corners of shapes to selected shapes
- Copy shapetype and all adjustments of shapes to selected shapes
- Replace fonts
- Set proofing language for all slides
- Format table
- E-mail selected slides (as PDF or PPT)
- Copy storyline to clipboard
- Paste storyline in shape
- Remove animations from all slides
- Remove slide entry transitions from all slides
- Remove speaker notes from all slides
- Remove comments from all slides
- Ticks and crosses
- Harvey Balls

# Platform support
All functions supported in Windows. 

The add in will work in OS X, with some minor issues:
* Some icons will not show correctly in the ribbon (underlying functionality will work)
* Export to E-mail (as PPT or PDF) is not yet supported

# Feature requests and contributions
I am happy to receive feature requests and code contributions! Let's make the best toolbar together. For feature requests please create new issue and label it as an enhancement (https://github.com/iappyx/Instrumenta/issues/new/choose). If you want to contribute, please make sure that the code can be freely used as open source code.

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
