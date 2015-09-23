# Calculator Add-in as an Excel Task Pane

This example shows a simple add-in as an Office Excel task pane. As you type in this Calculator, it puts that data into cells in an Excel spreadsheet, and then generates the corresponding formula.

![excelonlinesm](https://cloud.githubusercontent.com/assets/13560879/9948988/50da3012-5d5d-11e5-97ed-c3d9c0804ec5.png)

### Table of Contents
- [Overview](#overview)

- [Run in Office Playground](#run-in-office-playground)

- [Run in Excel 2013 Desktop](#run-in-excel-2013-desktop)

- [Run in Excel Online](#run-in-excel-online)

- [Additional resources](#additional-resources)

##Overview

This add-in reads its input from a set of number and operator keys, each defined in an HTML table cell. Each time you enter a number and press the Enter key, that number is entered into an Excel spreadsheet cell and the Excel cursor moves down one row. Clicking on an operator (such as "+") and then equals ("=") generates a formula (in this case a SUM) that is entered into the next cell down.

##Run in Office Playground

The easiest way to run this sample is to open it in the playground for Office Add-ins: 

1. Go to http://aka.ms/Vnp9gk
2. Log in using a Microsoft account.
3. Click the Run Project icon to launch the sample in Excel Online.


##Run in Excel 2013 Desktop

To run this sample in Excel 2013 Desktop:

1. Host these files on a local network share. For more information, please see: https://msdn.microsoft.com/EN-US/library/office/fp123503.aspx
2. Open up an Office app (Excel, Word or PowerPoint), open a document, and then select File > Options > Trust Center > Trust Center Settings > Trusted App Catalogs.
3. Type the location of the directory on your local network share into the Catalog Url text field, and click Add Catalog. Make sure the new location's Show in Menu check box is selected.
4. Click OK. Close the Office app and launch it again, so the changes take effect.
5. Go to Insert > My Apps > Shared Folder and select Calculator Add-in, and then click Insert. If you don't see the add-in, click Refresh.
6. The add-in's task pane opens next to your document.

##Run in Excel Online

**Note:** You will need a subscription to Office 365. If you don't have a subscription, obtain a free account for 30 days from https://portal.microsoftonline.com/Signup/MainSignUp.aspx?OfferId=6881A1CB-F4EB-4db3-9F18-388898DAF510&DL=DEVELOPERPACK

To run this sample in Excel Online:

1. You can host these files locally (on localhost), or online, such as on AWS, Azure, or Heroku. 
2. Edit the manifest file Calculator.xml and change the DefaultValue of the SourceLocation to the URL where the Home.html file is hosted.
3. Go to the Office 365 portal (https://portal.office.com) and click on Admin in the app launcher on the top left hand corner.
4. Select SharePoint > apps > App Catalog > Apps for Office.
5. Select the "+" button to add a new add-in, and choose Calculator.xml from its directory. Press OK and the add-in will install.
6. Open the app launcher from the top left hand corner and select Excel Online.
7. When the Excel Online app opens, open a blank spreadsheet, then go to Insert > Office Add-ins and select Calculator under My Organization. If you don't see the add-in, press Refresh. Press Insert and the add-in should appear.

##Additional resources

[Office Add-in publishing](https://msdn.microsoft.com/EN-US/library/office/fp123517.aspx)
