# ahk-daily--mainly AHK but other things as well
A collection of scripts I developed to use in my daily tasks.

## 1. PrinttoPDF.ahk
foo.pdf -> Microsoft Print to PDF -> foo_COPY.pdf

Prints pdfs to pdfs, appends "\_COPY" to the filename.

Drag and drop.

Uses Adobe Reader command line options:
https://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/Acrobat_SDK_developer_faq.pdf#page=24

## 2. Fedex Label Creator
Gui thing to fill out the shipping form on fedex ship manager webpage with Excel data.

After shipping, automatically prints the label, saves a screenshot of the created shipment page, and extracts the tracking number for further usage.

Most useful part of this horrific code is probably the fedex ship manager link structure used around line 250 and 720. All them dom elements, etc.

Don't forget to show the right directory to chrome.ahk and login credentials for fedex.com at line 850ish.

I won't update this code.
