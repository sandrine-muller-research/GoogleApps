# Sheet2Form

## What it does
This Google App Script is creating a set of forms from normalized data in a google spreadsheet. The code will generate a form for each tab. 

## Requirements
To auto-generate your forms you will need: 
1) a Google account and an explicit permission for the code to access the Google Spreadsheet and Drive
2) a Google Spreadsheet with 1 row per form item with the following headers:
    - Field/Question: text inserted for the question
    - Types	: normalized values for the time of item inserted. Please see below for a full description of the options
    - Description	: add a description for your fields
    - answer start the first header containing choices
    - answer end: the last header containing choices (if multiple questions with different number of choices, _answer end_ is the header of the column that contains the maximum  number of questions 
    - Required: YES or NO/empty cell
    - 
## Current form types supported
| Type label	| Description |
| ------------- | ----------- |
| SECTION	| add a section to a form. On preview, each section will be displayed at once on the webpage. You'll have to click "follwing" to move to the next section. |
| TEXT | a character limited text box | 
| PARAGRAPH | text box | 
| RADIO | radio button type answer | 
| SCALE | numerical scale | 
| CHECKBOX | check box button (you can choose as much as you want) | 
| MULTIPLE CHOICE | check box button (you can only choose one box) | 
| DROP DOWN | drop-down menu | 
| TEXT_BOUNDED | text field that is limited with a specific number | 
| IMAGE | image import. Provide in the answer field the image Google Drive ID. | 


## Usage
1) In your Google Sheet, go to _Tools_ > _Script Editor_ and add the _Sheet2Form.gs_ file. 
2) Reload the google sheet page
3) After a few second, you will see an additional tab called _Sheet2Form_ next to the _Help_ in your menu ribbon. 
4) Go to _Sheet2Form_ > _Export to form_
5) Once the generation is done, a pop up window will let you know. You will be able to see the forms under your Google Drive root.


