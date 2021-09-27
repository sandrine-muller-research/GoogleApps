function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sheet2Form')
    .addItem('Export to form', 'main')
    .addToUi();
}

function main() {

  // get spreadsheet:
  //var sheets = SpreadsheetApp.openById('1rfKUQryiVVWB3PrjaPR0bE9Mi7yNqKjecWtLe8Q645w').getSheets();
  


  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++){

    var data = []; //question= [], required= [], choices = [], form_type = [];

    // get current sheet name:
    var form_title = sheets[i].getName()

    // create & name form:  
    var form = FormApp.create(form_title)  
       .setTitle(form_title);

    // parse sheet
    data = parse_sheet(sheets[i]);

    // populate form:
    for (var j=0 ; j<data.form_type.length ; j++){
      form = create_form_object(form, data.form_type[j], data.question[j], data.choices[j], data.required[j])
    }
    debugger;
  }
};

function create_form_object(form, casetype, title_str, choices, is_required) {

  if (is_required == "YES"){
    required = true;
  } else {
    required = false;
  }

  switch(casetype){
    case 'TEXT': 
      form.addTextItem()  
       .setTitle(title_str)  
       .setRequired(required);
      break;

    case 'SECTION':
      form.addPageBreakItem()
        .setTitle(title_str)  
        //.setRequired(is_required);
      break;

    case 'SCALE':
      form.addScaleItem()
          .setTitle(title_str)
          .setBounds(1, 5);
      break;

    case 'PARAGRAPH':  
      form.addParagraphTextItem()  
          .setTitle(title_str)  
          .setRequired(required);
      break;
      
    case 'RADIO':  
      form.addMultipleChoiceItem()  
          .setTitle(title_str)  
          .setChoiceValues(choices)  
          .setRequired(required);
      break;  
      
    case 'CHECKBOX':  
      form.addCheckboxItem()  
          .setTitle(title_str)  
          .setChoiceValues(choices)
          .setRequired(required);
      break;

    case 'MULTIPLE CHOICE':  
      form.addMultipleChoiceItem()  
          .setTitle(title_str)  
          .setChoiceValues(choices)
          .setRequired(required);
      break;

    case 'DROP DOWN': 
    // Open a form by ID and add a new list item.
      form.addListItem()
          .setTitle('Do you prefer cats or dogs?')
          .setChoices(choices)
          .setRequired(required);

    case 'TEXT_BOUNDED':
    // Add a text item to a form and require it to be a number within a range -> lower and upper bound in choice vector
      var textValidation = FormApp.createTextValidation()
        .requireNumberBetween(choices[0], choices[1]) //verify!!!
        .build();
      form.addTextItem()
        .setTitle(title_str)
        .setValidation(textValidation);

    default:
        // ignore item
  }
  return form;
}

function parse_sheet(sheet_name) {

  //var sheet = SpreadsheetApp.openById('1rfKUQryiVVWB3PrjaPR0bE9Mi7yNqKjecWtLe8Q645w');

  //var sheet_name = sheet.getSheetByName('Work-life')

  //var sheet_name = sheet.getSheetByName(sheet_name)

  const [header, ...data] = sheet_name
    .getDataRange()
    .getDisplayValues();

  const answer_start = header.indexOf('answer start');
  const answer_end = header.indexOf('answer end');
  var question = [];
  var required = [];
  var choices = [];
  var form_type = [];
  var out = {};
  header.forEach((title, i) => {
    const d = data.map(row => row[i]);
    if (title == 'Field/Question'){question.push(d)}
    if (title == 'Required'){required.push(d);}
    if (title == 'Types'){form_type.push(d);} 
  });
  data.forEach((row,i) => {
    choices.push(row.slice(answer_start, (answer_end+1)).filter(e => e!=""));
  });

  out.form_type = form_type.flat(1);
  out.question = question.flat(1);
  out.choices = choices;
  out.required = required.flat(1);

  //debugger;
  return out;
  
};

