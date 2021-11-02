/*====================================================================================================================================*
  Sheet2Form by Sandrine Muller
  ====================================================================================================================================
  Version:      1.0.0
  Project Page: https://github.com/sandrine-muller-research/GoogleApps/tree/main/FormGenerator
  License:      MIT License
  Doc:          https://github.com/sandrine-muller-research/GoogleApps/blob/main/FormGenerator/README.md
  ------------------------------------------------------------------------------------------------------------------------------------
  This Google App Script is creating a set of forms from normalized data in a google spreadsheet. The code will generate a form for each tab. 
  
  For bug reports see https://github.com/sandrine-muller-research/GoogleApps/issues
  ------------------------------------------------------------------------------------------------------------------------------------
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sheet2Form')
    .addItem('Export to form', 'main')
    .addToUi();
}

function main() {

  // get spreadsheet:
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  var file_sheet_dest = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var folders_sheet_dest = file_sheet_dest.getParents().next();

  for (var i=0 ; i<sheets.length ; i++){

    // get current sheet name:
    var form_title = sheets[i].getName()

    // create & name form:  
    var form = FormApp.create(form_title)  
       .setTitle(form_title);

    // parse sheet
    var data = parse_sheet(sheets[i]);

    // populate form:
    var qu=[], ans = [];
    for (var j=0 ; j<data.form_type.length ; j++){
      if (data.form_type[j]=="GRID"){
        // verify the list of choices is the same than previously:
        if ( arr_equals(ans,data.choices[j]) && (ans.length != 0) ) {
          // if ans is not empty
          qu.push(data.question[j]);
        }else if (ans.length == 0){
          ans = data.choices[j];
          qu.push(data.question[j]);
        }else{
          form.addGridItem()
              .setRows(qu)
              .setColumns(ans);
          qu = [];
          qu.push(data.question[j]);
          ans = data.choices[j];
          if (j == (data.form_type.length-1)){ // if last question on the form -> submit
            form = create_form_object(form, data.form_type[j], data.question[j], data.choices[j], data.required[j]);
          }
        }
      }else{
        if (ans.length != 0){// previous set of questions were GRID
          form.addGridItem()
              .setRows(qu)
              .setColumns(ans);
          qu = [];
          ans = [];
        }
        form = create_form_object(form, data.form_type[j], data.question[j], data.choices[j], data.required[j]);
      }
      
    }
    var file_form_source = DriveApp.getFileById(form.getId());
    var folders_form_source = file_form_source.getParents().next();
    moveFiles(file_form_source,folders_form_source, folders_sheet_dest); 

  }
};

function arr_equals(arrA,arrB){
  var out = true;
  var difference = arrA.filter(x => !arrB.includes(x));
  if ( (arrA.length==0) || (arrB.length==0) || (difference.length > 0)){
    out = false;
  }
  return out;
}

function moveFiles(file,source_folder, dest_folder) {
    dest_folder.addFile(file);
    source_folder.removeFile(file);
}

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
          .setBounds(1,5)
          .setLabels(choices[0], choices[1]);
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

    case 'DROPDOWN': 
      var item = form.addListItem()
      c = [];
      for (var i=0 ; i<choices.length ; i++){
        c.push(item.createChoice(choices[i]));
      }
    // Open a form by ID and add a new list item.
      item.setTitle(title_str)
      item.setChoices(c)
      item.setRequired(required);

      break;

    case 'TEXT_BOUNDED':
    // Add a text item to a form and require it to be a number within a range -> lower and upper bound in choice vector
      var textValidation = FormApp.createTextValidation()
        .requireNumberBetween(choices[0], choices[1]) //verify!!!
        .build();
      form.addTextItem()
        .setTitle(title_str)
        .setValidation(textValidation);
      break;

    case 'IMAGE': 
    // Open a form by ID and add a new list item.
      var img = DriveApp.getFileById(choices[0]);
      form.addImageItem()
          .setTitle(title_str)
          .setImage(img);
      break;

    default:
        // ignore item
  }
  return form;
}

function parse_sheet(sheet_name) {

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
    var d = data.map(row => row[i]);
    if (title == 'Field/Question'){question.push(d)}
    if (title == 'Required'){required.push(d);}
    if (title == 'Type'){form_type.push(d);} 
  });
  data.forEach((row) => {
    choices.push(row.slice(answer_start, (answer_end+1)).filter(e => e!=""));
  });

  out.form_type = form_type.flat(1);
  out.question = question.flat(1);
  out.choices = choices;
  out.required = required.flat(1);

  return out;
  
};
