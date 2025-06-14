function auto_fill_inscriptions(e) {
  /* Google Apps script created by Sandrine Muller on 2022-08-20
  Automatisation of inscriptions from forms
  Last edited, 2022-08-30 by Sandrine Muller 
  To fix: copy and deletion of justificatifs
  make sure IDs in datasheet reflect the new copied IDs
  when multiple justificatifs, only copying the first
  petite section ne renvoie pas 150*/

  // get event info:
  var d = e.namedValues;
  var keys = getKeys(d);
  var vals = getVals(d);
  flag = true;
  if(keys.length>4){
    var cnt = 0;
    for(var i = 1;i<68;i = i + 1){// check if inscription form changed
      if(vals[i] != [ '' ]){
        cnt = cnt + 1;
      }
    }
    var cnt2 = 0;
    for(var i = 1;i<68;i = i + 1){// check if attachements changed
      if(vals[i] != [ '' ]){
        cnt2 = cnt2 + 1;
      }
    }
  }else{
    var spt = SpreadsheetApp.openById('1RTHHxaFClsgfhU7aqrbPK8S4Q7mnEllPy_2t6VgiEs8');
    var ss = spt.getSheetByName("données familles");
    var values = ss.getDataRange().getValues();
    for(n=0;n<values.length;++n){
      var cell = values[n][1] ; // x is the index of the column starting from 0
      console.log(cell);
      console.log(d.Adresse);
      if(cell == d.Adresse){
        var rnn = n;
      }
    }
    flag = false;
  }

  // to allow modification, get event info:
  var spt = SpreadsheetApp.openById('1RTHHxaFClsgfhU7aqrbPK8S4Q7mnEllPy_2t6VgiEs8');
  var ss = spt.getSheetByName("données familles");
  var rn = e.range.getRow();
  var row = ss.getRange('A' + rn.toString() + ':' + 'CA'  + rn.toString()).getValues()[0];

  // get choices variables
  var date = Utilities.formatDate(row[0], spt.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  var email = row[1];
  var first_inscription = row[10];
  var fratrie = row[12];
  var QF = row[65];
  var nb_mois = row[66];
  var reglement_mode = row[67];
  var PS = row[11];
  var Enfant_nom = row[2];
  var Enfant_prenoms = row[3];

  // Get information from the current year:
  var ss = SpreadsheetApp.openById("1P5Gecu4ZMc6B-KNc6V_YcrFfNyGeK6Ql1D8KjwN-JyU");
  var start_year = ss.getRange("mensuel!B1").getValues();
  var end_year = ss.getRange("mensuel!C1").getValues();
  var min_tranche_1child = ss.getRange("mensuel!A4:A8").getValues();
  min_tranche_1child = min_tranche_1child;
  var max_tranche_1child = ss.getRange("mensuel!B4:B7").getValues();
  max_tranche_1child = max_tranche_1child;
  var value_tranche_1child = ss.getRange("mensuel!D4:D8").getValues();
  value_tranche_1child = value_tranche_1child;
  var min_tranche_children = ss.getRange("mensuel!A14:A18").getValues();
  min_tranche_children = min_tranche_children;
  var max_tranche_children = ss.getRange("mensuel!B14:B17").getValues();
  max_tranche_children = max_tranche_children;
  var value_tranche_children = ss.getRange("mensuel!D14:D18").getValues();
  value_tranche_children = value_tranche_children;
  var value_tranche_PS = ss.getRange("mensuel!D20").getValues();
  console.log(value_tranche_PS);
  var frais_inscription = ss.getRange("inscription!B2").getValues();
  var frais_materiel = ss.getRange("inscription!B4:C4").getValues(); // list, [0] : 1er enfant, [1] : 2+
  frais_materiel = frais_materiel[0];
  var frais_adhesion = ss.getRange("adhesion!B1").getValues();
  
  if(cnt>0 && flag){
    
    console.log("manage inscription");

    // get patterns:
    var ff = set_up_uploaded_files();
    var uploaded_files = ff[0];
    var uploaded_files_names = ff[1];
    var naming_var_in_template = set_up_words_patterns_in_template_file();

    // Create student subfolder:
    var folder_inscription_name = start_year[0].toString() + "-" + end_year[0].toString() + "_" + Enfant_nom.toUpperCase() + "_" + Enfant_prenoms;
    var parentFolderID = '10vVDPpXs4-vECL3JxONe36t52jpVon8w'
    var template_response_folder = getSubFolder(folder_inscription_name,parentFolderID);
    if (!template_response_folder){ // If childFolder is not defined it creates it inside the parentFolder
      parentFolder = DriveApp.getFolderById(parentFolderID);
      var template_response_folder = parentFolder.createFolder(folder_inscription_name);}
    else{// if folder exist, get all previous children to remove
      files_to_remove = template_response_folder.getFiles();
      file_to_remove_ids = [];
      while (files_to_remove.hasNext()){
        file_to_remove = files_to_remove.next();
        file_to_remove_ids.push(file_to_remove.getId());
      }
    }
    
    // response files:
    var template_file = DriveApp.getFileById('1-uHig5Se_3UuvmMevHWRULGAMgbENPkKjSmim8AjHEw');
    var template_file_reglement = DriveApp.getFileById('1JyquLjuJEwaur1MnIspnUets4YctdnFzjf12h37_dGE');
    var output_file_name = start_year[0].toString() + "-" + end_year[0].toString() + "_" + "Inscription" + "_" + Enfant_nom.toUpperCase() + "_" + Enfant_prenoms;
    var response_file = template_file.makeCopy(output_file_name,template_response_folder);
    var response_file_doc = DocumentApp.openById(response_file.getId());  
    var body = response_file_doc.getBody();
    
    var output_file_name_reglement = start_year[0].toString() + "-" + end_year[0].toString() + "_" + "Reglement" + "_" + Enfant_nom.toUpperCase() + "_" + Enfant_prenoms;
    var response_file_reglement = template_file_reglement.makeCopy(output_file_name_reglement,template_response_folder);
    var response_file_doc_reglement = DocumentApp.openById(response_file_reglement.getId());
    var body_reglement = response_file_doc_reglement.getBody();

    // fill template file with current year values:
    ////// Year:
    body.replaceText("{{Start_year}}",start_year);
    body.replaceText("{{End_year}}",end_year);

    ////// Tranches values:
    var T1_child = ["{{tranche1_1}}","{{tranche2_1}}","{{tranche3_1}}","{{tranche4_1}}","{{tranche5_1}}"];
    var T2_child = ["{{tranche1_2}}","{{tranche2_2}}","{{tranche3_2}}","{{tranche4_2}}"];
    var Tv_child = ["{{tranche1_value}}","{{tranche2_value}}","{{tranche3_value}}","{{tranche4_value}}","{{tranche5_value}}"];
    var T1_children = ["{{tranche1_1_fr}}","{{tranche2_1_fr}}","{{tranche3_1_fr}}","{{tranche4_1_fr}}","{{tranche5_1_fr}}"];
    var T2_children = ["{tranche1_2_fr}}","{tranche2_2_fr}}","{{tranche3_2_fr}}","{{tranche4_2_fr}}"];
    var Tv_children = ["{{tranche1_value_fr}}","{{tranche2_value_fr}}","{{tranche3_value_fr}}","{{tranche4_value_fr}}","{{tranche5_value_fr}}"];

    body.replaceText("{{tranchePS_value}}",PS.toString());
    for (var counter = 0; counter < T2_child.length; counter = counter + 1) {
      if(fratrie == "1er enfant"){
        if ((QF>=min_tranche_1child[counter]) && (QF<=max_tranche_1child[counter])){
          var frais_mensuels = value_tranche_1child[counter];
        }
      }else{
        if ((QF>=min_tranche_children[counter]) && (QF<=max_tranche_children[counter])){
          var frais_mensuels = value_tranche_children[counter];
        }
      }
      
      body.replaceText(T1_child[counter],min_tranche_1child[counter].toString());
      body.replaceText(T2_child[counter],max_tranche_1child[counter].toString());
      body.replaceText(Tv_child[counter],value_tranche_1child[counter].toString());
      body.replaceText(T1_children[counter],min_tranche_children[counter].toString());
      body.replaceText(T2_children[counter],max_tranche_children[counter].toString());
      body.replaceText(Tv_children[counter],value_tranche_children[counter].toString());

    }
    var counter = 4;
    if(fratrie == "1er enfant"){
      if ((QF>=min_tranche_1child[counter])){
        var frais_mensuels = value_tranche_1child[counter];
      }
    }else{
      if ((QF>=min_tranche_children[counter])){
        var frais_mensuels = value_tranche_children[counter];
      }
    }
    body.replaceText(T1_child[counter],min_tranche_1child[counter].toString());
    body.replaceText(Tv_child[counter],value_tranche_1child[counter].toString());
    body.replaceText(T1_children[counter],min_tranche_children[counter].toString());
    body.replaceText(Tv_children[counter],value_tranche_children[counter].toString());

    if (PS == "Oui"){
      frais_mensuels = value_tranche_PS[0];
      console.log("petite section OK");
    }
    console.log(frais_mensuels);

    ////// frais:
    body.replaceText("{{frais_inscription}}",frais_inscription.toString());
    body.replaceText("{{frais_materiel_1}}",frais_materiel[0].toString());
    body.replaceText("{{frais_materiel_2}}",frais_materiel[1].toString());
    body.replaceText("{{Montant_mensualite}}",nb_mois.toString());
    body.replaceText("{{Reglement_mode}}",reglement_mode);
    body.replaceText("{{Montant_adhesion}}",frais_adhesion.toString());
    var frais_annuels = frais_mensuels*10;
    if (nb_mois=="12 mois sans frais"){
      frais_mensuels = Math.round(frais_mensuels*10/12,4);
    }
    body.replaceText("{{Montant_frais_mensualite}}",frais_mensuels.toString());
    body.replaceText("{{Montant_frais_total}}",frais_annuels.toString());
    if(fratrie == "1er enfant"){
      body.replaceText("{{Montant_inscription}}",frais_inscription.toString());
    }else{
      body.replaceText("{{Montant_inscription}}","0");
    }
    if(first_inscription == "Oui"){
      body.replaceText("{{Montant_fournitures}}",frais_materiel[0].toString());
      var tot = frais_inscription[0][0] + frais_materiel[0] + frais_adhesion[0][0];
      body.replaceText("{{Montant_inscription_fournitures_adhesion}}",tot.toString());
    }else{
      body.replaceText("{{Montant_fournitures}}",frais_materiel[1].toString());
      var tot = frais_materiel[1] + frais_adhesion[0][0];
      body.replaceText("{{Montant_inscription_fournitures_adhesion}}",tot.toString());
    }

    var i = 0;
    for (var counter = 2; counter < 65; counter = counter + 1) {
      if((counter!=10)&&(counter!=11)&&(counter!=12)&&(counter!=43)){
        var tmp = row[counter];
        if (isValidDate(tmp)){
          tmp = Utilities.formatDate(tmp, spt.getSpreadsheetTimeZone(), "dd/MM/yyyy");
        }
        body.replaceText(naming_var_in_template[i],tmp.toString());
        if((counter==2)||(counter==3)||(counter==17)||(counter==18)||(counter==29)||(counter==30)){
          body_reglement.replaceText(naming_var_in_template[i],tmp.toString());
        }
        i = i + 1;
      }
      
    }
    body_reglement.replaceText("{{date}}",date.toString());

    response_file_doc_reglement.saveAndClose();
    response_file_doc.saveAndClose();

    // create pdfs:
    reglement_pdf = saveaspdf(response_file_doc_reglement,template_response_folder);
    inscription_pdf = saveaspdf(response_file_doc,template_response_folder);

    // send email with documents:
    sendEmailWithAttachment(Enfant_prenoms + " " + Enfant_nom,email,[inscription_pdf,reglement_pdf]);
    ss.getRange('BZ' + rn.toString()).setValues([['email envoye']]);
  }

  if(cnt2>0 && flag){

    
    console.log("manage piece jointes");

    // rename uploaded files and append to dossier d'inscription:
    for(var counter = 0; counter < uploaded_files.length; counter = counter + 1){
      if(!row[counter+68]){}else{
        if (Array.isArray(row[counter])){
          for(var i = 0; i < row.length; i = i + 1){
            var tmp = getIdFromUrl(row[counter][i]).toString();
            var folderId = DriveApp.getFileById(tmp).getParents().next().getId();
            if(tmp.length != 0){
              var uf = DriveApp.getFileById(tmp);
              var uff = DriveApp.getFolderById(folderId);
              var FileRename =  uf.makeCopy(start_year[0].toString() + "-" + end_year[0].toString() + uploaded_files_names[counter] + "_" + i.toString() + "_" + Enfant_nom.toUpperCase() + "_" + Enfant_prenoms,template_response_folder);
              uff.removeFile(uf);
            }
          }
        }else{
            var tmp = getIdFromUrl(row[counter+68]).toString();
            var folderId = DriveApp.getFileById(tmp).getParents().next().getId();
            var uf = DriveApp.getFileById(tmp);
            var uff = DriveApp.getFolderById(folderId);
            var FileRename =  uf.makeCopy(start_year[0].toString() + "-" + end_year[0].toString() + "_" + uploaded_files_names[counter] + "_" + Enfant_nom.toUpperCase() + "_" + Enfant_prenoms,template_response_folder);
            uff.removeFile(uf);
        }
      }
    }
    counter = 43;
    if (Array.isArray(row[counter])){
      for(var i = 0; i < row.length; i = i + 1){
        var tmp = getIdFromUrl(row[counter][i]).toString();
        var folderId = DriveApp.getFileById(tmp).getParents().next().getId();
        if(tmp.length != 0){
          var uf = DriveApp.getFileById(tmp);
          var uff = DriveApp.getFolderById(folderId);
          var FileRename =  uf.makeCopy(start_year[0].toString() + "-" + end_year[0].toString() + "_diplome_" + i.toString() + "_" + Enfant_nom.toUpperCase() + "_" + Enfant_prenoms,template_response_folder);
          uff.removeFile(uf);
        }
      }
    }else{
      var tmp = getIdFromUrl(row[counter]).toString();
      var folderId = DriveApp.getFileById(tmp).getParents().next().getId();
      if(tmp.length != 0){
        var uf = DriveApp.getFileById(tmp);
        var uff = DriveApp.getFolderById(folderId);
        var FileRename =  uf.makeCopy(start_year[0].toString() + "-" + end_year[0].toString() + "_diplome_" + Enfant_nom.toUpperCase() + "_" + Enfant_prenoms,template_response_folder);
        uff.removeFile(uf);
      }
    }
  }

  if(!flag){
    
    console.log("copy values");
    ss.getRange('CA' + rnn.toString()+ ':' + 'CB' + rnn.toString()).setValues([vals.splice(0, 2)]);
  }

  if(cnt>0 && flag){
    // remove previous files in student folder:
    if(file_to_remove_ids != undefined){
      for(var i=0;i<file_to_remove_ids.length;i=i+1) {
        file_to_remove = DriveApp.getFileById(file_to_remove_ids[i]);
        file_to_remove.setTrashed(true);
      }
    }
  }
  
}

//////////////////////////////////////////////////// UTILS
function getIdFromUrl(url) { return url.match(/[-\w]{25,}/);}

function getKeys(dict){
  var keys = [];
  for(var k in dict){
    keys.push(k);
  } 
  return keys;
}

function getVals(dict){
  var vals = [];
  for(var k in dict){
    vals.push(dict[k]);
  } 
  return vals;
}

function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}

function getSubFolder(childFolderName, parentFolderID){
  var parentFolder, parentFolders;
  var childFolder, childFolders;
  // Gets FolderIterator for parentFolder
  parentFolder = DriveApp.getFolderById(parentFolderID);

  // If parentFolder is not defined it sets it to root folder
  if (!parentFolder) { parentFolder = DriveApp.getRootFolder(); }

  // Gets FolderIterator for childFolder
  childFolders = parentFolder.getFoldersByName(childFolderName);
  /* Checks if FolderIterator has Folders with given name
  Assuming there's only a childFolder with given name... */ 

  while (childFolders.hasNext()) {
    childFolder = childFolders.next();
  }

  return childFolder;
}

function set_up_words_patterns_in_template_file (){
    var naming_var_in_template = ["{{Enfant_nom}}","{{Enfant_prenoms}}","{{Enfant_sexe}}","{{Enfant_date_naissance}}","{{Enfant_lieu_naissance}}","{{Enfant_nationalite}}","{{Enfant_domiciliation}}","{{Enfant_garde_precedent}}","{{Enfant_difficultes}}","{{Enfant_passions}}","{{Enfant_interet}}","{{Enfant_raisons_inscription}}","{{Responsable1_nom}}","{{Responsable1_prenom}}","{{Responsable1_sexe}}","{{Responsable1_date_naissance}}","{{Responsable1_lieu_naissance}}","{{Responsable1_nationalite}}","{{Responsable1_profession}}","{{Responsable1_adresse}}","{{Responsable1_telephone_maison}}","{{Responsable1_telephone_travail}}","{{Responsable1_email}}","{{Responsable1_situation_familiale}}","{{Responsable2_nom}}","{{Responsable2_prenom}}","{{Responsable2_sexe}}","{{Responsable2_date_naissance}}","{{Responsable2_lieu_naissance}}","{{Responsable2_nationalite}}","{{Responsable2_profession}}","{{Responsable2_adresse}}","{{Responsable2_telephone_maison}}","{{Responsable2_telephone_travail}}","{{Responsable2_email}}","{{Responsable2_situation_familiale}}","{{Parent_investissement_associatif_task}}","{{Parent_investissement_associatif_passion}}","{{Vaccin_Polio}}","{{Vaccin_ROR}}","{{Vaccin_autre}}","{{Allergie_alimentaire}}","{{Allergie_aliments}}","{{Allergie_medicamenteuse}}","{{Allergie_medicaments}}","{{Maladie_asthme}}","{{Maladie_asthme_medicaments}}","{{Traitement_medical_y_n}}","{{Traitement_medical}}","{{Medecin_nom}}","{{Medecin_prenom}}","{{Medecin_telephone}}","{{Medecin_adresse}}","{{Soin_autorisation}}","{{Autres_informations}}","{{Autorisation_image}}","{{Autorisation_transport}}","{{Autorisation_chercher}}","{{Autorisation_appeler}}"];

  return naming_var_in_template
}

function set_up_uploaded_files (){
  var uploaded_files = ["{{piece_livret}}","{{piece_jugement}}","{{piece_sante}}","{{piece_identite}}","{{piece_domicile}}","{{piece_assurance}}","{{piece_radiation}}","{{piece_dossier_scolaire}}","{{piece_CAF}}","{{piece_adhesion}}","{{piece_reglement}}"];  
  var uploaded_files_names = ["livret_famille","jugement_divorce","carnet_sante","identite","domiciliation","assurance","radiation_etablissement_scolaire","dossier_scolaire","quotient_familial","casiers_judiciaires"];

  return [uploaded_files,uploaded_files_names];
}

function saveaspdf(file,folder){
  docblob = file.getAs('application/pdf');
  /* Add the PDF extension */
  docblob.setName(file.getName() + ".pdf");
  var file_pdf = folder.createFile(docblob);

  return file_pdf;
}

function sendEmailWithAttachment(nom,email,files)
{
  var att = [];  
  for(i=0;i<files.length;i = i+1){
    var id = files[i].getId();
    var file = DriveApp.getFileById(id);
    att.push(file.getAs(MimeType.PDF))
  }

  var template = HtmlService.createTemplateFromFile('email-template-depot-doc');
  template.nom = nom;
  var message = template.evaluate().getContent();
  
  MailApp.sendEmail({
    to: email,
    subject: "2ème étape pour l'inscription à l'école Aurore Boréale",
    htmlBody: message,
    attachments: att
  });
  
}


