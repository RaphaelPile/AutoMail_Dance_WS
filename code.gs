function email() {

// Récupère les données

var ss = SpreadsheetApp.getActiveSpreadsheet()
var rep_form = ss.getSheetByName("Réponses au formulaire 1") ;
//var template_mail = ss.getSheetByName("Template mails") ;
var der_ligne = rep_form.getLastRow() ;
var now = new Date();
var str_dm_confirm = "Date mail confirm" ;
var str_dm_paid = "Date mail paiement" ;
var str_dm_att = "Date mail liste attente" ;
var str_dm_part = "Date mail attente part" ;
//var prix = "15€" ;

/*
msg_paiement = getCellRangeByLineName(template_mail,"Paiement",2)
msg_conf_la = getCellRangeByLineName(template_mail,"Confirmation liste attente",2)
msg_conf_part = getCellRangeByLineName(template_mail,"Confirmation partenaire",2)
msg_conf = getCellRangeByLineName(template_mail,"Confirmation",2)
msg_RIB= getCellRangeByLineName(template_mail,"RIB",2)
msg_Relance = getCellRangeByLineName(template_mail,"Relance",2)
msg_la = getCellRangeByLineName(template_mail,"Liste attente",2)
msg_att_part = getCellRangeByLineName(template_mail,"Attente partenaire",2)
*/


var is_send_new_mail = false ;

for (var i=2; i < der_ligne+1; i++){
  var message = "";

  name = getCellValueByColumnName(rep_form, "Nom", i)
  surname = getCellValueByColumnName(rep_form, "Prénom", i)
  role = getCellValueByColumnName(rep_form, "Quel est ton rôle (pour cette séance) ?", i)
  part = getCellValueByColumnName(rep_form, "Nom du partenaire d'inscription, si tu t'inscris en binôme", i)
  prix = getCellValueByColumnName(rep_form, "Prix", i)

  is_confirm = getCellValueByColumnName(rep_form, "Confirmation", i)
  is_mail_confirm = getCellValueByColumnName(rep_form, str_dm_confirm, i)
  date_mail_confirme = Date(is_mail_confirm)

  is_paid = getCellValueByColumnName(rep_form, "Paiement", i)
  is_mail_paid = getCellValueByColumnName(rep_form, str_dm_paid, i)

  is_latt = getCellValueByColumnName(rep_form, "Liste attente", i)
  is_mail_latt = getCellValueByColumnName(rep_form, str_dm_att, i)

  is_part = getCellValueByColumnName(rep_form, "Attente partenaire", i)
  is_mail_part = getCellValueByColumnName(rep_form, str_dm_part, i)
  
  message += "Bonjour " + surname + " " + name + ",\n";
  message += "\n";

  if (is_paid == "OK"){
    if (is_mail_paid == ""){
      //message += getCellValueByLineName(template_mail,"Paiement",2)
      message += "Nous avons bien reçu le paiement.\n"
      message += "Ton inscription est validée !\n" ;
      setCellValueByColumnName(rep_form, str_dm_paid, i, now)
      is_send_new_mail = true ;
    }
  }
  else if (is_confirm == "OK"){
    if (is_mail_confirm == ""){

      if (is_part == "OK"){
        message += "Nous avons bien reçu l'inscription de ton/ta partenaire." + "\n";
        message += "Ton inscription en tant que " + role + " est donc enregistrée." + "\n";
      }
      else if (is_latt == "OK"){
        message += "Bonne nouvelle ! Tu es sorti.e de la liste d'attente." + "\n";
        message += "Ton inscription en tant que " + role + " est donc enregistrée." + "\n";
      }
      else{
        message += "Nous avons bien enregistré ton inscription en tant que " + role + " .\n";
      }

      message += "\n";
      message += "Pour valider ton inscription, merci d'adresser dès que possible un virement de " + prix + "€  à l'ordre de :\n"
      message += "West In Lille - IBAN : FR62 1144 9000 0101 3897 1001 Z83 - BIC : BDEIFRPPXXX\n"
      message += "Si possible avec ' " + name + " COACHING' comme libellé du virement. \n"

      setCellValueByColumnName(rep_form, str_dm_confirm, i, now)
      is_send_new_mail = true ; 
    }
    else if (parseInt((now-is_mail_confirm)/(24*3600*1000)) > 4){ // Tous les 5j
      message += "Nous n'avons pas encore reçu le paiement de ton inscription.\n";
      message += "Merci d'adresser au plus vite un virement de " + prix + " à l'ordre de :\n";
      message += "West In Lille - IBAN : FR62 1144 9000 0101 3897 1001 Z83 - BIC : BDEIFRPPXXX\n" ;
      message += "Si possible avec ' " + name + " COACHING' comme libellé du virement. \n" ;
      setCellValueByColumnName(rep_form, str_dm_confirm, i, now) ;
      is_send_new_mail = true ; 
    }
  }
  else if (is_latt == "OK"){
    if (is_mail_latt == ""){
      message += "Tu es actuellement dans la liste d'attente du rôle " + role + ".\n" ;
      message += "On te fait signe dès qu'il est possible de t'inscrire!\n"
      message += "Néanmoins, on te recommande de chercher un.e partenaire d'inscription de ton côté. \n"
      message += "Si tu trouves, demande bien de préciser ton nom dans la case 'Nom du partenaire d'inscription'. \n"

      setCellValueByColumnName(rep_form, str_dm_att, i, now)
      is_send_new_mail = true ; 
    }
  }

  else if (is_part == "OK"){
    if (is_mail_part == ""){
      message += "Nous attendons que ton/ta partenaire (" + part + ") remplisse aussi le formulaire pour enregistrer ton inscription.\n" ;
      message += "On te fait signe dès que c'est reçu ! \n"

      setCellValueByColumnName(rep_form, str_dm_part, i, now)
      is_send_new_mail = true ; 
    }

  }
  
  if (is_send_new_mail){
    let mail = rep_form.getRange(i,4).getValue()
    //let mail = "raphael.of.p@gmail.com" ; // Pour test, evite d'envoyer au stagiaire pour l'instant

    message += "\n";
    message += "A bientôt,\n"
    message += "L'équipe West In Lille"
    GmailApp.sendEmail(mail,"Coaching perfectionnement 2",message);
    is_send_new_mail = false ; //reset
  }

}
}

function getCellValueByColumnName(sheet, columnName, row) {
  let cell = getCellRangeByColumnName(sheet, columnName, row);
  if (cell != null) {
    return cell.getValue();
  }
}

function getCellRangeByColumnName(sheet, columnName, row) {
  let data = sheet.getDataRange().getValues();
  let column = data[0].indexOf(columnName);
  if (column != -1) {
    return sheet.getRange(row, column + 1, 1, 1);
  }
}

function getCellValueByLineName(sheet, lineName, col) {
  let cell = getCellRangeByLineName(sheet, lineName, col);
  if (cell != null) {
    return cell.getValue();
  }
}

function getCellRangeByLineName(sheet, lineName, col) {
  let data = sheet.getDataRange().getValues();
  transposedata = Transpose(data);
  let row = transposedata[0].indexOf(lineName);
  if (row != -1) {
    return sheet.getRange(row + 1, col, 1, 1);
  }
}

function Transpose(a){
  return Object.keys(a[0]).map( function (c) { return a.map(function (r) { return r[c];}); });
}

function setCellValueByColumnName(sheet, columnName, row, new_value) {
  let cell = getCellRangeByColumnName(sheet, columnName, row);
}
