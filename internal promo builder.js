// initialization of the UI library instance
  var ui = SpreadsheetApp.getUi();
// initialization of needed variables 
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var buildersheet= spreadsheet.getSheetByName("Builder");
  var codessheet= spreadsheet.getSheetByName("Codes");

function generateCode() {
  var values = buildersheet.getRange(5,5,12,1); // (row, column, numRows, numColumns)  
  
  // get parts of the cp code
  var link = values.getCell(1,1).getValue().replace(/\s/g, "") ;
  var page_position = values.getCell(2,1).getValue();
  var carousel_position = values.getCell(3,1).getValue();
  var season = values.getCell(4,1).getValue();
  var lob = values.getCell(5,1).getValue();
  var gender = values.getCell(6,1).getValue();
  var bu = values.getCell(7,1).getValue();
  var category = values.getCell(8,1).getValue();
  var campaign = values.getCell(9,1).getValue();
  var landing_page_type = values.getCell(10,1).getValue();
  
  var freeform = values.getCell(11,1).getValue();
  var freeform_validation = values.getCell(12,1).getValue();
  
  // Check if compulsory values are not empty and in right format if not - throw error
  var errors = ["Please correct following problems and try generating the code again: \n"];
    //blanks
    var blanks = false;
    for (i = 2; i < 9; i++) { 
      if(values.getCell(i, 1).getValue() == ""){
        blanks = true;
      }
    }
    if(blanks == true){
      errors.push("- One of the required field is not filled.");
    }

    // Replace spaces with - hyphens (FC)
    if(campaign.toString().match(/ /)){
      campaign = campaign.replace(/ /g, '-');
    }
    if(freeform.toString().match(/ /)){
      freeform = freeform.replace(/ /g, '-');
    }
  
    // Free form check
    if(!freeform.toString().match(/^[a-zA-Z0-9\-]*$/)){
       errors.push(" - Free form contains prohibited characters. Please use just alphanumeric characters and hyphens (-).");
    }
    if(freeform.length > 40){
      errors.push(" - Text in the free form is too long. Please use max 35 chars.");
    }
    if(freeform_validation == "Yes" && freeform == ""){
      errors.push(" - Missing value in free form. Please provide free form value or uncheck the checkbox.");
    }
    //throw error if any
    if(errors.length > 1){
      ui.alert(errors.join("\n"))
    }
    
  // All values filled properly
  else{    
  
  // Generator logic
        // CP code
    var cp_code = page_position + carousel_position + ":" + season + ":" + lob + ":" + gender + ":" + bu + ":" + category + ":" + campaign + ":" + landing_page_type;
    
        if(freeform != ""){
          cp_code = cp_code + ":" + freeform;
        };
        
        var no_duplicates = checkDuplicate(cp_code);
        if(no_duplicates > 0){
          cp_code = cp_code + "_dupl" + no_duplicates;
        }
        
        // Link
        // Separate the hash part if any append the cp code and then #
        var hash_part_link = "";
        if(link.match(/#/)){
          link = link.match(/(.*)(#.*)/);
          link_without_hash = link[1];
          hash_part_link = link[2];
          link = link_without_hash;
        }
        if(link.match(/nike.com\/(xf\/|us\/|xl\/|ar\/|br\/|ca\/|mx\/|au\/|hk\/|in\/|id\/|my\/|nz\/|ph\/|sg\/|th\/|cn\/|tw\/|jp\/|kr\/|gb\/|be\/|cz\/|dk\/|de\/|es\/|fi\/|fr\/|gr\/|hu\/|ie\/|il\/|it\/|lu\/|nl\/|no\/|at\/|pl\/|pt\/|ru\/|si\/|se\/|ch\/|tr\/)/)){
          links = addLocales(link);
        }
        // Adding locale if there is none.
        // Fast workaround...
        else{
         var l;
         l = link.match(/(.*nike\.com)(.*)/);
         l = l[1] + "/en/gb" + l[2];
         links = addLocales(l);
        }
        
        // Append the CP code
        if(links[0].match(/\?/)){
          var links_with_cp_code = [];
          for(i=0; i<links.length; i++){
            links_with_cp_code.push(links[i] + "&intpromo=" + cp_code);
          }
          
        }
        else{     
          var links_with_cp_code = [];
          for(i=0; i<links.length; i++){
            links_with_cp_code.push(links[i] + "?intpromo=" + cp_code);
            
          }
          
        }
        // Append hash back if any
         var links_with_cp_code_and_hash = [];
          for(i=0; i<links_with_cp_code.length; i++){
            links_with_cp_code_and_hash.push(links_with_cp_code[i] + hash_part_link);
          }
        saveCode(links_with_cp_code[0], cp_code, no_duplicates, links_with_cp_code_and_hash);
      }
}

function generateCodeTop9() {
  var values = buildersheet.getRange(5,5,12,1); // (row, column, numRows, numColumns)  
  
  // get parts of the cp code
  var link = values.getCell(1,1).getValue().replace(/\s/g, "") ;
  var page_position = values.getCell(2,1).getValue();
  var carousel_position = values.getCell(3,1).getValue();
  var season = values.getCell(4,1).getValue();
  var lob = values.getCell(5,1).getValue();
  var gender = values.getCell(6,1).getValue();
  var bu = values.getCell(7,1).getValue();
  var category = values.getCell(8,1).getValue();
  var campaign = values.getCell(9,1).getValue();
  var landing_page_type = values.getCell(10,1).getValue();
  
  var freeform = values.getCell(11,1).getValue();
  var freeform_validation = values.getCell(12,1).getValue();
  
  // Check if compulsory values are not empty and in right format if not - throw error
  var errors = ["Please correct following problems and try generating the code again: \n"];
    //blanks
    var blanks = false;
    for (i = 2; i < 9; i++) { 
      if(values.getCell(i, 1).getValue() == ""){
        blanks = true;
      }
    }
    if(blanks == true){
      errors.push("- One of the required field is not filled.");
    }

    // Replace spaces with - hyphens (FC)
    if(campaign.toString().match(/ /)){
      campaign = campaign.replace(/ /g, '-');
    }
    if(freeform.toString().match(/ /)){
      freeform = freeform.replace(/ /g, '-');
    }
  
    // Free form check
    if(!freeform.toString().match(/^[a-zA-Z0-9\-]*$/)){
       errors.push(" - Free form contains prohibited characters. Please use just alphanumeric characters and hyphens (-).");
    }
    if(freeform.length > 40){
      errors.push(" - Text in the free form is too long. Please use max 35 chars.");
    }
    if(freeform_validation == "Yes" && freeform == ""){
      errors.push(" - Missing value in free form. Please provide free form value or uncheck the checkbox.");
    }
    //throw error if any
    if(errors.length > 1){
      ui.alert(errors.join("\n"))
    }
    
  // All values filled properly
  else{    
  
  // Generator logic
        // CP code
    var cp_code = page_position + carousel_position + ":" + season + ":" + lob + ":" + gender + ":" + bu + ":" + category + ":" + campaign + ":" + landing_page_type;
    
        if(freeform != ""){
          cp_code = cp_code + ":" + freeform;
        };
        
        var no_duplicates = checkDuplicate(cp_code);
        if(no_duplicates > 0){
          cp_code = cp_code + "_dupl" + no_duplicates;
        }
        
        // Link
        // Separate the hash part if any append the cp code and then #
        var hash_part_link = "";
        if(link.match(/#/)){
          link = link.match(/(.*)(#.*)/);
          link_without_hash = link[1];
          hash_part_link = link[2];
          link = link_without_hash;
        }
        if(link.match(/nike.com\/(xf\/|us\/|xl\/|ar\/|br\/|ca\/|mx\/|au\/|hk\/|in\/|id\/|my\/|nz\/|ph\/|sg\/|th\/|cn\/|tw\/|jp\/|kr\/|gb\/|be\/|cz\/|dk\/|de\/|es\/|fi\/|fr\/|gr\/|hu\/|ie\/|il\/|it\/|lu\/|nl\/|no\/|at\/|pl\/|pt\/|ru\/|si\/|se\/|ch\/|tr\/)/)){
          links = addLocalesTop9(link);
        }
        // Adding locale if there is none.
        // Fast workaround...
        else{
         var l;
         l = link.match(/(.*nike\.com)(.*)/);
         l = l[1] + "/en/gb" + l[2];
         links = addLocalesTop9(l);
        }
        
        // Append the CP code
        if(links[0].match(/\?/)){
          var links_with_cp_code = [];
          for(i=0; i<links.length; i++){
            links_with_cp_code.push(links[i] + "&intpromo=" + cp_code);
          }
          
        }
        else{     
          var links_with_cp_code = [];
          for(i=0; i<links.length; i++){
            links_with_cp_code.push(links[i] + "?intpromo=" + cp_code);
            
          }
          
        }
        // Append hash back if any
         var links_with_cp_code_and_hash = [];
          for(i=0; i<links_with_cp_code.length; i++){
            links_with_cp_code_and_hash.push(links_with_cp_code[i] + hash_part_link);
          }
        saveCode(links_with_cp_code[0], cp_code, no_duplicates, links_with_cp_code_and_hash);
      }
}

/////////////////////////support functions/////////////////////////

function addLocales(link_with_locale){
  // Parses the link to an array. 
  // Element with index 0 is the full link. 1 is the part before locale, 2 first part of locale - country (i. e. "/nl/"),
  // 3 second part of locale i. e. "en_gb", 4 is the part of the link behind locale
  
  parsed_link = link_with_locale.match(/(.*nike.com)\/(xf\/|us\/|xl\/|ar\/|br\/|ca\/|mx\/|au\/|hk\/|in\/|id\/|my\/|nz\/|ph\/|sg\/|th\/|cn\/|tw\/|jp\/|kr\/|gb\/|be\/|cz\/|dk\/|de\/|es\/|fi\/|fr\/|gr\/|hu\/|ie\/|il\/|it\/|lu\/|nl\/|no\/|at\/|pl\/|pt\/|ru\/|si\/|se\/|ch\/|tr\/)(.{5})(.*)/);
  //var link_without_locale = parsed_link[1] + parsed_link[4];
  
  var locales = [
                  '',
                  '/gb/en_gb',
                  '/be/en_gb',
                  '/be/nl_nl',
                  '/be/fr_fr',
                  '/be/de_de',
                  '/dk/en_gb',
                  '/de/de_de',
                  '/gr/el_gr',
                  '/es/es_es',
                  '/es/ca_es',
                  '/fi/en_gb',
                  '/fr/fr_fr',
                  '/ie/en_gb',
                  '/il/en_gb',
                  '/it/it_it',
                  '/lu/en_gb',
                  '/lu/fr_fr',
                  '/lu/de_de',
                  '/nl/nl_nl',
                  '/nl/en_gb',
                  '/pl/pl_pl',
                  '/no/en_gb',
                  '/at/de_de',
                  '/at/en_gb',
                  '/pt/en_gb',
                  '/se/en_gb',
                  '/ch/en_gb',
                  '/ch/fr_fr',
                  '/ch/de_de',
                  '/ch/it_it'
  ];
/*  var locales = [
    '',
    '/gb/en_gb',
    '/be/en_gb',
    '/be/nl_nl',
    '/be/fr_fr',
    '/be/de_de',
    '/dk/en_gb',
    '/cz/en_gb',
    '/de/de_de',
    '/gr/el_gr',
    '/es/es_es',
    '/es/ca_es',
    '/fi/en_gb',
    '/fr/fr_fr',
    '/hu/en_gb',
    '/ie/en_gb',
    '/il/en_gb',
    '/it/it_it',
    '/lu/en_gb',
    '/lu/fr_fr',
    '/lu/de_de',
    '/nl/nl_nl',
    '/nl/en_gb',
    '/no/en_gb',
    '/at/de_de',
    '/at/en_gb',
    '/pl/pl_pl',
    '/pt/en_gb',
    '/ru/ru_ru',
    '/si/en_gb',
    '/se/en_gb',
    '/ch/en_gb',
    '/ch/fr_fr',
    '/ch/de_de',
    '/ch/it_it',
    '/tr/tr_tr'];*/
  
  
  var links_with_locales = [];
  for(i=0; i<locales.length; i++){
    
    links_with_locales.push(parsed_link[1] + locales[i] + parsed_link[4]);
  }  
  return(links_with_locales);  
}

function addLocalesTop9(link_with_locale){
  // Parses the link to an array. 
  // Element with index 0 is the full link. 1 is the part before locale, 2 first part of locale - country (i. e. "/nl/"),
  // 3 second part of locale i. e. "en_gb", 4 is the part of the link behind locale
  
  parsed_link = link_with_locale.match(/(.*nike.com)\/(xf\/|us\/|xl\/|ar\/|br\/|ca\/|mx\/|au\/|hk\/|in\/|id\/|my\/|nz\/|ph\/|sg\/|th\/|cn\/|tw\/|jp\/|kr\/|gb\/|be\/|cz\/|dk\/|de\/|es\/|fi\/|fr\/|gr\/|hu\/|ie\/|il\/|it\/|lu\/|nl\/|no\/|at\/|pl\/|pt\/|ru\/|si\/|se\/|ch\/|tr\/)(.{5})(.*)/);
  //var link_without_locale = parsed_link[1] + parsed_link[4];
  
  var locales = [
                  '',
                  '/lu/en_gb',
                  '/lu/fr_fr',
                  '/lu/de_de',
                  '/it/it_it',
                  '/es/es_es',
                  '/es/ca_es',
                  '/be/nl_nl',
                  '/pl/pl_pl',
                  '/gr/el_gr'
  ];  
  
  var links_with_locales = [];
  for(i=0; i<locales.length; i++){
    
    links_with_locales.push(parsed_link[1] + locales[i] + parsed_link[4]);
  }  
  return(links_with_locales);  
}

// returns number of duplicates
function checkDuplicate(cp_code){
  var unique_check_cell = codessheet.getRange("G1");
  unique_check_cell.setFormula("=COUNTIF(E:E,\"" + cp_code +"\")");
  
  var unique_check = unique_check_cell.getValue();
  unique_check_cell.setValue("KEEP THIS CELL CLEAR");
  
  return unique_check;
}

function saveCode(cp_with_link, cp_without_link, no_duplicates, links_with_cp_code) {
  var user = Session.getEffectiveUser().getEmail(); //gets user's email
  
  // flag duplicates
  if(no_duplicates > 0){
      var cp_dedup = cp_without_link.match(/(.*)(_dupl.*)/);
      var save_duplicate_alert = ui.alert('You are trying to create following Internal Promo code, which already exists.\n'+ 
                                           cp_dedup[1] +
                                           '\n\n Are you sure want to create a duplicate? \n All previously created codes are stored in the codes sheet.', ui.ButtonSet.YES_NO);
      
      if(save_duplicate_alert == ui.Button.YES){
        codessheet.appendRow([new Date(), user, cp_with_link, cp_without_link, cp_dedup[1]]);
        var all_links = "";
        for(i=1; i<links_with_cp_code.length; i++){
          all_links = all_links + links_with_cp_code[i] + "\n";
        }
        
        var message = "Internal Promo code:\n\n" + cp_without_link + "\n\n" +      
                      "Links with the Internal Promo code and all locales:\n(Paste to Excel for better readability)\n\n" + all_links + "\n\n" +
                      "The link without locale and tracking code are also stored in the 'Codes' sheet for later usage.";
        ui.alert(message);
      }
  }
  else{
    codessheet.appendRow([new Date(), user, cp_with_link, cp_without_link, cp_without_link]);
      // Output - alert with the CP code link
      var all_links = "";
        for(i=1; i<links_with_cp_code.length; i++){
          all_links = all_links + links_with_cp_code[i] + "\n";
        }
        
        var message = "Internal Promo code:\n\n" + cp_without_link + "\n\n" +      
                      "Links with the Internal Promo code and all locales:\n(Paste to Excel for better readability)\n\n" + all_links + "\n\n" +
                      "The link without locale and tracking code are also stored in the 'Codes' sheet for later usage."; 
           ui.alert(message);
  } 
}
