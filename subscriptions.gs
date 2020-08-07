// Function to pull subscriptions from WooCommerce RestAPI
function subscription_sync() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_name = doc.getSheetByName('subscriptions').getName();
  
  fetch_subscriptions(sheet_name);
  
  // Sort data by date
  var sheet = doc.getSheetByName(sheet_name);
  var data = sheet.getRange(9,1,sheet.getLastRow() - 8,sheet.getLastColumn());
  data.sort({column: 7, ascending: true});
  
  // Reset date formula
  var cell = sheet.getRange("B6");
  cell.setFormula("=MAX(G:G)");
  
}


// function to fetch subscriptions from the WooCommerce API
function fetch_subscriptions(sheet_name) {
  var ck = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B4").getValue();
  var cs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B5").getValue();
  var website = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B3").getValue();
  var manualDate = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B6").getValue(); // Set your subscription start date in spreadsheet in cell B6
  var m = new Date(manualDate).toISOString();
  
  // Set the page you wish to begin on -- I normally only have to change this the first time I run the script to import ALL the subscription history
  var page = 1;
  
  var surl = website + "/wp-json/wc/v1/subscriptions?consumer_key=" + ck + "&consumer_secret=" + cs + "&after=" + m + "&per_page=100" + "&page=" + page; 
  
  var url = surl
  Logger.log(url)
  
  var options =
      {
      
        "method": "GET",
        "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
        "muteHttpExceptions": true,
        
      };
  
  var result = UrlFetchApp.fetch(url, options);
  
  if (result.getResponseCode() == 200) {
    
    var params = JSON.parse(result.getContentText());
    
  }
  
  page++;
  
  // as long as data is returned, loop through each page and parse the data for import
  while (params != "") {
    import_subscriptions_to_sheet(sheet_name, params);
    
    surl = website + "/wp-json/wc/v1/subscriptions?consumer_key=" + ck + "&consumer_secret=" + cs + "&after=" + m + "&per_page=100" + "&page=" + page;
    url = surl;
    result = UrlFetchApp.fetch(url, options);
    
    if (result.getResponseCode() == 200) {
      params = JSON.parse(result.getContentText());
      page++;
    }
    
  }
    
}


// Function to import subscriptions from JSON API response from server
function import_subscriptions_to_sheet(sheet_name, params) {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var temp = doc.getSheetByName(sheet_name);
  var consumption = {};
  var arrayLength = params.length;
  
  for (var i = 0; i < arrayLength; i++) {
    var a, c, e, d, f, x;
    var count = 1;
    var container = [];
    
    a = container.push(params[i]["id"]);
    a = container.push(count);
    a = container.push(params[i]["parent_id"]);
    a = container.push(params[i]["status"]);
    a = container.push(params[i]["order_key"]);
    a = container.push(params[i]["currency"]);
    a = container.push(new Date(params[i]["date_created"] + "-0500"));
    a = container.push(new Date(params[i]["date_modified"] + "-0500"));
    a = container.push(params[i]["customer_id"]);
    a = container.push(params[i]["billing"]["first_name"]);
    a = container.push(params[i]["billing"]["last_name"]);
    a = container.push(params[i]["billing"]["company"]);
    a = container.push(params[i]["billing"]["address_1"]);
    a = container.push(params[i]["billing"]["address_2"]);
    a = container.push(params[i]["billing"]["city"]);
    a = container.push(params[i]["billing"]["state"]);
    a = container.push(params[i]["billing"]["country"]);
    a = container.push(params[i]["billing"]["postcode"]);
    a = container.push(params[i]["billing"]["phone"]);
    a = container.push(params[i]["billing"]["email"]);
    a = container.push(params[i]["shipping"]["first_name"] + " "+ params[i]["shipping"]["last_name"]+" "+ params[i]["shipping"]["address_2"]+" "+ params[i]["shipping"]["address_1"]+" "+params[i]["shipping"]["city"]+" "+params[i]["shipping"]["state"]+" "+params[i]["shipping"]["postcode"]+" "+params[i]["shipping"]["country"]); 
    a = container.push(params[i]["discount_total"]);
    a = container.push(params[i]["discount_tax"]);
    a = container.push(params[i]["shipping_total"]);
    a = container.push(params[i]["shipping_tax"]);
    a = container.push(params[i]["cart_tax"]);
    a = container.push(params[i]["total"]);
    a = container.push(params[i]["total_tax"]);
    a = container.push(params[i]["created_via"]);
    a = container.push(params[i]["payment_method_title"]);
    a = container.push(params[i]["date_paid"]);
    a = container.push(params[i]["customer_note"]);
    a = container.push(params[i]["billing_period"]);
    a = container.push(params[i]["start_date"]);
    a = container.push(params[i]["end_date"]);
    a = container.push(params[i]["next_payment_date"]);
        
    // loop to parse coupon lines from the results
    e = params[i]["coupon_lines"].length;
    
    for (var s = 0; s < e; s++) {
      var coupon_code = params[i]["coupon_lines"][s]["code"];
      var coupon_discount = params[i]["coupon_lines"][s]["discount"];
      
    }
    
    // push coupon_lines to container for import
    a = container.push(coupon_code);
    a = container.push(coupon_discount);
        
    // Loop to parse line_items from the results
    c = params[i]["line_items"].length;
    
    for (var k = 0; k < c; k++) {
      var line_id, item, prod_id, qty, order_qty, price, total;
      
      line_id = params[i]["line_items"][k]["id"];
      qty = params[i]["line_items"][k]["quantity"];
      order_qty = qty;
      item = params[i]["line_items"][k]["name"];
      prod_id = params[i]["line_items"][k]["product_id"];
      price = params[i]["line_items"][k]["price"];
      total = params[i]["line_items"][k]["total"];
      
      // While loop to handle creating separate lines when quantity is greater than 1
      for (var q = 0; q < qty; q++) {
        var container2 = [];
        var qty_value = 1;
        
        // Push all data in container to container2
        for (x in container) {
          container2.push(container[x]);
        }
        
        // update count value in array
        container2[1] = count;
                
        // push individual line item to container for import
        a = container2.push(order_qty);
        a = container2.push(total);
        a = container2.push(line_id);
        a = container2.push(prod_id);
        a = container2.push(qty_value);
        a = container2.push(item);
        a = container2.push(price);
        a = container2.push(qty_value*price);
                
        // code to parse meta_data within each order line
        e = params[i]["line_items"][k]["meta"].length;
        
        var meta_occupancy = "";
        var meta_cabintype = "";
        var meta_payment = "";
        var meta_partypass = "";
        var meta_roomtype = "";
        
        for (var s = 0; s < e; s++) {
          var keys, value;
          
          keys = params[i]["line_items"][k]["meta"][s]["key"];
          value = params[i]["line_items"][k]["meta"][s]["value"];
          
          if (keys == "occupancy") {
            var meta_occupancy = value;
          } else if (keys == "cabin-type") {
            var meta_cabintype = value;
          } else if (keys == "payment-type") {
            var meta_payment = value;
          } else if (keys == "party-pass") {
            var meta_partypass = value;
          } else if (keys == "room-type") {
            var meta_roomtype = value;
          }
          
        }
        
        // push line_item meta_data to container for import
        a = container2.push(meta_occupancy);
        a = container2.push(meta_cabintype);
        a = container2.push(meta_payment);
        a = container2.push(meta_partypass);
        a = container2.push(meta_roomtype);
               
        // set default values if the order had greater than 1 quantity
        // this just removes extra data that I don't want, or need
        if (q > 0) {
          container2[9] = "TBD"; container2[10] = "TBD"; container2[11] = ""; container2[12] = ""; 
          container2[13] = ""; container2[14] = ""; container2[15] = ""; container2[16] = ""; 
          container2[17] = ""; container2[18] = ""; container2[19] = ""; container2[20] = ""; 
          container2[21] = ""; container2[22] = ""; container2[23] = ""; container2[24] = ""; 
          container2[25] = ""; container2[26] = ""; container2[27] = ""; container2[28] = ""; 
          container2[29] = ""; container2[30] = ""; container2[31] = ""; container2[32] = ""; 
          container2[33] = ""; container2[34] = ""; container2[35] = ""; container2[36] = ""; 
          container2[37] = ""; container2[38] = ""; container2[39] = "";
        } 
        
        // push data to the spreadsheet
        temp.appendRow(container2);
        
        // increment the line counter
        count++;
        
      }     
            
    }
    
  }
      
}
