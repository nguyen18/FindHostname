// Script to find hostname based off the ip address using the Google DoH API
// link to article with functions to use API: https://medium.com/@aimhuge/reverse-dns-lookup-in-google-sheets-234c75966e55
// link to article used for sheet mapping: https://techandeco.medium.com/how-to-turn-a-google-sheet-into-a-simple-vacation-requesting-app-843e2c8bf2a4 

/**
 * Sheet header names
 */
const Header = {
  IpAddress: 'IP',
  Hostname: 'Hostname'
};

/**
 * Create menu button to find host name
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Actions')
    .addItem('Find Hostnames', 'mapping')  // create set time off button under schedule event category
    .addToUi();
}

/**
 * map out the rows of sheet to hostname column by name
 */
function mapping() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let dataRange = sheet.getDataRange().getValues();
  let headers = dataRange.shift();

  let rows = dataRange
    .map((row, i) => asObject(headers, row, i))
    .filter(row => row[Header.Hostname] == "")
    .map(findHostname)
    .map(row => writeRowToSheet(sheet, headers, row));
}


/**
 * Convert the row arrays into objects.
 * Start with an empty object, then create a new field
 * for each header name using the corresponding row value.
 * 
 * @param {string[]} headers - list of column names.
 * @param {any[]} rowArray - values of a row as an array.
 * @param {int} rowIndex - index of the row.
 */
function asObject(headers, rowArray, rowIndex) {
  return headers.reduce(
    (row, header, i) => {
      row[header] = rowArray[i];
      return row;
    }, { rowNumber: rowIndex + 1 });
}

/**
 * Rewrites a row into the sheet.
 * 
 * @param {SpreadsheetApp.Sheet} sheet - tab in sheet.
 * @param {string[]} headers - list of column names.
 * @param {Object} row - values in a row.
 */
function writeRowToSheet(sheet, headers, row) {
  let rowArray = headers.map(header => row[header]);
  let rowNumber = sheet.getFrozenRows() + row.rowNumber;
  sheet.getRange(rowNumber, 1, 1, rowArray.length).setValues([rowArray]);
}

/**
 * find the hostname using reverseLookup and return the row
 */
function findHostname(row) {
  row[Header.Hostname] = reverseLookup(row[Header.IpAddress]);
  return row;
}

/**
 * nslookup using Google DoH (DNS over HTTPS) API.
 *
 * @param {"google.com"} name    A well-formed domain name to resolve.
 * @param {"A"} type           Type of data to be returned, such as A, AAA, MX, NS...
 * @return {String}            Resolved IP address, or list of addresses separated by new lines
 * @customfunction
 */
function NSLookup(name,type) {
  var url = "https://dns.google.com/resolve?name=" + name + "&type=" + type;
  var response = UrlFetchApp.fetch(url);
  var responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    throw new Error( responseCode + ": " + response.message );
  }
  var responseText = response.getContentText(); // Get the response text
  var json = JSON.parse(responseText); // Parse the JSON text
  var answers = json.Answer.map(function(ans) {
    return ans.data
  }).join('\n'); // Get the values
  return answers;
}

/**
 * reverse lookup using Google DoH (DNS over HTTPS) API.
 *
 * @param {"1.1.1.1"} ip    An ip address to lookup.
 * @return {String}         Resolved Fully Qualified Domain Name
 * @customfunction
 */
function reverseLookup(ip) {
  ip = ip.split(".").reverse().join(".")
  
  var url = "https://dns.google.com/resolve?name=" + ip + ".in-addr.arpa&type=PTR";
  var response = UrlFetchApp.fetch(url);
  var responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    throw new Error( responseCode + ": " + response.message );
  }
  var responseText = response.getContentText(); // Get the response text
  var json = JSON.parse(responseText); // Parse the JSON text
  try{
    return json.Answer ? json.Answer[0].data : "no data" ;
}catch(err){
    return "no data: " + err ;
  }
}

/**
 * reverse lookup using Google DoH (DNS over HTTPS) API.
 *
 * @param {"1.1.1.1\n2.2.2.2"} ip    A new line separated list of IPs
 * @return {String}         Resolved List of Fully Qualified Domain Names
 * @customfunction
 */
function reverseLookups(ip){
  ips = ip.split('\n')
result = ips.map(function(ip) {
    var res = reverseLookup(ip);
    return res || "no data"
  }).join('\n');
  
  return result;
}
