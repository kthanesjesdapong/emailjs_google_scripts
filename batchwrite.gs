var search_query = 'QUERY PARAMS';
var header_fields = ['First Name', 'Last Name', 'Email', 'Phone Number']
var sheet_name='NAME OF GOOGLE SHEET'

//Clears current sheets and sets headers for them
function setHeaderConfigs () {
    console.log('Clearing sheets...');
  var currentSheet = SpreadsheetApp.getActiveSheet();
  currentSheet.clear();
  //Styling for headers
  var header_range = currentSheet.getRange("A1:D1");
  var header_style = SpreadsheetApp.newTextStyle()
    .setForegroundColor("black")
    .setFontSize(14)
    .setBold(true)
    .build();
  header_range.setTextStyle(header_style);
  header_range.setHorizontalAlignment('center')
  currentSheet.appendRow(header_fields)
  insertCustomerContact()
  return currentSheet;
}

//Returns an Array of Thread Bodies
function getThreadBodies () {
    var threads = GmailApp.search(search_query, 0 , 10);
    const threadBodies = [];
    for (let i = 0 ; i < threads.length ;i++) {
      const threadTable = getMessageBody(threads[i],i);
      threadBodies.push(threadTable);
    }
    return threadBodies
}

//Returns an Array of CustomerContacts
function getCustomers(threadBodies) {
    const customers = [];
    for (let j = 0 ; j < threadBodies.length; j++) {
      var field = threadBodies[j];
      const fieldTd = getTdTexts(field)
      const rowFieldTd = [getCustomerContact(fieldTd)]
      customers.push(rowFieldTd);
    }
  return customers;
}

//Calls, getCustomers, and threadBodies and inserts within our spreadsheet
function insertCustomerContact() {
    var currentSheet = SpreadsheetApp.getActiveSheet()
    //An Array of threadBodies
    const threadBodies = getThreadBodies();
    //An Array Containing Customer Contacts, where customers[i] is the customerContact of the Ith customer
    const customers = getCustomers(threadBodies)
    for (const customer of customers) {
      for (const customerContact of customer) {
        currentSheet.appendRow(customerContact);
      }
    }
  }

//Gets the table from an email
//Couldn't use xml parser so I had to parse through tbody instead
function getMessageBody(thread, index) {
  var message = thread.getMessages()[0];
  var body = message.getBody();
  var indexOrigin = body.search('<table');
  var indexEnd = body.search('</table')
  var messageTable = body.substring(indexOrigin,indexEnd+8);
  return messageTable
}

//Grabs Fields from td
function getTdTexts (subStr) {
    let temp = ''
    let tempArr = []
    let open = false;
    for (let i = 0 ; i < subStr.length ; i++) {
      if (subStr[i] === ' ' && subStr[i-1] == ';' && subStr[i-2] == 'p') {
        open = true;
        closed = false;
      }
      if (subStr[i] ==='<' && subStr[i+1] == '/') {
        open = false;
      }
      if (open === true && subStr[i] !== ' ' && subStr[i] !== ':') {
        temp+= subStr[i];
      }
      if (open === false && temp.length > 0) {
        tempArr.push(temp);
        temp = '';
      }
    }
    return tempArr;
  }

//Grabs only pertinent fields from tdTexts
//Might have to change this function depending on the table youre grabbing information from
function getCustomerContact (arrOfFields) {
  const res = [];
  for (let i = 1 ; i < arrOfFields.length ;i+=2) res.push(arrOfFields[i]);
  return res;
}
