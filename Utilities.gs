function main() {
    // Get utility costs and venmo/zelle payments from emails 
    costsMain();
    paymentsMain();
    
}
function getFirstDateOfQuarter(date) {
    // Get the first date in a quarter
    const quarter = Math.ceil((date.getMonth() + 1) / 3);
    const year = date.getFullYear();

    switch (quarter) {
        case 1:
            return new Date(year, 0, 1); // January 1st
        case 2:
            return new Date(year, 3, 1); // April 1st
        case 3:
            return new Date(year, 6, 1); // July 1st
        case 4:
            return new Date(year, 9, 1); // October 1st
    }
}

function dateString(date) {
    // wrapper for local date string
    return date.toLocaleDateString();
}
function costsMain() {
    saveCoststoSheet(getUtilityCosts());
}
function getUtilityCosts() {
    // get the cost of utilities for all emails under the below label
    const app = GmailApp;
    // Create labels in your gmail or update this to a different search text
    const threads = app.search("label:utilities");
    // threads contain messages

    let today = new Date();

    // Example of hardcoded expenses
    // SEWAGE $99 per quarter
    let startOfQuarter = getFirstDateOfQuarter(today).toLocaleDateString();
    let sewage = { "UTILITY": "Sewage", "AMOUNT": '99.00', "Dates": { "RECEIVED_DATE": startOfQuarter, "DUE_DATE": startOfQuarter } };
    let data = [];
    // Example
    if (today.getFullYear() == 2024) {
        data.push(sewage);
    }

    for (var i = 0; i < threads.length; i++) {
        // If you only need the first message otherwise add another for loop
        let messages = threads[i].getMessages();
        let subject = messages[0].getSubject();
        let content = messages[0].getPlainBody();
        // email receipt date
        let receivedDate = messages[0].getDate().toLocaleDateString();
        // Utility name REGEX - modify the names based on the companies you work with
        const utilityExpr = /(Edison|SCE|Starlink|Verizon|Internet|Waste|Sewage)/mi;
        // regex for dates preceeded by the word Date and optionally colon or \s chars
        const dateExpr = /([a-zA-Z ]+Date)(?:[: \t\n\r\f\v]{1,2})(\d{1,2}\/\d{1,2}\/\d{2,4})/gm
        // regex for US dollar amounts 
        const amountExpr = /(?:\$\d+\.\d{0,2})/m;

        let cleanDates = { "RECEIVED_DATE": receivedDate };
        if (content) {
            //   Log content for debugging
            // console.log(content);
            // string.match returns matches and or groups
            let dollarAmount = content.match(amountExpr);
            if (!dollarAmount) {
                // Not all dollar amounts are formatted as currency
                dollarAmount = content.match(/(?<!Due:)\d+\.\d{0,2}/m);

            }
            // Find the utility in the email body or subject
            let utility = content.match(utilityExpr);
            if (!utility) {
                utility = subject.match(utilityExpr);
            }
            // Custom logic for a different date format
            if (utility[0] == "Internet") {
                console.log(utility[0])
                let earthDateExpr = /(\d{1,2}\/\d{1,2}\/\d{2,4})/gm;
                let dates = content.match(earthDateExpr);
                console.log(dates);
                if (dates) {
                    cleanDates["Invoice Date"] = dates[0];
                    cleanDates["Due Date"] = dates[1];
                }
            } else {
                // Some utility emails contain multiple dates
                let dates = content.match(dateExpr);

                if (dates) {
                    // console.log(dates);
                    let groups = dates.groups;
                    if (groups) {
                        // console.log("Has groups",groups);

                        cleanDates[groups[0].trim()] = groups[1];
                    } else {
                        // Separating the date label from the date it is due
                        let splitChars = [":", "Date", " "];
                        for (let d of dates) {
                            for (let splitter of splitChars) {
                                if (d.search(splitter) > 0) {
                                    let splitDate = d.split(splitter);
                                    if (splitter === "Date") {
                                        cleanDates[splitDate[0].trim() + ` ${splitter}`] = splitDate[1].trim();
                                    } else {
                                        cleanDates[splitDate[0].trim()] = splitDate[1].trim();
                                    }

                                    break
                                }

                            }
                        }
                    }
                }
            }
            // append map/object to data array
            data.push({ "UTILITY": utility[0], "AMOUNT": dollarAmount[0].replace("$", ""), "Dates": cleanDates });


        }
    }
    // console.log(data)
    return data
}
function saveCoststoSheet(data) {
    // Save json data to SpreadSheet
    console.log(data)
    const drive = DriveApp;
    const spreadsheet = SpreadsheetApp;
    const spreadSheetName = "Expenses";
    // Mapping fields for SQL tables in BigQuery
    const fieldMap = { "Payment Due Date": "DUE_DATE", "Due Date": "DUE_DATE", "Invoice Date": "INVOICE_DATE", "Statement Date": "INVOICE_DATE" };
    // Drive search for filename
    let exists = drive.searchFiles(`title contains "${spreadSheetName}"`);

    // Get or create spreadsheet if it does not exist
    let newSheet = getSpreadsheet(spreadSheetName, true);
    const activeSheet = newSheet.getSheetByName("Utilities"); //newSheet.getActiveSheet();
    activeSheet.activate();
    if (sheetIsBlank(activeSheet)) {
        // Set default field values 
        activeSheet.getRange("A1:C1").setValues([["UTILITY", "AMOUNT", "RECEIVED_DATE"]]);

    }

    let today = new Date();

    for (let obj of data) {
        // iterate through utilities
        let recDate = new Date(obj["Dates"]["RECEIVED_DATE"]);

        // if email receipt date is in the past check if it has been recorded
        // all records should be checked for duplicates
        // text finder searches for the give text within the sheet
        let uteFinder = activeSheet.createTextFinder(obj["UTILITY"]);

        // find all occurences of text
        let uteRange = uteFinder.findAll();

        let recordExists = false;
        for (let rng of uteRange) {
            // received date is two columns to the left of the utility
            let offsetDate = rng.offset(0, 2).getValue();
            // console.log(rng.offset(0, 2).getA1Notation(), offsetDate, obj["UTILITY"], recDate, offsetDate.valueOf() == recDate.valueOf())

            if (offsetDate.valueOf() == recDate.valueOf()) {
                // if the record exists break out of loop
                recordExists = true;

                break
            }
        }
        if (recordExists) {
            // if the record exists continue to the next utility
            continue
        }

        // next blank row for writing data
        let lastRow = activeSheet.getLastRow() + 1;
        // Write utility and amount in the next row
        activeSheet.getRange(lastRow, 1).setValue(obj["UTILITY"]);
        activeSheet.getRange(lastRow, 2).setValue(obj["AMOUNT"]);
        let dateArray = obj["Dates"];
        // iterate over the dates, writing them in the appropriate column
        for (let dateKey in dateArray) {
            findColumnInsertValue(activeSheet, lastRow, dateKey in fieldMap ? fieldMap[dateKey] : dateKey, dateArray[dateKey])
        }


    }
    spreadsheet.flush()
}

function findColumnInsertValue(activeSheet, row, columnName, value) {
    // Check if columnName exists, creating the column if it doesn't and add new rows of value. Does not check for duplicates
    console.log("row: ", row, "colName: ", columnName, "val: ", value)
    let textFinder = activeSheet.createTextFinder(columnName);
    // find the column 
    let found = textFinder.findNext();
    // console.log(found);
    if (found) {
        let foundCol = found.getColumn();

        let actRange = activeSheet.getRange(row, foundCol);
        if (!actRange.getValue()) {
            actRange.setValue(value);
        } else {
            while (actRange.getValue()) {
                actRange = actRange.offset(1, 0);
            }
            actRange.setValue(value);
        }
    } else {
        let lastCol = sheetIsBlank(activeSheet) ? 1 : activeSheet.getLastColumn() + 1;

        row = 1;
        let actRange = activeSheet.getRange(row, lastCol);
        actRange.setValue(columnName);
        actRange.offset(1, 0).setValue(value);
    }
}

function paymentsMain() {
    // Main payments function
    savePayments(getPayments());
}



function getPayments() {
    // get all venmo utility payments
    const app = GmailApp;
    const threads = app.search("label:payments");

    // threads contain messages
    let today = new Date();
    let data = [];
    for (let thread of threads) {

        let messages = thread.getMessages();
        for (let message of messages) {
            let subject = message.getSubject();
            let content = message.getPlainBody();
            let receivedDate = message.getDate().toLocaleDateString();
            // Zelle or Venmo payments, add more if you use something else
            let source = subject.match("Zelle") ? "Zelle" : "Venmo";
            let tranId = "";
            let amount = "";
            // Venmo logic
            if (source == "Venmo") {
                // transaction ID / confirmation number
                tranId = content.match(/Transaction\sID\s{1,2}(\d+)/m)[1]; // Groups at index 1
                amount = subject.split("$")[1];
            } else {
                // Zelle logic
                // Replace Landlord with the person you pay on Zelle
                if (!content.match(/Landlord/m)) {
                    continue;
                }
                // transaction ID / confirmation number
                tranId = content.match(/Confirmation:\s\*(\w+)/m)[1]; // Groups at index 1
                // console.log(tranId);
                amount = content.match(/\$([\d.]+)/m)[1]; // Groups at index 1


            }
            // REGEX to identify if the email is for utlitiy or rent payments. Change as needed
            let paymentType = content.match(/Utilities|Rent/)[0];
            data.push({ "PAYMENT_DATE": receivedDate, "PAYMENT_ID": tranId, "PAYMENT_TYPE": paymentType, "AMOUNT": amount });
        }

    }
    // console.log(data);
    return data
}

function savePayments(data) {
    // Store venmo/zelle payments in Spreadsheet
    console.log(data)
    const drive = DriveApp;
    const spreadsheet = SpreadsheetApp;
    const spreadSheetName = "Expenses";
    // get or create spreadsheet if it does not exist
    let newSheet = getSpreadsheet(spreadSheetName, true);

    let sheetName = "Payments";
    let activeSheet = newSheet.getSheetByName(sheetName);
    if (!activeSheet) {
        activeSheet = newSheet.insertSheet(sheetName);
    }
    activeSheet.activate();
    let lastRow = 0;
    // Add column headers if they don't exist
    if (sheetIsBlank(activeSheet)) {
        // sheet is blank
        activeSheet.getRange("A1:D1").setValues([["PAYMENT_DATE", "PAYMENT_ID", "PAYMENT_TYPE", "AMOUNT"]]);
        lastRow = 1;
    } else {
        lastRow = activeSheet.getLastRow() + 1;
    }
    // next blank row for writing data
    for (let obj of data) {
        lastRow = sheetIsBlank(activeSheet) ? 1 : activeSheet.getLastRow() + 1;
        let textFinder = activeSheet.createTextFinder(obj["PAYMENT_ID"]);
        let found = textFinder.findNext();
        if (found) {
            // if PAYMENT_ID exists continue to next object
            continue
        }
        for (let key in obj) {
            // Write data to sheet
            findColumnInsertValue(activeSheet, lastRow, key, obj[key]);
        }
    }
}
function getSpreadsheet(spreadSheetName, createNew = false) {
    // Get Spreadsheet by name optionally create a new one
    const drive = DriveApp;
    const spreadsheet = SpreadsheetApp;
    // standardize date fields
    let exists = drive.searchFiles(`title contains "${spreadSheetName}"`);
    // Drive search for filename
    let newSheet = {};
    // create spreadsheet if it does not exist
    if (exists.hasNext()) {
        let file = exists.next();
        if (file.getName() == spreadSheetName) {
            newSheet = spreadsheet.open(file);
            console.log("Spreadsheet exists");
            return newSheet;
        }
    } else {
        if (createNew) {
            newSheet = spreadsheet.create(spreadSheetName);
            console.log(`Created new sheet "${spreadSheetName}"`);
            return newSheet;
        } else {
            console.log("Spreadsheet was not found");

        }
    }
    return null;
}
function getOrCreateSheet(spreadsheet, sheetName) {
    // Get or create a Sheet by name in a Spreadsheet
    let activeSheet = spreadsheet.getSheetByName(sheetName);
    if (!activeSheet) {
        activeSheet = spreadsheet.insertSheet(sheetName);
    }
    activeSheet.activate();
    return activeSheet;
}
function sheetIsBlank(activeSheet) {
    // If A1 is blank the entire sheet is considered blank
    if (activeSheet.getRange("A1").isBlank()) {
        return true
    }
    return false
}
function saveRent() {
    // Store upcoming rent costs in Rent sheet
    let rentAmount = 1000.00;
    let data = [];
    let today = new Date();
    let firstDate = new Date("12/1/2024");
    let year = today.getFullYear();
    let month = today.getMonth(); // RETURNS 0 INDEXED MONTHS
    // console.log(month);
    let nextYear = month == 11 ? year + 1 : year; // RETURNS 0 INDEXED MONTHS
    let nextMonth = nextYear == year ? month + 1 : 1;
    let thisMonth = new Date(year, month, 1);
    // console.log(month,thisMonth);
    data.push({ "DUE_DATE": dateString(firstDate), "AMOUNT": rentAmount });
    data.push({ "DUE_DATE": dateString(thisMonth), "AMOUNT": rentAmount });
    if (today > firstDate) {
        for (let i = 0; i <= nextMonth; i++) {
            let nextDate = new Date(nextYear, i, 1);
            if (nextDate > firstDate) {
                data.push({ "DUE_DATE": dateString(nextDate), "AMOUNT": rentAmount });

            }
        }

    }
    // console.log(data);
    let spreadsheet = getSpreadsheet("Expenses");
    if (spreadsheet == null) {
        throw "Sheet does not exist";
    }
    let activeSheet = getOrCreateSheet(spreadsheet, "Rent");
    for (let obj of data) {
        lastRow = sheetIsBlank(activeSheet) ? 1 : activeSheet.getLastRow() + 1;
        let textFinder = activeSheet.createTextFinder(obj["DUE_DATE"]);
        let found = textFinder.findNext();
        if (found) {
            // if DUE_DATE exists continue to next object
            continue
        }
        for (let key in obj) {
            // write data to sheet
            findColumnInsertValue(activeSheet, lastRow, key, obj[key]);
        }
    }
}
