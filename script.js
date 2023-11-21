const excel = require('excel4node');
const fs = require('fs');
const data = {
  "result": {
    "requestId": "c217de8f-3ebb-4c94-a588-427b0f0ea933",
    "gstin": "29AAWCS3552Q1Z6",
    "gstr1": {
      "102020": {
        "B2BA": {
          "_id": "64e9a33628625ae9f5f509d6",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "AT": {
          "_id": "64e9a33628625ae9f5f5099e",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "B2CLA": {
          "_id": "64e8959f262341a19e5eef2e",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "B2CL": {
          "_id": "64e9a33628625ae9f5f5097a",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "B2CS": {
          "_id": "64e99d67a41c31988e165e24",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "CDNRA": {
          "_id": "64e9a33628625ae9f5f5099a",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "CDNUR": {
          "_id": "64e894e7cff7a0af8667ca4e",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "CDNURA": {
          "_id": "64e895bd262341a19e5ef2ca",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "HSNSUM": {
          "_id": "64e9a33628625ae9f5f50972",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "EXPA": {
          "_id": "64e894e8cff7a0af8667ca5a",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number might have been amended earlier. Search with latest amended document."
          }
        },
        "TXP": {
          "_id": "64e99d6ba41c31988e165e4c",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "ATA": {
          "_id": "64e9a33628625ae9f5f5098e",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "B2CSA": {
          "_id": "64e9a33628625ae9f5f509b6",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "B2B": {
          "_id": "64e9a33628625ae9f5f5098a",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "CDNR": {
          "_id": "64e9a33628625ae9f5f509ca",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "DOCISS": {
          "_id": "64e895c2262341a19e5ef2d2",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "EXP": {
          "_id": "64e8959d262341a19e5eef2a",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number might have been amended earlier. Search with latest amended document."
          }
        },
        "RETSUM": {
          "_id": "64e9a33628625ae9f5f50a6a",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "InvoiceChecksumValue": "25ff374294ade309d1199e73254e450f033ba1cad0d85e740e3a05f9862c5f59",
            "summaryType": "",
            "secSum": [
              {
                "returnSection": "CDNUR",
                "InvoiceChecksumValue": "74313561d1897af3dc03f4fae174960d28968f92b49230523faca462b848db60",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "EXPA",
                "InvoiceChecksumValue": "3b7546ed79e3e5a7907381b093c5a182cbf364c5dd0443dfa956c8cca271cc33",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "DOC_ISSUE",
                "InvoiceChecksumValue": "",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "TXPDA",
                "InvoiceChecksumValue": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "HSN",
                "InvoiceChecksumValue": "",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "EXP",
                "InvoiceChecksumValue": "3b7546ed79e3e5a7907381b093c5a182cbf364c5dd0443dfa956c8cca271cc33",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "CDNURA",
                "InvoiceChecksumValue": "74313561d1897af3dc03f4fae174960d28968f92b49230523faca462b848db60",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "NIL",
                "InvoiceChecksumValue": "",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "CDNRA",
                "InvoiceChecksumValue": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "CDNR",
                "InvoiceChecksumValue": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "B2CL",
                "InvoiceChecksumValue": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "B2CSA",
                "InvoiceChecksumValue": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "B2CS",
                "InvoiceChecksumValue": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "AT",
                "InvoiceChecksumValue": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "ATA",
                "InvoiceChecksumValue": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "TXPD",
                "InvoiceChecksumValue": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "B2BA",
                "InvoiceChecksumValue": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "B2CLA",
                "InvoiceChecksumValue": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              },
              {
                "returnSection": "B2B",
                "InvoiceChecksumValue": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
                "totalRecordCount": "",
                "totalRecordsValue": "",
                "totalIGST": "",
                "totalCGST": "",
                "totalSGST": "",
                "totalCessValue": "",
                "totalTaxableValueOfRecords": ""
              }
            ]
          }
        },
        "TXPA": {
          "_id": "64e9a33628625ae9f5f509ee",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        },
        "NIL": {
          "_id": "64e9a33628625ae9f5f50a0e",
          "gstin": "29AAWCS3552Q1Z6",
          "month": "102020",
          "data": {
            "message": "No document found in selected financial year. Either this document exists in a different financial year or the document number is incorrect."
          }
        }
      }
    }
  }
}

function jsonToExcel(data) {
    // Create a new instance of a Workbook class
    let workbook = new excel.Workbook();

    // Add Worksheets to the workbook
    let summarySheet = workbook.addWorksheet('Summary');
    let detailsSheet = workbook.addWorksheet('Details');

    // Set headers for Summary sheet
    let summaryHeadings = Object.keys(data.Summary);
    summaryHeadings.forEach((heading, index) => {
        summarySheet.cell(1, index + 1).string(heading);
    });

    // Set values for Summary sheet
    let summaryValues = Object.values(data.Summary);
    summaryValues.forEach((value, index) => {
        summarySheet.cell(2, index + 1).string(value);
    });

    // Set headers for Details sheet
    let detailsHeadings = Object.keys(data.Details[0]);
    detailsHeadings.forEach((heading, index) => {
        detailsSheet.cell(1, index + 1).string(heading);
    });

    // Set values for Details sheet
    let detailsRows = data.Details;
    detailsRows.forEach((row, rowIndex) => {
        let detailValues = Object.values(row);
        detailValues.forEach((value, valueIndex) => {
            detailsSheet.cell(rowIndex + 2, valueIndex + 1).string(value);
        });
    });

    // Write to a file
    workbook.write('data.xlsx');
}

jsonToExcel(data);