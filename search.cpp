//It prints the details and marks only

#include <iostream>
#include <OpenXLSX.hpp>
#include <string>
#include <algorithm> // For transform and remove_if

using namespace std;
using namespace OpenXLSX;

// Function to trim spaces from a string
string trim(const string& str) {
    string trimmed = str;
    trimmed.erase(trimmed.begin(), find_if(trimmed.begin(), trimmed.end(), [](unsigned char ch) { return !isspace(ch); }));
    trimmed.erase(find_if(trimmed.rbegin(), trimmed.rend(), [](unsigned char ch) { return !isspace(ch); }).base(), trimmed.end());
    return trimmed;
}

// Convert string to lowercase for case-insensitive comparison
string toLower(const string& str) {
    string lowerStr = str;
    transform(lowerStr.begin(), lowerStr.end(), lowerStr.begin(), ::tolower);
    return lowerStr;
}

int main() {
    string filePath = "C:\\Users\\smart\\OneDrive\\Desktop\\cpp2Excel\\student_marks.xlsx"; // Ensure correct path

    try {
        XLDocument doc;
        doc.open(filePath); // Open the Excel file
        auto workbook = doc.workbook();

        string searchPIN;
        cout << "Enter PIN number to search: ";
        cin >> searchPIN;

        // Trim and convert searchPIN to lowercase
        searchPIN = toLower(trim(searchPIN));

        bool found = false;

        // Iterate through all sheets
        for (const auto& sheetName : workbook.worksheetNames()) {
            XLWorksheet sheet = workbook.worksheet(sheetName);
            size_t rowCount = sheet.rowCount();
            size_t colCount = sheet.columnCount();

            // Iterate through rows to find the PIN
            for (size_t row = 2; row <= rowCount; ++row) { // Skipping header row (row 1)
                XLCellValue cellValue = sheet.cell(row, 2).value(); // Column 2 is "Pin Number"

                string pinValue;
                if (cellValue.type() == XLValueType::String) {
                    pinValue = cellValue.get<string>();
                } else if (cellValue.type() == XLValueType::Integer) {
                    pinValue = to_string(cellValue.get<int64_t>());
                } else if (cellValue.type() == XLValueType::Float) {
                    pinValue = to_string(cellValue.get<double>());
                }

                // Trim and convert to lowercase before comparison
                pinValue = toLower(trim(pinValue));

                if (pinValue == searchPIN) { // Check PIN
                    found = true;
                    cout << "Details found in sheet: " << sheetName << endl;

                    // Print the student's details
                    for (size_t col = 1; col <= colCount; ++col) {
                        XLCellValue headerValue = sheet.cell(1, col).value();
                        XLCellValue dataValue = sheet.cell(row, col).value();

                        string header = (headerValue.type() != XLValueType::Empty) ? headerValue.get<string>() : "Unknown";
                        string data;

                        if (dataValue.type() == XLValueType::String) {
                            data = dataValue.get<string>();
                        } else if (dataValue.type() == XLValueType::Integer) {
                            data = to_string(dataValue.get<int64_t>());
                        } else if (dataValue.type() == XLValueType::Float) {
                            data = to_string(dataValue.get<double>());
                        }

                        cout << header << ": " << data << endl;
                    }
                    break;
                }
            }
            if (found) break;
        }

        if (!found) {
            cout << "PIN not found in any sheet." << endl;
        }

        doc.close();
    } catch (const exception& e) {
        cerr << "Error: " << e.what() << endl;
    }

    return 0;
}
