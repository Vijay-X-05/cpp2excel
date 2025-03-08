#include <iostream>
#include <OpenXLSX.hpp>
#include <vector>
#include <string>
#include <algorithm>

using namespace std;
using namespace OpenXLSX;

// Function to trim spaces from a string
string trim(const string& str) {
    string trimmed = str;
    trimmed.erase(trimmed.begin(), find_if(trimmed.begin(), trimmed.end(), [](unsigned char ch) { return !isspace(ch); }));
    trimmed.erase(find_if(trimmed.rbegin(), trimmed.rend(), [](unsigned char ch) { return !isspace(ch); }).base(), trimmed.end());
    return trimmed;
}

// Convert string to lowercase
string toLower(const string& str) {
    string lowerStr = str;
    transform(lowerStr.begin(), lowerStr.end(), lowerStr.begin(), ::tolower);
    return lowerStr;
}

int main() {
    string filePath = "C:\\Users\\smart\\OneDrive\\Desktop\\cpp2Excel\\marks_sheet.xlsx"; // Update the correct path
    vector<size_t> targetColumns = {5, 9, 13, 17, 21, 25, 29, 33, 37, 41, 45}; // Columns E, I, M, Q, U, Y, AC, AG, AK, AO, AS

    try {
        XLDocument doc;
        doc.open(filePath);
        auto workbook = doc.workbook();

        cout << "Remedial Class Eligible Students (Failed Subjects):\n";
        cout << "---------------------------------------------------\n";
        
        // Iterate through all sheets
        for (const auto& sheetName : workbook.worksheetNames()) {
            XLWorksheet sheet = workbook.worksheet(sheetName);
            size_t rowCount = sheet.rowCount();
            cout << "************************************************************\n";
            cout << "**Semister: " << sheet.name() << "**\n";
            cout << "************************************************************\n\n";

            // Iterate through rows (starting from 2 to skip the header row)
            for (size_t row = 2; row <= rowCount; ++row) {
                XLCellValue pinCellValue = sheet.cell(row, 2).value(); // Column B is "Pin Number"
                string pinNumber;

                if (pinCellValue.type() != XLValueType::Empty) {
                    if (pinCellValue.type() == XLValueType::String) {
                        pinNumber = pinCellValue.get<string>();
                    } else if (pinCellValue.type() == XLValueType::Integer) {
                        pinNumber = to_string(pinCellValue.get<int64_t>());
                    }
                }

                bool hasFailedSubjects = false;
                string failedSubjectsInfo;

                // Check specified columns for marks below 35
                for (size_t col : targetColumns) {
                    XLCellValue subjectHeader = sheet.cell(1, col).value();
                    XLCellValue markValue = sheet.cell(row, col).value();

                    string subjectName = (subjectHeader.type() != XLValueType::Empty) ? subjectHeader.get<string>() : "Unknown";
                    int marks = 0;

                    if (markValue.type() == XLValueType::Integer) {
                        marks = markValue.get<int64_t>();
                    } else if (markValue.type() == XLValueType::Float) {
                        marks = static_cast<int>(markValue.get<double>());
                    }

                    if (marks > 0 && marks < 35) { // If marks are below 35
                        hasFailedSubjects = true;
                        failedSubjectsInfo += "    " + subjectName + " - " + to_string(marks) + "\n";
                    }
                }

                if (hasFailedSubjects) {
                    cout << "PIN: " << pinNumber << "\n";
                    cout << failedSubjectsInfo;
                    cout << "------------------------------------\n";
                }
            }
        }

        doc.close();
    } catch (const exception& e) {
        cerr << "Error: " << e.what() << endl;
    }

    return 0;
}
