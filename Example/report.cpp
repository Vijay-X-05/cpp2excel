#include <iostream>
#include <OpenXLSX.hpp>
#include <string>
#include <vector>

using namespace std;
using namespace OpenXLSX;

int main() {
    string filePath = "C:\\Users\\smart\\OneDrive\\Desktop\\cpp2Excel\\diploma_data.xlsx"; // Ensure correct path

    try {
        XLDocument doc;
        doc.open(filePath); // Open the Excel file
        auto workbook = doc.workbook();
        
        vector<string> failedStudents;

        // Iterate through all sheets
        for (const auto& sheetName : workbook.worksheetNames()) {
            XLWorksheet sheet = workbook.worksheet(sheetName);
            size_t rowCount = sheet.rowCount();
            size_t colCount = sheet.columnCount();

            // Check if there's enough data
            if (rowCount < 2 || colCount < 3) { // Assuming at least 3 columns exist (S.No, Pin Number, and Subjects)
                cout << "Not enough data in sheet: " << sheetName << endl;
                continue;
            }

            // Get subject names from the header row (assuming subjects start from column 3)
            vector<string> subjectNames;
            for (size_t col = 3; col <= colCount; ++col) {
                subjectNames.push_back(sheet.cell(1, col).value().get<string>());
            }

            // Iterate through students (rows), skipping the header (row 1)
            for (size_t row = 2; row <= rowCount; ++row) {
                string pinNumber = sheet.cell(row, 2).value().get<string>();
                string failedSubjects = "";

                // Check subject marks (starting from column 3)
                for (size_t col = 3; col <= colCount; ++col) {
                    XLCellValue cellValue = sheet.cell(row, col).value();
                    if (cellValue.type() == XLValueType::Integer) {
                        int marks = cellValue.get<int64_t>();
                        if (marks < 40) { // Assuming below 40 is fail
                            failedSubjects += subjectNames[col - 3] + " (" + to_string(marks) + "), ";
                        }
                    }
                }

                // If the student failed in any subject, add them to the report
                if (!failedSubjects.empty()) {
                    failedStudents.push_back("Pin: " + pinNumber + " | Failed: " + failedSubjects);
                }
            }
        }

        doc.close();

        // Display failed students
        if (failedStudents.empty()) {
            cout << "No students failed in any subject." << endl;
        } else {
            cout << "List of students who failed:\n";
            for (const auto& student : failedStudents) {
                cout << student << endl;
            }
        }

    } catch (const exception& e) {
        cerr << "Error: " << e.what() << endl;
    }

    return 0;
}
