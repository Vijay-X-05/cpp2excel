#include <iostream>
#include <fstream>
#include <OpenXLSX/OpenXLSX.hpp>

using namespace OpenXLSX;
using namespace std;

int main() {
    // Open the Excel file
    XLDocument doc;
    doc.open("C:\\Users\\smart\\OneDrive\\Desktop\\cpp2Excel\\diploma_data.xlsx");
    auto sheet = doc.workbook().worksheet("Sheet1"); // Change if the sheet name is different

    // Open text file to save the output
    ofstream outFile("C:\\Users\\smart\\OneDrive\\Desktop\\cpp2Excel\\Example\\failed_students.txt");

    if (!outFile) {
        cerr << "Error: Could not create output file." << endl;
        return 1;
    }

    // Read headers
    vector<string> headers;
    for (auto col = 1; col <= sheet.row(1).cellCount(); ++col) {
        headers.push_back(sheet.cell(XLCellReference(1, col)).value().get<string>());
    }

    // Process student records
    outFile << "Failed Students Report\n\n";

    for (auto row = 2; row <= sheet.rowCount(); ++row) { // Start from row 2 (skip headers)
        string studentDetails;
        vector<string> failedSubjects;

        for (auto col = 1; col <= sheet.row(row).cellCount(); ++col) {
            auto cellValue = sheet.cell(XLCellReference(row, col)).value();

            if (col == 1) { // Assuming first column has student ID or details
                studentDetails = cellValue.get<string>();
            } else {
                if (cellValue.type() == XLValueType::Integer || cellValue.type() == XLValueType::Float) {
                    int marks = cellValue.get<int>();
                    if (marks < 35) { // If marks are below passing criteria
                        failedSubjects.push_back(headers[col - 1] + ": " + to_string(marks));
                    }
                }
            }
        }

        // If student failed in any subject, print & save their details
        if (!failedSubjects.empty()) {
            outFile << "Student: " << studentDetails << "\n";
            for (const auto& subject : failedSubjects) {
                outFile << "  - " << subject << "\n";
            }
            outFile << "---------------------------------\n";
        }
    }

    // Close files
    outFile.close();
    doc.close();

    cout << "Report saved to failed_students.txt\n";
    return 0;
}
