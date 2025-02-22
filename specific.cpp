//It returns the data by giving colomn and row number

#include <iostream>
#include <OpenXLSX.hpp> // Include the OpenXLSX library

using namespace std;
using namespace OpenXLSX; // OpenXLSX namespace

int main() {
    // Path to the Excel file (Update to your actual file path)
    string filePath = "C:\\Users\\smart\\OneDrive\\Desktop\\cpp2Excel\\student_marks.xlsx";  

    try {
        // Open the Excel file
        XLDocument doc;
        doc.open(filePath);

        // Access the first worksheet
        auto sheet = doc.workbook().worksheet("Sheet1");

        // Specify the row and column to fetch data (Example: Row 3, Column 2)
        int targetRow = 10;
        int targetCol = 3;

        // Retrieve and print the specific cell value
        cout << "Data at Row " << targetRow << ", Column " << targetCol << ": "
             << sheet.cell(targetRow, targetCol).value().getString() << endl;

        doc.close(); // Close the document
    } catch (const exception& e) {
        cerr << "Error: " << e.what() << endl;
    }

    return 0;
}
