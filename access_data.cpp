//It's just display the excel data in the output

#include <iostream>
#include <OpenXLSX.hpp> // Include the OpenXLSX library

using namespace std;
using namespace OpenXLSX; // OpenXLSX namespace

int main() {
    // Path to the Excel file (Update to your actual file path)
    string filePath = "C:\\Users\\smart\\OneDrive\\Desktop\\cpp2Excel\\new.xlsx";  

    try {
        // Open the Excel file
        XLDocument doc;
        doc.open(filePath);

        // Access the first worksheet
        auto sheet = doc.workbook().worksheet("Sheet1"); 

        // Read and print data from the first 5 rows and columns
        for (int row = 1; row <= 5; ++row) {
            for (int col = 1; col <= 5; ++col) {
                cout << sheet.cell(row, col).value().getString() << "\t"; // Print cell value
            }
            cout << endl;
        }

        doc.close(); // Close the document
    } catch (const exception& e) {
        cerr << "Error: " << e.what() << endl;
    }

    return 0;
}
