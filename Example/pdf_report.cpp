#include <iostream>
#include <OpenXLSX.hpp>
#include <hpdf.h>
#include <vector>
#include <string>

using namespace std;
using namespace OpenXLSX;

// Function to create a PDF file
void generatePDF(const vector<vector<string>>& failedStudents) {
    HPDF_Doc pdf = HPDF_New(NULL, NULL);
    if (!pdf) {
        cerr << "Error: Cannot create PDF object." << endl;
        return;
    }

    HPDF_Page page = HPDF_AddPage(pdf);
    HPDF_SetCompressionMode(pdf, HPDF_COMP_ALL);
    HPDF_Page_SetSize(page, HPDF_PAGE_SIZE_A4, HPDF_PAGE_PORTRAIT);

    float x = 50, y = 800;
    HPDF_Page_BeginText(page);
    HPDF_Page_SetFontAndSize(page, HPDF_GetFont(pdf, "Helvetica", NULL), 12);
    HPDF_Page_TextOut(page, x, y, "Failed Students Report");
    y -= 30;

    for (const auto& student : failedStudents) {
        string line = "PIN: " + student[0] + ", Subject: " + student[1] + ", Marks: " + student[2];
        HPDF_Page_TextOut(page, x, y, line.c_str());
        y -= 20;
        if (y < 50) {  // Add a new page if needed
            page = HPDF_AddPage(pdf);
            HPDF_Page_SetFontAndSize(page, HPDF_GetFont(pdf, "Helvetica", NULL), 12);
            y = 800;
        }
    }

    HPDF_Page_EndText(page);
    HPDF_SaveToFile(pdf, "failed_students_report.pdf");
    HPDF_Free(pdf);
    cout << "PDF report saved as 'failed_students_report.pdf'" << endl;
}

int main() {
    string filePath = "C:\\Users\\smart\\OneDrive\\Desktop\\cpp2Excel\\diploma_data.xlsx"; 

    try {
        XLDocument doc;
        doc.open(filePath);
        auto sheet = doc.workbook().worksheet("Sheet1");

        size_t rowCount = sheet.rowCount();
        size_t colCount = sheet.columnCount();

        vector<vector<string>> failedStudents;

        for (size_t row = 2; row <= rowCount; ++row) {  
            string pin = sheet.cell(row, 2).value().get<string>();  

            for (size_t col = 3; col <= colCount; ++col) {  
                string subject = sheet.cell(1, col).value().get<string>();  
                int marks = sheet.cell(row, col).value().get<int>();  

                if (marks < 35) {  
                    failedStudents.push_back({pin, subject, to_string(marks)});
                }
            }
        }

        doc.close();

        if (!failedStudents.empty()) {
            generatePDF(failedStudents);
        } else {
            cout << "No failed students found." << endl;
        }

    } catch (const exception& e) {
        cerr << "Error: " << e.what() << endl;
    }

    return 0;
}
