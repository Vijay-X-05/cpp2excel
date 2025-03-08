#include <iostream>
#include <OpenXLSX.hpp>
#include <podofo.h>

using namespace std;
using namespace OpenXLSX;
using namespace PoDoFo;

const int PASS_MARKS = 35;

int main() {
    string filePath = "C:\\Users\\smart\\OneDrive\\Desktop\\cpp2Excel\\diploma_data.xlsx";  
    string pdfPath = "failed_students_report.pdf";  

    try {
        XLDocument doc;
        doc.open(filePath);
        auto sheet = doc.workbook().worksheet("Sheet1");  

        size_t rowCount = sheet.rowCount();
        size_t colCount = sheet.columnCount();

        // Initialize PDF document
        PdfStreamedDocument pdfDoc(pdfPath.c_str());
        PdfFont* font = pdfDoc.CreateFont("Arial");
        font->SetFontSize(12);

        PdfPage* page = pdfDoc.CreatePage(PdfPage::CreateStandardPageSize(PdfPageSize::A4));
        PdfPainter painter;
        painter.SetPage(page);
        painter.SetFont(font);

        double yPosition = 750;  
        painter.DrawText(200, yPosition, "Failed Students Report");
        yPosition -= 30;

        for (size_t row = 2; row <= rowCount; ++row) {
            string pin = sheet.cell(row, 2).value().get<string>();  
            bool failed = false;
            string failReport = "PIN " + pin + " failed in: ";

            for (size_t col = 3; col <= colCount; ++col) {
                XLCellValue cellValue = sheet.cell(row, col).value();
                int marks = cellValue.get<int64_t>();  

                if (marks < PASS_MARKS) {  
                    string subject = sheet.cell(1, col).value().get<string>();  
                    failReport += subject + " (" + to_string(marks) + "), ";
                    failed = true;
                }
            }

            if (failed) {
                painter.DrawText(50, yPosition, failReport.c_str());
                yPosition -= 20;

                if (yPosition < 50) {  
                    painter.FinishPage();
                    page = pdfDoc.CreatePage(PdfPage::CreateStandardPageSize(PdfPageSize::A4));
                    painter.SetPage(page);
                    painter.SetFont(font);
                    yPosition = 750;
                }
            }
        }

        painter.FinishPage();
        pdfDoc.Close();
        doc.close();

        cout << "PDF report generated: " << pdfPath << endl;
    } catch (const exception& e) {
        cerr << "Error: " << e.what() << endl;
    }

    return 0;
}
