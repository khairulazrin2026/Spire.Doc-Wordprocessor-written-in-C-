
#include"Spire.Doc.o.h"
using namespace std;
using namespace Spire::Doc;


void addcellcontent(intrusive_ptr<TableCell> cell, wstring Content)
{
    // Add a new paragraph to the table cell.
    intrusive_ptr<Paragraph> para = cell->AddParagraph();

    // Append the specified text content to the paragraph inside the cell.
    para->AppendText(Content.c_str());

    // Center-align the text horizontally within the paragraph.
    para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

    // Set the background (shading) color of the table cell to CadetBlue.
    cell->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::GetCadetBlue());

    // Vertically center the content within the table cell.
    cell->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
}

int main()
{
    // Create a new document instance.
    intrusive_ptr<Document> doc = new Document();

    // Initialize a document navigator to navigate and modify the document content.
    intrusive_ptr<DocumentNavigator> navigator = new DocumentNavigator(doc);

    // Begin creating a new table in the document.
    navigator->StartTable();

    // Insert the first cell into the current row of the table and get its reference.
    intrusive_ptr<TableCell> cell1 = navigator->InsertCell();

    // Add content to the first cell of the first row.
    addcellcontent(cell1, L"Row 1, Cell 1");

    // Insert the second cell into the current row of the table and get its reference.
    intrusive_ptr<TableCell> cell2 = navigator->InsertCell();

    // Add content to the second cell of the first row.
    addcellcontent(cell2, L"Row 1, Cell 2");

    // Insert the third cell into the current row of the table and get its reference.
    intrusive_ptr<TableCell> cell3 = navigator->InsertCell();

    // Add content to the third cell of the first row.
    addcellcontent(cell3, L"Row 1, Cell 3");

    // End the current row and move to the next row in the table.
    navigator->EndRow();

    // Insert the first cell into the new row of the table and get its reference.
    intrusive_ptr<TableCell> cell4 = navigator->InsertCell();

    // Add content to the first cell of the second row.
    addcellcontent(cell4, L"Row 2, Cell 1");

    // End the table creation process.
    navigator->EndTable();

    // Move the navigator's cursor to a specific cell in the first table
    navigator->MoveToCell(0, 0, 1, 0);

    // Insert (overwrite) the text at the current cursor position inside the target cell.
    navigator->Write(L"new content");

    // Get the formatting object of the current cell where the navigator is positioned.
    intrusive_ptr<CellFormat> cellformat = navigator->GetCellFormat();

    // Clear all existing formatting applied to the current cell.
    cellformat->ClearFormatting();

    // Set the background (shading) color of the cell to GreenYellow.
    cellformat->GetShading()->SetBackgroundPatternColor(Color::GetCadetBlue());

    // Save the modified document to the specified output file in DOCX format.
    doc->SaveToFile(L"OperateTableCell.docx", FileFormat::Docx);

    // Close the document to release internal resources.
    doc->Close();

    // Explicitly dispose of the document object to free memory immediately.
    doc->Dispose();
}


