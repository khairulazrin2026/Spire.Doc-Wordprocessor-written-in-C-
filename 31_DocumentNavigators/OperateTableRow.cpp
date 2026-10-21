
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
    // Create a new empty document instance.
    intrusive_ptr<Document> doc = new Document();

    // Create a document navigator to help navigate and modify the document content.
    intrusive_ptr<DocumentNavigator> navigator = new DocumentNavigator(doc);

    // Start creating a new table and get its reference.
    intrusive_ptr<Table> table = navigator->StartTable();

    // Initialize the table with 2 rows and 2 columns.
    table->ResetCells(2, 2);

    // Set the table width to 100% of the available page width.
    table->SetPreferredWidth(new PreferredWidth(WidthType::Percentage, (short)100));

    // Apply a single black border with line width 1 to all sides of the table.
    table->GetFormat()->GetBorders()->SetBordersAttribute(BorderStyle::Single, 1, Color::GetBlack());

    // Set the height of the first row to 30 points.
    table->GetFirstRow()->SetHeight(30);

    // Get a reference to the first cell in the first row (row 0, column 0).
    intrusive_ptr<TableCell> cell1 = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0);

    // Add content to the first cell of the first row.
    addcellcontent(cell1, L"Row 1, Cell 1");

    // Get a reference to the second cell in the first row (row 0, column 1).
    intrusive_ptr<TableCell> cell2 = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(1);

    // Add content to the second cell of the first row.
    addcellcontent(cell2, L"Row 1, Cell 2");

    // Get a reference to the first cell in the second row (row 1, column 0).
    intrusive_ptr<TableCell> cell3 = table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0);

    // Add content to the first cell of the second row.
    addcellcontent(cell3, L"Row 2, Cell 1");

    // Get a reference to the second cell in the second row (row 1, column 1).
    intrusive_ptr<TableCell> cell4 = table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(1);

    // Add content to the second cell of the second row.
    addcellcontent(cell4, L"Row 2, Cell 2");

    // Finalize the table creation.
    navigator->EndTable();

    // Delete the first row (row index 0) of the first table (table index 0) in the document.
    navigator->DeleteRow(0, 0);

    // Save the modified document to a new file named "OperateTableRow.docx" in DOCX format.
    doc->SaveToFile(L"OperateTableRow.docx", FileFormat::Docx);

    // Close the document to release internal resources.
    doc->Close();

    // Explicitly dispose of the document object to free memory immediately.
    doc->Dispose();
}


