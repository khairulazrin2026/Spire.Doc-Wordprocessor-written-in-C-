
#include"Spire.Doc.o.h"
using namespace std;
using namespace Spire::Doc;

int main()
{
	wstring outputFile = L"SetTableStyle.docx";

    // Create a new Document instance using intrusive_ptr for automatic memory management
    intrusive_ptr<Document> document = new Document();

    // Add a new table style named "TestTableStyle1" to the document's style collection
    // Dynamic_cast is used to convert the generic Style object to a TableStyle object
    intrusive_ptr<TableStyle> tableStyle = Object::Dynamic_cast<TableStyle>(
        document->GetStyles()->Add(StyleType::TableStyle, L"TestTableStyle1")
    );

    // Set the border color of the table style to Blue
    tableStyle->GetBorders()->SetColor(Color::GetBlue());

    // Set the border type to Single line
    tableStyle->GetBorders()->SetBorderType(BorderStyle::Single);

    // Set the horizontal alignment of the table to Center (relative to the page)
    tableStyle->SetHorizontalAlignment(RowAlignment::Center);

    // --- Create the first section and table ---
    // Add a new section to the document
    intrusive_ptr<Section> section = document->AddSection();

    // Add a new table to the section
    intrusive_ptr<Table> table = section->AddTable();

    // Initialize the table with 1 row and 1 column
    table->ResetCells(1, 1);

    // Access the first cell (Row 0, Cell 0), add a paragraph, and append text
    table->GetRows()->GetItemInRowCollection(0)
        ->GetCells()->GetItemInCellCollection(0)
        ->AddParagraph()->AppendText(L"Aligned to the center of the page");

    // Set the preferred width of the table to 300 points
    table->SetPreferredWidth(PreferredWidth::FromPoints(300));

    // Apply the previously defined "TestTableStyle1" (Centered, Blue borders) to this table
    table->ApplyStyle(tableStyle);

    // --- Prepare for the second table ---
    // Add an empty paragraph to create some vertical space between tables
    section->AddParagraph()->AppendText(L"");

    // --- Create the second table style (Left Indented) ---
    // Add a new table style named "TestTableStyle2"
    tableStyle = Object::Dynamic_cast<TableStyle>(
        document->GetStyles()->Add(StyleType::TableStyle, L"TestTableStyle2")
    );

    // Set a left indent of 55 points for the table style
    tableStyle->SetLeftIndent(55);

    // Set the border color to Green
    tableStyle->GetBorders()->SetColor(Color::GetGreen());

    // Set the border type to Single line
    tableStyle->GetBorders()->SetBorderType(BorderStyle::Single);

    // --- Create the second table ---
    // Add another table to the same section
    table = section->AddTable();

    // Initialize with 1 row and 1 column
    table->ResetCells(1, 1);

    // Add text to the first cell indicating the alignment method
    table->GetRows()->GetItemInRowCollection(0)
        ->GetCells()->GetItemInCellCollection(0)
        ->AddParagraph()->AppendText(L"Aligned according to left indent");

    // Set the preferred width to 300 points (same as the first table)
    table->SetPreferredWidth(PreferredWidth::FromPoints(300));

    // Apply the second style (Left Indented, Green borders) to this table
    // Note: Some versions use GetFormat()->SetStyle, others use ApplyStyle directly on the table
    table->GetFormat()->SetStyle(tableStyle);

    // --- Save and Close ---
    // Save the document to the specified output file in DOCX 2013 format
    document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);

    // Close the document to release resources
    document->Close();
}