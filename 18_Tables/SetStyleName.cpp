
#include"Spire.Doc.o.h"
using namespace std;
using namespace Spire::Doc;

int main()
{
	wstring outputFile = L"SetStyleName.docx";
	// Initialize a new Document object with automatic memory management via intrusive_ptr
	intrusive_ptr<Document> document = new Document();

	// Add a new section to the document
	intrusive_ptr<Section> section = document->AddSection();

	// Create a new table style named "TestTableStyle1" and cast it to TableStyle pointer
	intrusive_ptr<TableStyle> tableStyle = Object::Dynamic_cast<TableStyle>(document->GetStyles()->Add(StyleType::TableStyle, L"TestTableStyle1"));

	// Set the left indentation for the table style to 55 points
	tableStyle->SetLeftIndent(55);

	// Set the color of all borders in the table style to Green
	tableStyle->GetBorders()->SetColor(Color::GetGreen());

	// Set the border line style to Single (solid line)
	tableStyle->GetBorders()->SetBorderType(BorderStyle::Single);

	// Set the horizontal alignment of the table to Right based on the row alignment setting
	tableStyle->SetHorizontalAlignment(RowAlignment::Right);

	// Add a new table to the current section
	intrusive_ptr<Table>  table = section->AddTable();

	// Reset and initialize the table structure with 5 rows and 3 columns
	table->ResetCells(5, 3);

	// Access the first cell (Row 0, Cell 0), add a paragraph to it, and append the text
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"Aligned according to left indent");

	// Set the preferred width of the table to 300 points
	table->SetPreferredWidth(PreferredWidth::FromPoints(300));

	// Apply the previously created custom style "TestTableStyle1" to this table
	table->GetFormat()->SetStyleName(L"TestTableStyle1");

	// Save the document to the specified file path using the DOCX 2013 format
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);

	// Close the document to release all associated resources and finalize the file
	document->Close();
}