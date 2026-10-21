
#include"Spire.Doc.o.h"
using namespace std;
using namespace Spire::Doc;

int main()
{
	wstring outputFile = L"SetTableStyleOptions.docx";

	// Initialize a new Document object with automatic memory management via intrusive_ptr
	intrusive_ptr<Document> document = new Document();

	// Add a new section to the document
	intrusive_ptr<Section> section = document->AddSection();

	// Create a new table style named "TestTableStyle1" and cast it to TableStyle pointer
	intrusive_ptr<TableStyle> tableStyle = Object::Dynamic_cast<TableStyle>(document->GetStyles()->Add(StyleType::TableStyle, L"TestTableStyle1"));

	// Configure the default borders for the table style: set color to black and style to double line
	tableStyle->GetBorders()->SetColor(Color::GetBlack());
	tableStyle->GetBorders()->SetBorderType(BorderStyle::Double);

	// Enable row striping with a stripe size of 3 rows
	tableStyle->SetRowStripe(3);

	// Set background color for odd row stripes to light blue
	tableStyle->GetConditionalStyles()->GetItem(TableConditionalStyleType::OddRowStripe)->GetShading()->SetBackgroundPatternColor(Color::GetLightBlue());

	// Set background color for even row stripes to light cyan
	tableStyle->GetConditionalStyles()->GetItem(TableConditionalStyleType::EvenRowStripe)->GetShading()->SetBackgroundPatternColor(Color::GetLightCyan());

	// Enable column striping with a stripe size of 1 column
	tableStyle->SetColumnStripe(1);

	// Set background color for even column stripes to light pink
	tableStyle->GetConditionalStyles()->GetItem(TableConditionalStyleType::EvenColumnStripe)->GetShading()->SetBackgroundPatternColor(Color::GetLightPink());

	// Add a new table to the section
	intrusive_ptr<Table> table = section->AddTable();

	// Initialize the table structure with 5 rows and 3 columns
	table->ResetCells(5, 3);

	// Access the first cell (Row 0, Column 0), add a paragraph, and insert centered text
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"Aligned to the center of the page");

	// Set the preferred width of the table to 500 points
	table->SetPreferredWidth(PreferredWidth::FromPoints(500));

	// Apply the custom table style to the table
	table->ApplyStyle(tableStyle);

	// Explicitly enable the column stripe option in the table style options
	table->GetFormat()->SetStyleOptions(TableStyleOptions::ColumnStripe);

	// Save the document to the specified file path in DOCX 2013 format
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);

	// Close the document to release all associated resources
	document->Close();
}