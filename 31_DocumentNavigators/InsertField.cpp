
#include"Spire.Doc.o.h"
using namespace std;
using namespace Spire::Doc;

int main()
{
	wstring outputFile = L"InsertField.docx";

    // Initialize a new Document object with automatic memory management via intrusive_ptr
    intrusive_ptr<Document> document = new Document();

    // Create a DocumentNavigator object to facilitate easy content insertion and formatting.
    intrusive_ptr<DocumentNavigator> navigator = new DocumentNavigator(document);

    // Get the current section of the document where content will be added.
    intrusive_ptr<Section> section = navigator->GetCurrentSection();

    // Add a new paragraph to the section.
    intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

    // Write the label text "Add page fields:" followed by a paragraph break.
    navigator->Writeln(L"Add page fields:");

    // Insert a PAGE field with numeric formatting.
    navigator->InsertField(L"PAGE \\# \"Page 0\"");

    // Insert a paragraph break (empty line).
    navigator->Writeln();

    // Insert the same PAGE field but with a custom result text to simulate a placeholder or initial value.
    navigator->InsertField(L"PAGE \\# \"Page 0\"", L"3");

    // Insert another paragraph break.
    navigator->Writeln();

    // Insert a built-in PAGE field using FieldType enumeration, with the field result displayed (true = show result).
    navigator->InsertField(FieldType::FieldPage, true);

    // Insert another paragraph break.
    navigator->Writeln();

    // Insert another PAGE field, but this time hide the field result (false = show field code instead of result).
    navigator->InsertField(FieldType::FieldPage, false);

    // Save the document to the specified file path using the DOCX 2013 format
    document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);

    // Close the document to release all associated resources and finalize the file
    document->Close();

    // Explicitly dispose of the document object to free memory immediately.
    document->Dispose();
}

