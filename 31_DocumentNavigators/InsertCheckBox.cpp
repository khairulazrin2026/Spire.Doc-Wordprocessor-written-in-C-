
#include"Spire.Doc.o.h"
using namespace std;
using namespace Spire::Doc;

int main()
{
	wstring outputFile = L"InsertCheckBox.docx";

    // Initialize a new Document object with automatic memory management via intrusive_ptr
    intrusive_ptr<Document> document = new Document();

    // Create a DocumentNavigator object to facilitate easy content insertion and formatting.
    intrusive_ptr<DocumentNavigator> navigator = new DocumentNavigator(document);

    // Get the current section of the document where content will be added.
    intrusive_ptr<Section> section = navigator->GetCurrentSection();

    // Add a new paragraph to the section.
    intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

    // Append descriptive text indicating a checkbox control will follow (default unselected).
    paragraph->AppendText(L"Add CheckBox Content Control: ");

    // Append additional label text for the first checkbox example.
    paragraph->AppendText(L"Default: Unselected: ");

    // Move the cursor to the beginning of the first paragraph (index 1, offset 0) to insert content precisely.
    navigator->MoveToParagraph(1, 0);

    // Insert an unchecked checkbox content control named "Checkbox1" with a size of 20 points.
    navigator->InsertCheckBox(L"Checkbox1", false, 20);

    // Add a new paragraph for the second checkbox example.
    paragraph = section->AddParagraph();

    // Append label text indicating this checkbox is selected by default.
    paragraph->AppendText(L"Default: Selected: ");

    // Move the cursor to the beginning of the second paragraph (index 2, offset 0).
    navigator->MoveToParagraph(2, 0);

    // Insert a checked checkbox content control named "Checkbox2" with a size of 20 points.
    navigator->InsertCheckBox(L"Checkbox2", true, 20);

    // Add a new paragraph for the third checkbox example.
    paragraph = section->AddParagraph();

    // Append label text explaining that this checkbox starts unchecked but will be set to checked via API.
    paragraph->AppendText(L"Default: Uncheck. Set to check now: ");

    // Move the cursor to the beginning of the third paragraph (index 3, offset 0).
    navigator->MoveToParagraph(3, 0);

    // Insert a checkbox named "Checkbox3" that is initially unchecked (false), then programmatically marked as checked (true), with size 20.
    navigator->InsertCheckBox(L"Checkbox3", false, true, 20);

    // Add a new paragraph for the fourth checkbox example.
    paragraph = section->AddParagraph();

    // Append label text for a checkbox that appears selected by default in the UI but is actually stored as unchecked.
    paragraph->AppendText(L"Default selection, unselected by default: ");

    // Move the cursor to the beginning of the fourth paragraph (index 4, offset 0).
    navigator->MoveToParagraph(4, 0);

    // Insert a checkbox named "Checkbox4" that appears checked (true) but is stored as unchecked (false), with size 20.
    navigator->InsertCheckBox(L"Checkbox4", true, false, 20);

    // Save the document to the specified file path using the DOCX 2013 format
    document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);

    // Close the document to release all associated resources and finalize the file
    document->Close();

    // Explicitly dispose of the document object to free memory immediately.
    document->Dispose();
}

