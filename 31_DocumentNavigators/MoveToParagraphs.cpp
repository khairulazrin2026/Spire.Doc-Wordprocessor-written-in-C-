
#include"Spire.Doc.o.h"
using namespace std;
using namespace Spire::Doc;

int main()
{
    wstring inputFile = L"Data/Sample.docx";

    // Create a new empty document instance.
    intrusive_ptr<Document> doc = new Document();

    // Create a document navigator to help navigate and modify the document content.
    intrusive_ptr<DocumentNavigator> navigator = new DocumentNavigator(doc);

    // Load an existing Word document from the specified relative file path.
    doc->LoadFromFile(inputFile.c_str());

    // Move the navigator's cursor to the first section of the document (section index 0).
    navigator->MoveToSection(0);

    // Move the cursor to the third paragraph (index 2) within the current section, at character offset 0.
    navigator->MoveToParagraph(2, 0);

    // Insert new text at the current cursor position, overwriting any existing content from that point onward.
    navigator->Write(L"This is new content......");

    // Save the modified document to a new file.
    doc->SaveToFile(L"MoveToParagraphs.docx", FileFormat::Docx);

    // Close the document to release internal resources.
    doc->Close();

    // Explicitly dispose of the document object to free memory immediately.
    doc->Dispose();
}

