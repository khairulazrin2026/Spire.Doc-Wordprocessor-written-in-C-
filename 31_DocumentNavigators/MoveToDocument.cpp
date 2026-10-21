
#include"Spire.Doc.o.h"
using namespace std;
using namespace Spire::Doc;

int main()
{
    wstring inputFile = L"Data/Sample.docx";

    // Create a new empty document instance.
    intrusive_ptr<Document> document = new Document();

    // Create a document navigator to help navigate and modify the document content.
    intrusive_ptr<DocumentNavigator> navigator = new DocumentNavigator(document);

    // Load an existing Word document from the specified relative file path.
    document->LoadFromFile(inputFile.c_str());

    // Move the cursor to the very beginning of the document.
    navigator->MoveToDocumentStart();

    // Write a new line of text at the start of the document.
    navigator->Writeln(L"Insert the content at the beginning of the document.");

    // Write another line of text immediately after the previous one at the start.
    navigator->Writeln(L"This is new content.");

    // Move the cursor to the very end of the document.
    navigator->MoveToDocumentEnd();

    // Insert a blank line at the end of the document.
    navigator->Writeln();

    // Insert a new line of text at the end of the document.
    navigator->Writeln(L"Insert the content at the end of the document.");

    // Save the modified document to a new file.
    document->SaveToFile(L"MoveToDocument.docx", FileFormat::Docx);

    // Close the document to release internal resources.
    document->Close();

    // Explicitly dispose of the document object to free memory immediately.
    document->Dispose();

}

