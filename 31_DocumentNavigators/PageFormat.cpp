
#include"Spire.Doc.o.h"
using namespace std;
using namespace Spire::Doc;

int main()
{
    // Create a new empty document instance.
    intrusive_ptr<Document> doc = new Document();

    // Create a document navigator to help navigate and modify the document content.
    intrusive_ptr<DocumentNavigator> navigator = new DocumentNavigator(doc);

    // Load an existing Word document from the specified relative file path.
    doc->LoadFromFile(L"Data/Sample.docx");

    // Move the navigator's cursor to the first section (section index 0) of the document.
    navigator->MoveToSection(0);

    // Set the page margins for the current section.
    navigator->GetPageSetup()->SetMargins(new MarginsF(100, 80, 100, 80));

    // Set the page size of the current section to Letter.
    navigator->GetPageSetup()->SetPageSize(PageSize::Letter());

    // Save the modified document to a new file.
    doc->SaveToFile(L"PageFormat.docx", FileFormat::Docx);

    // Close the document to release internal resources.
    doc->Close();

    // Explicitly dispose of the document object to free memory immediately.
    doc->Dispose();

}


