
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
    doc->LoadFromFile(L"Data/MoveToHeaderAndFooter.docx");

    // Move the navigator's cursor to the first section of the document (section index 0).
    navigator->MoveToSection(0);

    // Navigate to the footer of the first page in the current section.
    navigator->MoveToHeaderFooter(HeaderFooterType::FooterFirstPage);

    // Write a new line of text into the first-page footer.
    navigator->Writeln(L"The footer on the first page.");

    // Navigate to the header of the first page in the current section.
    navigator->MoveToHeaderFooter(HeaderFooterType::HeaderFirstPage);

    // Write a new line of text into the first-page header.
    navigator->Writeln(L"The header on the first page.");

    // Save the modified document to a new file named "MoveToHeaderAndFooter.docx" in DOCX format.
    doc->SaveToFile(L"MoveToHeaderAndFooter.docx", FileFormat::Docx);

    // Close the document to release internal resources.
    doc->Close();

    // Explicitly dispose of the document object to free memory immediately.
    doc->Dispose();
}


