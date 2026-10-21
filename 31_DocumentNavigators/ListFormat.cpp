
#include"Spire.Doc.o.h"
using namespace std;
using namespace Spire::Doc;

int main()
{
    // Create a new document instance.
    intrusive_ptr<Document> document = new Document();

    // Add a new section to the document.
    intrusive_ptr<Section> section = document->AddSection();

    // Add the first paragraph to the section.
    intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

    // Append text to the first paragraph.
    paragraph->AppendText(L"This is the first paragraph.");

    // Add a second paragraph to the section.
    paragraph = section->AddParagraph();

    // Append text to the second paragraph.
    paragraph->AppendText(L"This is the second paragraph.");


    // Add a third paragraph to the section.
    paragraph = section->AddParagraph();

    // Append text to the third paragraph.
    paragraph->AppendText(L"This is the third paragraph.");

    // Create a document navigator to facilitate navigation and formatting.
    intrusive_ptr<DocumentNavigator> navigator = new DocumentNavigator(document);

    // Apply bullet list style to the current position (first paragraph by default).
    navigator->GetListFormat()->ApplyBulletStyle();

    // Move the navigator's cursor to the third paragraph (index 2) at character offset 0.
    navigator->MoveToParagraph(2, 0);

    // Apply bullet list style to the third paragraph.
    navigator->GetListFormat()->ApplyBulletStyle();

    // Save the document to a file named "ListFormat.docx" in DOCX format.
    document->SaveToFile(L"ListFormat.docx", FileFormat::Docx);

    // Close the document to release internal resources.
    document->Close();

    // Explicitly dispose of the document object to free memory immediately.
    document->Dispose();

}

