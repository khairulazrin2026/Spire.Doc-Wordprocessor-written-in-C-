
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

    // Move the navigator's cursor to the first section of the document (section index 0).
    navigator->MoveToSection(0);

    // Move the cursor to the first paragraph (index 0) at character position 0 within that paragraph.
    navigator->MoveToParagraph(0, 0);

    // Set the line spacing rule for the current paragraph to "Multiple" (enables custom line spacing multiplier).
    navigator->GetParagraphFormat()->SetLineSpacingRule(LineSpacingRule::Multiple);

    // Set the line spacing to 1.5 times the default font size (assuming 12-point font: 1.5 * 12 = 18 points).
    navigator->GetParagraphFormat()->SetLineSpacing(1.5 * 12);

    // Set the left indent of the current paragraph to 5 points.
    navigator->GetParagraphFormat()->SetLeftIndent(5);

    // Move the cursor to the third paragraph (index 2) at character position 0.
    navigator->MoveToParagraph(2, 0);

    // Set the background color of the current paragraph to blue.
    navigator->GetParagraphFormat()->SetBackColor(Color::GetBlue());

    // Save the modified document to a new file.
    doc->SaveToFile(L"ParagraphFormat.docx", FileFormat::Docx);

    // Close the document to release internal resources.
    doc->Close();

    // Explicitly dispose of the document object to free memory immediately.
    doc->Dispose();
}


