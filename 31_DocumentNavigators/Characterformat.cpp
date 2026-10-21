
#include"Spire.Doc.o.h"
using namespace std;
using namespace Spire::Doc;

int main()
{
	wstring outputFile = L"Characterformat.docx";

    // Initialize a new Document object with automatic memory management via intrusive_ptr
    intrusive_ptr<Document> document = new Document();

    // Create a DocumentNavigator object to facilitate easy content insertion and formatting.
    intrusive_ptr<DocumentNavigator> navigator = new DocumentNavigator(document);

    // Write plain text into the document (without special formatting yet).
    navigator->Write(L"Write plain text into the document (without special formatting yet).");

    // Insert a line break after the previous text.
    navigator->InsertBreak(BreakType::LineBreak);

    // Enable underline formatting for subsequent text.
    navigator->GetCharacterFormat()->SetUnderlineStyle(UnderlineStyle::Single);

    // Set bold formatting for subsequent text.
    navigator->GetCharacterFormat()->SetBold(TRUE);

    // Enable shadow effect for subsequent text.
    navigator->GetCharacterFormat()->SetIsShadow(TRUE);

    // Set the text color to blue for subsequent text.
    navigator->GetCharacterFormat()->SetTextColor(Color::GetBlue());

    // Write formatted text using the current character formatting settings.
    navigator->Write(L"Write formatted text using the current character formatting settings.");

    // Insert another line break.
    navigator->InsertBreak(BreakType::LineBreak);

    // Save the current character formatting settings onto an internal stack for later reuse.
    navigator->PushCharacterFormat();

    // Clear all character formatting to default (e.g., no bold, no color, etc.).
    navigator->GetCharacterFormat()->ClearFormatting();

    // Write text with cleared (default) formatting.
    navigator->Write(L"Write text with cleared (default) formatting");

    // Insert another line break.
    navigator->InsertBreak(BreakType::LineBreak);

    // Restore the previously saved character formatting from the stack.
    navigator->PopCharacterFormat();

    // Write text using the restored formatting.
    navigator->Write(L"Write text using the restored formatting.");

    // Save the document to the specified file path using the DOCX 2013 format
    document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);

    // Close the document to release all associated resources and finalize the file
    document->Close();

    // Explicitly dispose of the document object to free memory immediately.
    document->Dispose();
}