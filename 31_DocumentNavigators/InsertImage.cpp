
#include"Spire.Doc.o.h"
using namespace std;
using namespace Spire::Doc;

int main()
{
	wstring outputFile = L"InsertImage.docx";
    wstring imgpath = L"Data/E-iceblue.png";
    // Initialize a new Document object with automatic memory management via intrusive_ptr
    intrusive_ptr<Document> document = new Document();

    // Create a DocumentNavigator object to facilitate easy content insertion and formatting.
    intrusive_ptr<DocumentNavigator> navigator = new DocumentNavigator(document);

    // Write a line of text into the document indicating that an image will be inserted directly.
    navigator->Writeln(L"Insert the picture directly:");

    // Insert an image at the current cursor position using the specified image file path.
    navigator->InsertImage(imgpath.c_str());

    // Add a new section to the document.
    intrusive_ptr<Section> section = document->AddSection();

    // Move the cursor to the second section (index 1, since sections are zero-based).
    navigator->MoveToSection(1);

    // Write a line of text explaining that the next image will have its dimensions set.
    navigator->Writeln(L"Set the width and height of the image:");

    // Insert an image with specified width and height (both set to 100 pixels).
    navigator->InsertImage(imgpath.c_str(), 100, 100);

    // Insert a page break to start content on a new page.
    navigator->InsertBreak(BreakType::PageBreak);

    // Write a line of text describing more advanced image positioning and formatting.
    navigator->Writeln(L"Set the width, height, offset, and wrapping style of the image:");

    // Insert an image with detailed positioning and formatting
    navigator->InsertImage(imgpath.c_str(), HorizontalOrigin::LeftMarginArea, 100, VerticalOrigin::Paragraph, 50, 100, 100, TextWrappingStyle::Through);

    // Save the document to the specified file path using the DOCX 2013 format
    document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);

    // Close the document to release all associated resources and finalize the file
    document->Close();

    // Explicitly dispose of the document object to free memory immediately.
    document->Dispose();
}

