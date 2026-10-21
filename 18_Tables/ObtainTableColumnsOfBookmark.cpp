
#include"Spire.Doc.o.h"
using namespace std;
using namespace Spire::Doc;

int main()
{
	wstring inputFile = L"Data/BookmarksInTable.docx";
	wstring outputFile = L"GetBookmarkColumns.txt";

	// Create a new Document object managed by an intrusive_ptr.
	intrusive_ptr<Document> document = new Document();

	// Load the PDF document from the file path specified.
	document->LoadFromFile(inputFile.c_str());

	// Retrieve the bookmarks collection from the document, find the bookmark named L"Test1",
	int firstColumn = document->GetBookmarks()->FindByName(L"Test1")->GetFirstColumn();

	// Similarly, get the index of the last column associated with the "Test1" bookmark.
	int lastColumn = document->GetBookmarks()->FindByName(L"Test1")->GetLastColumn();

	// Output the retrieved first column index to the standard console.
	std::cout << firstColumn << std::endl;

	// Output the retrieved last column index to the standard console.
	std::cout << lastColumn << std::endl;
}