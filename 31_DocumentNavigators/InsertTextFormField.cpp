
#include"Spire.Doc.o.h"
using namespace std;
using namespace Spire::Doc;

int main()
{
    // Create a new empty Word document.
    intrusive_ptr<Document> document = new Document();

    // Initialize a DocumentNavigator to help insert content and form fields into the document.
    intrusive_ptr<DocumentNavigator> navigator = new DocumentNavigator(document);

    // Write a label indicating the type of the following form field: Calculation.
    navigator->Write(L"TextFormFieldType.Calculation: ");

    // Insert a calculation-type text form field.
    navigator->InsertTextFormField(L"CalculationTextField", TextFormFieldType::Calculation, L"0", L"=3+1", 30);

    // Insert a line break (carriage return) to move to the next line.
    navigator->Writeln();

    // Write a label for a number-only text form field.
    navigator->Write(L"TextFormFieldType.NumberText: ");

    // Insert a number-input text form field.
    navigator->InsertTextFormField(L"NumberText", TextFormFieldType::NumberText, L"0", L"100", 30);

    // Insert a line break.
    navigator->Writeln();

    // Write a label for a date-input text form field.
    navigator->Write(L"TextFormFieldType.DateText: ");

    // Insert a date-formatted text form field.
    navigator->InsertTextFormField(L"DateText", TextFormFieldType::DateText, L"yyyy/M/d", L"2025/8/1", 30);

    // Insert a line break.
    navigator->Writeln();

    // Enable automatic field updating. 
    document->SetIsUpdateFields(true);

    // Save the resulting document to a file.
    document->SaveToFile(L"InsertTextFormField.docx", FileFormat::Docx);

    // Close the document to release internal resources.
    document->Close();

    // Explicitly dispose of the document object to free memory immediately.
    document->Dispose();
}

