#include "pptElements.h"

int main()
{
    setlocale(LC_ALL, "Russian");

    PPT presentation(L"Language 2007+.ppt");

    presentation.GetText(L"test1.txt");

    return 0;
}
