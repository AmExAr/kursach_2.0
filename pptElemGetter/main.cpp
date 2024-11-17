#include "pptElements.h"

int main()
{
    setlocale(LC_ALL, "Russian");

    PPT presentation(L"Launguage.ppt");

    presentation.GetText(L"test1.txt");

    return 0;
}
