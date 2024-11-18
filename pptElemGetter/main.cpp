#include "pptElements.h"

int main()
{
    setlocale(LC_ALL, "Russian");

    PPT presentation(L"C:\\Users\\Artyom\\Downloads\\PPT365.ppt");

    presentation.GetText(L"test1.txt");

    return 0;
}
