#include "pptElements.h"

int main()
{
    setlocale(LC_ALL, "Russian");

    PPT presentation(L"C:\\Users\\Artyom\\Downloads\\FAT.ppt");

    //presentation.GetPicture(L"ПАПКА для картинок1");

    presentation.GetText(L"test2.txt");

    return 0;
}
