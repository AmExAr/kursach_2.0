#include "pptElements.h"

int wmain()
{
    setlocale(LC_ALL, "Russian");

    PPT presentation(L"2003_bouth.ppt");
    //PPT presentation(L"2007_bouth.ppt");
    //PPT presentation(L"PICTUR.ppt");

    presentation.GetPics((std::wstring)L"C:\\Users\\McPoh\\Desktop\\kursach_2.0\\pptElemGetter\\New_folder\\");

    return 0;
}
