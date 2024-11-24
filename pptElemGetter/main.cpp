#include "pptElements.h"
#include <boost/program_options.hpp>

namespace po = boost::program_options;

int main(int argc, const char *argv[])
{
    setlocale(LC_ALL, "Russian");
    try {
        po::options_description desc{"General options"};
        desc.add_options()
            ("help,h", "Show help")
            ("file,f", po::value<std::string>(), "Path to the presentation.")
            ("text,t", po::value<std::string>(), "Path to the output file.")
            ("pictures,p", po::value<std::string>(), "Path to the directory for output files.");


        po::variables_map vm;
        po::store(parse_command_line(argc, argv, desc), vm);
        po::notify(vm);


        if (vm.count("help")) {
                std::cout << desc << std::endl;
                return 0;
        }

        wchar_t *presentationFile = _wcsdup((std::wstring((vm["file"].as<std::string>()).begin(), (vm["file"].as<std::string>()).end())).c_str());
        std::wcout << "Read presentation from: " << *presentationFile << std::endl;
        PPT presentation(presentationFile);


        if (vm.count("text")) {
            std::wstring outputTextFile = std::wstring(vm["text"].as<std::string>().begin(), vm["text"].as<std::string>().end());

            std::wcout << "Output text file: " << outputTextFile << std::endl;

            presentation.GetText(_wcsdup(outputTextFile.c_str()));
        }


        if (vm.count("pictures")){
            std::wstring picturesDir = std::wstring((vm["pictures"].as<std::string>()).begin(), (vm["pictures"].as<std::string>()).end());

            CheckAndCreateDir(picturesDir);
            std::wcout << "Pictures will be saved to: " << picturesDir << std::endl;

            presentation.GetPics(picturesDir);
        }

    }
    catch (const std::exception &e) {
        std::cerr << e.what() << std::endl;
    }


    /*
    //PPT presentation(L"2003_bouth.ppt");
    //PPT presentation(L"2007_bouth.ppt");
    //PPT presentation(L"PICTUR.ppt");

    presentation.GetPics((std::wstring)L"C:\\Users\\McPoh\\Desktop\\kursach_2.0\\pptElemGetter\\New_folder\\");

    return 0;
    */
}
