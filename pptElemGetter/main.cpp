#include "pptElements.h"
#include <boost/program_options.hpp>

namespace po = boost::program_options;

int wmain(int argc, const wchar_t *argv[])
{
    setlocale(LC_ALL, "Russian");
    try {
        po::options_description desc{"General options"};
        desc.add_options()
            ("help,h", "Show help")
            ("file,f", po::wvalue<std::wstring>(), "Path to the presentation.")
            ("text,t", po::wvalue<std::wstring>(), "Path to the output file.")
            ("pictures,p", po::wvalue<std::wstring>(), "Path to the directory for output files.");


        po::variables_map vm;
        po::store(parse_command_line(argc, argv, desc), vm);
        po::notify(vm);


        if (vm.count("help") || argc == 1) {
                std::cout << desc << std::endl;
                return 0;
        }

        const wchar_t *presentationFile = (vm["file"].as<std::wstring>()).c_str();
        std::cout << "Pe" << std::endl;
        std::wcout << L"Read presentation from: " << *presentationFile << std::endl;
        PPT presentation(presentationFile);


        if (vm.count("text")) {
            const wchar_t *outputTextFile = (vm["text"].as<std::wstring>()).c_str();

            std::wcout << "Output text file: " << outputTextFile << std::endl;

            presentation.GetText(outputTextFile);
        }


        if (vm.count("pictures")){
            std::wstring picturesDir = vm["pictures"].as<std::wstring>();

            CheckAndCreateDir(picturesDir);
            std::wcout << "Pictures will be saved to: " << picturesDir << std::endl;

            presentation.GetPics(picturesDir);
        }

    }
    catch (const std::exception &e) {
        std::cerr << e.what() << std::endl;
    }

    return 0;

    /*
    //PPT presentation(L"2003_bouth.ppt");
    //PPT presentation(L"2007_bouth.ppt");
    //PPT presentation(L"PICTUR.ppt");

    presentation.GetPics((std::wstring)L"C:\\Users\\McPoh\\Desktop\\kursach_2.0\\pptElemGetter\\New_folder\\");

    return 0;
    */
}
