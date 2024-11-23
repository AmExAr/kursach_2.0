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
            ("file,f", po::value<std::string>()->default_value("Presentation.ppt"), "Path to the presentation. Default: Presentation.ppt")
            ("text,t", po::value<std::string>(), "Path to the output file. Default: output_from_PPT.txt")
            ("pictures,p", po::value<std::string>(), "Path to the directory for output files. Default: ./pictures/");


        po::variables_map vm;
        po::store(parse_command_line(argc, argv, desc), vm);
        po::notify(vm);


        if (vm.count("help")) {
                std::cout << desc << std::endl;
                return 0;
        }


        std::cout << vm["file"].as<std::string>() << std::endl;
        std::cout << vm["text"].as<std::string>() << std::endl;
        std::cout << vm["pictures"].as<std::string>() << std::endl;


        wchar_t *presentationFile = String2WChar(vm["file"].as<std::string>());
        std::wcout << "Read presentation from: " << *presentationFile << std::endl;
        PPT presentation(presentationFile);


        if (vm.count("text")) {
            std::string outputTextFile = vm["text"].as<std::string>();

            if (outputTextFile.empty()) { outputTextFile = ".\\output_from_PPT.txt"; }

            std::wcout << "Output text file: " << *String2WChar(outputTextFile) << std::endl;

            presentation.GetText(String2WChar(outputTextFile));
        }


        if (vm.count("pictures")){
            std::wstring picturesDir = vm["pictures"].as<std::wstring>();

            if (picturesDir.empty()) { picturesDir = L".\\pictures"; }

            CheckAndCreateDir(picturesDir);
            std::wcout << "Pictures will be saved to: " << picturesDir << std::endl;

            presentation.GetPics(picturesDir);
        }

    }
    catch (const std::exception &e) {
        std::cerr << e.what() << std::endl;
    }


    /*
    PPT presentation(L"2003_bouth.ppt");
    //PPT presentation(L"2007_bouth.ppt");
    //PPT presentation(L"PICTUR.ppt");

    presentation.GetPics((std::wstring)L"C:\\Users\\McPoh\\Desktop\\kursach_2.0\\pptElemGetter\\New_folder\\");

    return 0;
    */
}
