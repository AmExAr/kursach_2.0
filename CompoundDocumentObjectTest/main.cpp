#include <iostream>
#include <fstream>
#include <iomanip>
#include <vector>
#include <map>
#include <string>
#include <wtypes.h>
#include "CompDocObj.h"

#include <locale>
#include <codecvt>

using namespace std;

// Функция для поиска индекса потока по имени EntryName
int findStreamIndexByName(const vector<CompoundDocumentObject::CompoundDocument_DirectoryEntryStruct>& Directory, const wstring& targetName) {
    for (size_t i = 0; i < Directory.size(); ++i) {
        if (Directory[i].EntryName == targetName) {
            return static_cast<int>(i);
        }
    }
    return -1;
}

// Функция для чтения файла и заполнения данных
bool readFile(const string& fileName, BYTE*& dataBuffer, size_t& fileSize) {
    ifstream inFile(fileName, ios::binary);
    if (!inFile.is_open()) {
        wcout << L"Ошибка открытия файла!" << endl;
        return false;
    }

    inFile.seekg(0, ios_base::end);
    fileSize = inFile.tellg();
    wcout << L"Размер файла: " << fileSize << L" байт" << endl;

    dataBuffer = new BYTE[fileSize];
    inFile.seekg(0, ios_base::beg);
    inFile.read(reinterpret_cast<char*>(dataBuffer), fileSize);
    if (inFile.fail()) {
        wcout << L"Ошибка при чтении файла!" << endl;
        delete[] dataBuffer;
        return false;
    }

    return true;
}

// Функция для обработки потока
void processStream(const vector<BYTE>& targetStream) {
    size_t targetSize = targetStream.size();

    // Последовательность для поиска
    const BYTE sequence1[] = {0x0F, 0x00, 0xF0, 0x0F};
    const size_t sequence1Size = sizeof(sequence1) / sizeof(sequence1[0]);
    size_t lastSequenceIndex = SIZE_MAX;

    // Поиск последней последовательности 0F 00 F0 0F
    for (size_t i = 0; i <= targetSize - sequence1Size; ++i) {
        if (memcmp(targetStream.data() + i, sequence1, sequence1Size) == 0) {
            lastSequenceIndex = i;
        }
    }

    // Если последовательность найдена
    if (lastSequenceIndex != SIZE_MAX && lastSequenceIndex + sequence1Size < targetSize) {
        BYTE sizeByte1 = targetStream[lastSequenceIndex + sequence1Size];
        BYTE sizeByte2 = targetStream[lastSequenceIndex + sequence1Size + 1];
        uint16_t combinedSize = (static_cast<uint16_t>(sizeByte2) << 8) | static_cast<uint16_t>(sizeByte1);

        if (lastSequenceIndex + sequence1Size + 2 + combinedSize <= targetSize) {
            const BYTE sequence2[] = {0xA8, 0x0F};
            const size_t sequence2Size = sizeof(sequence2) / sizeof(sequence2[0]);
            const BYTE sequence3[] = {0xA0, 0x0F};
            const size_t sequence3Size = sizeof(sequence3) / sizeof(sequence3[0]);

            for (size_t j = 0; j <= combinedSize - sequence2Size; ++j) {
                if (memcmp(targetStream.data() + lastSequenceIndex + sequence1Size + 2 + j, sequence2, sequence2Size) == 0) {
                    if (j + sequence2Size + 1 < combinedSize) {
                        BYTE sizeByte3 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence2Size];
                        BYTE sizeByte4 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence2Size + 1];
                        uint16_t sizeText = (static_cast<uint16_t>(sizeByte4) << 8) | static_cast<uint16_t>(sizeByte3);

                        if (j + sequence2Size + 2 + sizeText <= combinedSize) {
                            string utf8String;
                            wstring utf16String;

                            for (size_t k = 0; k < sizeText; ++k) {
                                BYTE byte = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence2Size + 4 + k];
                                //wcout << hex << setw(2) << setfill(L'0') << byte;

                                if (isprint(byte)) {
                                    utf8String += static_cast<char>(byte);
                                } else {
                                    utf8String += '\n';
                                }

                                // Обработка символов UTF-16LE
                                if (k + 1 < sizeText) {
                                    BYTE byte1 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence2Size + 4 + k];
                                    BYTE byte2 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence2Size + 4 + k + 1];
                                    wchar_t wchar = static_cast<wchar_t>(byte2 << 8 | byte1);

                                    if (iswprint(wchar)) {
                                        utf16String += wchar;
                                    } else {
                                        utf16String += L'\n';
                                    }

                                    ++k;
                                }
                            }

                            // Вывод UTF-8 и UTF-16LE строк
                            cout << utf8String << endl;
                            wcout << utf16String << endl;
                        } else {
                            wcout << L"Ошибка: выход за пределы потока при извлечении байтов." << endl;
                        }
                        cout << endl;
                    }
                }

                // Поиск и обработка последовательности A0 0F аналогично
                if (memcmp(targetStream.data() + lastSequenceIndex + sequence1Size + 2 + j, sequence3, sequence3Size) == 0) {
                    const BYTE carriage_return[] = {0x0D};
                    if (j + sequence3Size + 1 < combinedSize) {
                        BYTE sizeByte5 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence3Size];
                        BYTE sizeByte6 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence3Size + 1];
                        uint16_t sizeText = (static_cast<uint16_t>(sizeByte6) << 8) | static_cast<uint16_t>(sizeByte5);

                        if (j + sequence3Size + 2 + sizeText <= combinedSize) {
                            string utf8String;
                            wstring utf16String;

                            for (size_t k = 0; k < sizeText; ++k) {
                                BYTE byte = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence3Size + 4 + k];
                                //wcout << hex << setw(2) << setfill(L'0') << byte;

                                if (byte == carriage_return[0]) {
                                    utf8String += '\n';
                                } else if (isprint(byte)) {
                                    utf8String += static_cast<char>(byte);
                                }

                                if (k + 1 < sizeText) {
                                    BYTE byte1 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence3Size + 4 + k];
                                    BYTE byte2 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence3Size + 4 + k + 1];
                                    wchar_t wchar = static_cast<wchar_t>(byte2 << 8 | byte1);

                                    if (iswprint(wchar)) {
                                        utf16String += wchar;
                                    } else {
                                        utf16String += L'\n';
                                    }

                                    ++k;
                                }
                            }

                            //cout << utf8String << endl;
                            wcout << utf16String << endl;
                        }
                    }
                }
            }
        } else {
            wcout << L"Ошибка: выход за пределы потока при извлечении байтов." << endl;
        }
    }
}

int main() {
    setlocale(LC_ALL, "Russian");

    BYTE* dataBuffer = nullptr;
    size_t fileSize = 0;

    // Чтение файла
    if (!readFile("Second.ppt", dataBuffer, fileSize)) {
        return 1; // Ошибка при чтении файла
    }

    wstring defaultExtension;
    vector<CompoundDocumentObject::CompoundDocument_DirectoryEntryStruct> Directory;
    map<wstring, ULONGLONG> NameMap;
    vector<BinaryBlock> DataStreams;

    DWORD realDataSize = CompoundDocumentObject::GetCompoundDocumentInfo(dataBuffer, fileSize, defaultExtension, Directory, NameMap, DataStreams);
    wcout << L"Размер файла проверенный: " << realDataSize << L" байт" << endl;
    wcout << L"Расширение по умолчанию: " << defaultExtension << endl << endl;

    // Вывод оглавления
    wcout << L"Оглавление:" << endl;
    for (const auto& entry : Directory) {
        wcout << L"Имя: " << entry.EntryName << L", Размер: " << entry.StreamSize << L" байт" << endl;
    }
    wcout << endl;

    // Сопоставление записей
    wcout << L"Сопоставление записей с индексами потоков:" << endl;
    for (const auto& pair : NameMap) {
        wcout << pair.first << L" : " << hex << pair.second << endl;
    }
    wcout << endl;

    // Поиск и обработка потока
    int targetStreamIndex = findStreamIndexByName(Directory, L"PowerPoint Document");
    if (targetStreamIndex >= 0 && targetStreamIndex < static_cast<int>(DataStreams.size())) {
        const vector<BYTE>& targetStream = DataStreams[targetStreamIndex];
        processStream(targetStream);
    } else {
        wcout << L"Ошибка: Поток с именем 'PowerPoint Document' не найден." << endl;
    }

    // Освобождение ресурсов
    delete[] dataBuffer;
    return 0;
}
