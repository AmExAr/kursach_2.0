#include <iostream>
#include <fstream>
#include <iomanip>
#include <vector>
#include <map>
#include <string>
#include <wtypes.h>
#include "CompDocObj.h"

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

int main() {
    setlocale(LC_ALL, "Russian");

    ifstream inFile("12.ppt", ios::binary);
    if (!inFile.is_open()) {
        wcout << L"Ошибка открытия файла!" << endl;
        return 1;
    }

    inFile.seekg(0, ios_base::end);
    size_t fileSize = inFile.tellg();
    wcout << L"Размер файла: " << fileSize << L" байт" << endl;

    BYTE *dataBuffer = new BYTE[fileSize];
    inFile.seekg(0, ios_base::beg);
    inFile.read(reinterpret_cast<char*>(dataBuffer), fileSize);
    if (inFile.fail()) {
        wcout << L"Ошибка при чтении файла!" << endl;
        delete[] dataBuffer;
        return 1;
    }

    wstring defaultExtension;
    vector<CompoundDocumentObject::CompoundDocument_DirectoryEntryStruct> Directory;
    map<wstring, ULONGLONG> NameMap;
    vector<BinaryBlock> DataStreams;

    DWORD realDataSize = CompoundDocumentObject::GetCompoundDocumentInfo(dataBuffer, fileSize, defaultExtension, Directory, NameMap, DataStreams);
    wcout << L"Размер файла проверенный: " << realDataSize << L" байт" << endl;
    wcout << L"Расширение по умолчанию: " << defaultExtension << endl << endl;

    // Вывод оглавления и сопоставления записей
    wcout << L"Оглавление:" << endl;
    for (const auto& entry : Directory) {
        wcout << L"Имя: " << entry.EntryName << L", Размер: " << entry.StreamSize << L" байт" << endl;
    }
    wcout << endl;

    wcout << L"Сопоставление записей с индексами потоков:" << endl;
    for (const auto& pair : NameMap) {
        wcout << pair.first << L" : " << hex << pair.second << endl;
    }
    wcout << endl;

    // Находим индекс потока с именем "PowerPoint Document"
    int targetStreamIndex = findStreamIndexByName(Directory, L"PowerPoint Document");

    // Проверяем, что индекс потока найден и находится в пределах размера DataStreams
    if (targetStreamIndex >= 0 && targetStreamIndex < static_cast<int>(DataStreams.size())) {
        const vector<BYTE>& targetStream = DataStreams[targetStreamIndex];
        size_t targetSize = targetStream.size();

        // Последовательность для поиска
        const BYTE sequence1[] = {0x0F, 0x00, 0xF0, 0x0F};
        const size_t sequence1Size = sizeof(sequence1) / sizeof(sequence1[0]);

        size_t lastSequenceIndex = SIZE_MAX; // Индекс последней найденной последовательности

        // Поиск последней последовательности 0F 00 F0 0F в потоке
        for (size_t i = 0; i <= targetSize - sequence1Size; ++i) {
            if (memcmp(targetStream.data() + i, sequence1, sequence1Size) == 0) {
                lastSequenceIndex = i; // Сохраняем индекс последнего вхождения
            }
        }

        // Проверяем, было ли найдено хотя бы одно вхождение
        if (lastSequenceIndex != SIZE_MAX) {
            // Извлекаем данные после последнего вхождения
            if (lastSequenceIndex + sequence1Size < targetSize) {
                BYTE sizeByte1 = targetStream[lastSequenceIndex + sequence1Size];       // Младший байт
                BYTE sizeByte2 = targetStream[lastSequenceIndex + sequence1Size + 1];   // Старший байт

                // Объединение двух байтов в одно 16-битное целое число в правильном порядке для little-endian
                uint16_t combinedSize = (static_cast<uint16_t>(sizeByte2) << 8) | static_cast<uint16_t>(sizeByte1);

                // Преобразование в десятичное значение и вывод
                wcout << std::dec << L"Размер (после последовательности 0F 00 F0 0F): " << combinedSize << L" (десятичный формат)" << endl;

                // Извлечение байтов после combinedSize
                if (lastSequenceIndex + sequence1Size + 2 + combinedSize <= targetSize) {
                    const BYTE sequence2[] = {0xA8, 0x0F};
                    const size_t sequence2Size = sizeof(sequence2) / sizeof(sequence2[0]);
                    const BYTE sequence3[] = {0xA0, 0x0F};
                    const size_t sequence3Size = sizeof(sequence3) / sizeof(sequence3[0]);

                    // Поиск последовательности A8 0F и A0 0F в извлеченной последовательности
                    for (size_t j = 0; j <= combinedSize - sequence2Size; ++j) {
                        if (memcmp(targetStream.data() + lastSequenceIndex + sequence1Size + 2 + j, sequence2, sequence2Size) == 0) {
                            // Найдено совпадение, выводим два байта, следующие за A8 0F
                            if (j + sequence2Size + 1 < combinedSize) {
                                BYTE sizeByte3 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence2Size];       // Младший байт
                                BYTE sizeByte4 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence2Size + 1];   // Старший байт

                                // Объединение двух байтов в одно 16-битное целое число в правильном порядке для little-endian
                                uint16_t sizeText = (static_cast<uint16_t>(sizeByte4) << 8) | static_cast<uint16_t>(sizeByte3);

                                // Преобразование в десятичное значение и вывод
                                wcout << std::dec << L"Размер (после последовательности A8 0F): " << sizeText << L" (десятичный формат)" << endl;

                                // Извлечение байтов после sizeText
                                if (j + sequence2Size + 2 + sizeText <= combinedSize) {
                                    wcout << L"Последовательность байтов от A8 0F до sizeText (в UTF-8):" << endl;

                                    // Переводим байты в строку UTF-8 с заменой неотображаемых символов на перенос строки
                                    string utf8String;
                                    for (size_t k = 0; k < sizeText; ++k) {
                                        BYTE byte = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence2Size + 4 + k];

                                        // Проверяем, является ли байт отображаемым символом
                                        if (isprint(byte)) {
                                            utf8String += static_cast<char>(byte);  // Добавляем байт как символ
                                        } else {
                                            utf8String += '\n';  // Неотображаемые символы заменяем на перенос строки
                                        }
                                    }

                                    // Выводим строку UTF-8
                                    cout << utf8String << endl;
                                } else {
                                    wcout << L"Ошибка: выход за пределы потока при извлечении байтов." << endl;
                                }
                            cout << endl;
                            }
                        }

                        if (memcmp(targetStream.data() + lastSequenceIndex + sequence1Size + 2 + j, sequence3, sequence3Size) == 0) {
                            // Найдено совпадение, выводим два байта, следующие за A8 0F
                            const BYTE carriage_return[] = {0x0D}; // Для enter (carriage return)
                            if (j + sequence3Size + 1 < combinedSize) {
                                BYTE sizeByte5 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence3Size];       // Младший байт
                                BYTE sizeByte6 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence3Size + 1];   // Старший байт

                                // Объединение двух байтов в одно 16-битное целое число в правильном порядке для little-endian
                                uint16_t sizeText = (static_cast<uint16_t>(sizeByte6) << 8) | static_cast<uint16_t>(sizeByte5);

                                // Преобразование в десятичное значение и вывод
                                wcout << std::dec << L"Размер (после последовательности A0 0F): " << sizeText << L" (десятичный формат)" << endl;
                                // Извлечение байтов после sizeText
                                if (j + sequence3Size + 2 + sizeText <= combinedSize) {
                                    wcout << L"Последовательность байтов от A8 0F до sizeText (в UTF-8):" << endl;

                                    // Переводим байты в строку UTF-8 с заменой неотображаемых символов на перенос строки
                                    string utf8String;
                                    for (size_t k = 0; k < sizeText; ++k) {
                                        BYTE byte = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence3Size + 4 + k];

                                        // Проверяем, является ли байт carriage return и заменяем его на '\n'
                                        if (byte == carriage_return[0]) {
                                            utf8String += '\n';
                                        }

                                        else if (isprint(byte)) {
                                            utf8String += static_cast<char>(byte);  // Добавляем байт как символ
                                        }
                                    }

                                    // Выводим строку UTF-8
                                    cout << utf8String << endl;
                                }
                            }
                        }
                    }
                } else {
                    wcout << L"Ошибка: выход за пределы потока при извлечении байтов." << endl;
                }
            } else {
                wcout << L"Следующие байты после последней последовательности выходят за пределы потока." << endl;
            }
        } else {
            wcout << L"Ошибка: последовательность 0F 00 F0 0F не найдена." << endl;
        }
    } else {
        wcout << L"Ошибка: Поток с именем 'PowerPoint Document' не найден." << endl;
    }

    delete[] dataBuffer;

    return 0;
}

