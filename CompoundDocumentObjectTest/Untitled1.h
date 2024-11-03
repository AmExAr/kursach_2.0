#include <iostream>
#include <fstream>
#include <iomanip>
#include <vector>
#include <map>
#include <string>
#include <wtypes.h>
#include "CompDocObj.h"

using namespace std;

// ������� ��� ������ ������� ������ �� ����� EntryName
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
        wcout << L"������ �������� �����!" << endl;
        return 1;
    }

    inFile.seekg(0, ios_base::end);
    size_t fileSize = inFile.tellg();
    wcout << L"������ �����: " << fileSize << L" ����" << endl;

    BYTE *dataBuffer = new BYTE[fileSize];
    inFile.seekg(0, ios_base::beg);
    inFile.read(reinterpret_cast<char*>(dataBuffer), fileSize);
    if (inFile.fail()) {
        wcout << L"������ ��� ������ �����!" << endl;
        delete[] dataBuffer;
        return 1;
    }

    wstring defaultExtension;
    vector<CompoundDocumentObject::CompoundDocument_DirectoryEntryStruct> Directory;
    map<wstring, ULONGLONG> NameMap;
    vector<BinaryBlock> DataStreams;

    DWORD realDataSize = CompoundDocumentObject::GetCompoundDocumentInfo(dataBuffer, fileSize, defaultExtension, Directory, NameMap, DataStreams);
    wcout << L"������ ����� �����������: " << realDataSize << L" ����" << endl;
    wcout << L"���������� �� ���������: " << defaultExtension << endl << endl;

    // ����� ���������� � ������������� �������
    wcout << L"����������:" << endl;
    for (const auto& entry : Directory) {
        wcout << L"���: " << entry.EntryName << L", ������: " << entry.StreamSize << L" ����" << endl;
    }
    wcout << endl;

    wcout << L"������������� ������� � ��������� �������:" << endl;
    for (const auto& pair : NameMap) {
        wcout << pair.first << L" : " << hex << pair.second << endl;
    }
    wcout << endl;

    // ������� ������ ������ � ������ "PowerPoint Document"
    int targetStreamIndex = findStreamIndexByName(Directory, L"PowerPoint Document");

    // ���������, ��� ������ ������ ������ � ��������� � �������� ������� DataStreams
    if (targetStreamIndex >= 0 && targetStreamIndex < static_cast<int>(DataStreams.size())) {
        const vector<BYTE>& targetStream = DataStreams[targetStreamIndex];
        size_t targetSize = targetStream.size();

        // ������������������ ��� ������
        const BYTE sequence1[] = {0x0F, 0x00, 0xF0, 0x0F};
        const size_t sequence1Size = sizeof(sequence1) / sizeof(sequence1[0]);

        size_t lastSequenceIndex = SIZE_MAX; // ������ ��������� ��������� ������������������

        // ����� ��������� ������������������ 0F 00 F0 0F � ������
        for (size_t i = 0; i <= targetSize - sequence1Size; ++i) {
            if (memcmp(targetStream.data() + i, sequence1, sequence1Size) == 0) {
                lastSequenceIndex = i; // ��������� ������ ���������� ���������
            }
        }

        // ���������, ���� �� ������� ���� �� ���� ���������
        if (lastSequenceIndex != SIZE_MAX) {
            // ��������� ������ ����� ���������� ���������
            if (lastSequenceIndex + sequence1Size < targetSize) {
                BYTE sizeByte1 = targetStream[lastSequenceIndex + sequence1Size];       // ������� ����
                BYTE sizeByte2 = targetStream[lastSequenceIndex + sequence1Size + 1];   // ������� ����

                // ����������� ���� ������ � ���� 16-������ ����� ����� � ���������� ������� ��� little-endian
                uint16_t combinedSize = (static_cast<uint16_t>(sizeByte2) << 8) | static_cast<uint16_t>(sizeByte1);

                // �������������� � ���������� �������� � �����
                wcout << std::dec << L"������ (����� ������������������ 0F 00 F0 0F): " << combinedSize << L" (���������� ������)" << endl;

                // ���������� ������ ����� combinedSize
                if (lastSequenceIndex + sequence1Size + 2 + combinedSize <= targetSize) {
                    const BYTE sequence2[] = {0xA8, 0x0F};
                    const size_t sequence2Size = sizeof(sequence2) / sizeof(sequence2[0]);
                    const BYTE sequence3[] = {0xA0, 0x0F};
                    const size_t sequence3Size = sizeof(sequence3) / sizeof(sequence3[0]);

                    // ����� ������������������ A8 0F � A0 0F � ����������� ������������������
                    for (size_t j = 0; j <= combinedSize - sequence2Size; ++j) {
                        if (memcmp(targetStream.data() + lastSequenceIndex + sequence1Size + 2 + j, sequence2, sequence2Size) == 0) {
                            // ������� ����������, ������� ��� �����, ��������� �� A8 0F
                            if (j + sequence2Size + 1 < combinedSize) {
                                BYTE sizeByte3 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence2Size];       // ������� ����
                                BYTE sizeByte4 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence2Size + 1];   // ������� ����

                                // ����������� ���� ������ � ���� 16-������ ����� ����� � ���������� ������� ��� little-endian
                                uint16_t sizeText = (static_cast<uint16_t>(sizeByte4) << 8) | static_cast<uint16_t>(sizeByte3);

                                // �������������� � ���������� �������� � �����
                                wcout << std::dec << L"������ (����� ������������������ A8 0F): " << sizeText << L" (���������� ������)" << endl;

                                // ���������� ������ ����� sizeText
                                if (j + sequence2Size + 2 + sizeText <= combinedSize) {
                                    wcout << L"������������������ ������ �� A8 0F �� sizeText (� UTF-8):" << endl;

                                    // ��������� ����� � ������ UTF-8 � ������� �������������� �������� �� ������� ������
                                    string utf8String;
                                    for (size_t k = 0; k < sizeText; ++k) {
                                        BYTE byte = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence2Size + 4 + k];

                                        // ���������, �������� �� ���� ������������ ��������
                                        if (isprint(byte)) {
                                            utf8String += static_cast<char>(byte);  // ��������� ���� ��� ������
                                        } else {
                                            utf8String += '\n';  // �������������� ������� �������� �� ������� ������
                                        }
                                    }

                                    // ������� ������ UTF-8
                                    cout << utf8String << endl;
                                } else {
                                    wcout << L"������: ����� �� ������� ������ ��� ���������� ������." << endl;
                                }
                            cout << endl;
                            }
                        }

                        if (memcmp(targetStream.data() + lastSequenceIndex + sequence1Size + 2 + j, sequence3, sequence3Size) == 0) {
                            // ������� ����������, ������� ��� �����, ��������� �� A8 0F
                            const BYTE carriage_return[] = {0x0D}; // ��� enter (carriage return)
                            if (j + sequence3Size + 1 < combinedSize) {
                                BYTE sizeByte5 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence3Size];       // ������� ����
                                BYTE sizeByte6 = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence3Size + 1];   // ������� ����

                                // ����������� ���� ������ � ���� 16-������ ����� ����� � ���������� ������� ��� little-endian
                                uint16_t sizeText = (static_cast<uint16_t>(sizeByte6) << 8) | static_cast<uint16_t>(sizeByte5);

                                // �������������� � ���������� �������� � �����
                                wcout << std::dec << L"������ (����� ������������������ A0 0F): " << sizeText << L" (���������� ������)" << endl;
                                // ���������� ������ ����� sizeText
                                if (j + sequence3Size + 2 + sizeText <= combinedSize) {
                                    wcout << L"������������������ ������ �� A8 0F �� sizeText (� UTF-8):" << endl;

                                    // ��������� ����� � ������ UTF-8 � ������� �������������� �������� �� ������� ������
                                    string utf8String;
                                    for (size_t k = 0; k < sizeText; ++k) {
                                        BYTE byte = targetStream[lastSequenceIndex + sequence1Size + 2 + j + sequence3Size + 4 + k];

                                        // ���������, �������� �� ���� carriage return � �������� ��� �� '\n'
                                        if (byte == carriage_return[0]) {
                                            utf8String += '\n';
                                        }

                                        else if (isprint(byte)) {
                                            utf8String += static_cast<char>(byte);  // ��������� ���� ��� ������
                                        }
                                    }

                                    // ������� ������ UTF-8
                                    cout << utf8String << endl;
                                }
                            }
                        }
                    }
                } else {
                    wcout << L"������: ����� �� ������� ������ ��� ���������� ������." << endl;
                }
            } else {
                wcout << L"��������� ����� ����� ��������� ������������������ ������� �� ������� ������." << endl;
            }
        } else {
            wcout << L"������: ������������������ 0F 00 F0 0F �� �������." << endl;
        }
    } else {
        wcout << L"������: ����� � ������ 'PowerPoint Document' �� ������." << endl;
    }

    delete[] dataBuffer;

    return 0;
}

