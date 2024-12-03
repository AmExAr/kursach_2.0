#include "pptElements.h"
#include <iostream>
#include <stdio.h>
#include <vector>
#include <map>
#include <cassert>
#include <wtypes.h>



namespace CompoundDocumentObject {

using namespace std;

bool GetStreamOffsets(
		LONG firstSecId, LONG maxSecId,
		const LONG *SAT, DWORD SAT_SizeInRecords,
		DWORD sectorSize, bool shortOffsets,
		vector<LONG> &streamOffsets
	)
{
	assert(maxSecId >= 0);
	assert(streamOffsets.empty());

	LONG currentSecId = firstSecId;
	size_t streamOffsetsSize = 0;
	LONG prevOffset = -1;

	do
	{
		// assert(currentSecId >= 0);
		// assert(currentSecId <= maxSecId);

		if(currentSecId < 0)
		{
			return false;
		}

		if(currentSecId > maxSecId && currentSecId < LONG(SAT_SizeInRecords))
		{
			// Специальный запас на слегка обрезанные файлы
		}
		else if(currentSecId > maxSecId)
		{
			return false;
		}

		size_t streamOffsetsSize = streamOffsets.size();
		if(streamOffsetsSize > size_t(maxSecId))
		{
			return false; // Не может быть цепочки длиннее, чем сама таблица SAT
		}

		LONG currentOffset = currentSecId*sectorSize;
		if(!shortOffsets) currentOffset += 512;
		if(currentOffset == prevOffset)
		{
			return false; // Зацикливание
		}

		streamOffsets.push_back(currentOffset);
		streamOffsetsSize++;
		currentSecId = SAT[currentSecId];
		prevOffset = currentOffset;

	} while(currentSecId >= 0); // -2 - штатный конец цепочки, остальные случаи могут встречаться в поврежденных документах

	return true;
}

DWORD ReadStreamDataByOffsets(
		const BYTE *documentData,
		DWORD documentDataSize,
		vector<LONG> streamOffsets,
		DWORD sectorSize,
		DWORD streamSize,
		BYTE *outDataBuffer,
		DWORD outDataBufferSize
	)
{
	DWORD outDataPosition = 0;

	for(vector<LONG>::const_iterator i=streamOffsets.begin(); i!=streamOffsets.end(); i++)
	{
		if(outDataPosition >= streamSize)	break;

		LONG currentOffset = *i;
		if(currentOffset >= documentDataSize)	break;

		DWORD readSize = min(sectorSize, streamSize-outDataPosition);

		// assert(currentOffset + readSize <= documentDataSize);
		if((currentOffset + readSize) > documentDataSize)
		{
			// Уменьшаем количество данных, которые надо считать
			readSize = documentDataSize-currentOffset;
		}

		// assert(outDataPosition + readSize <= outDataBufferSize);
		if((outDataPosition + readSize) > outDataBufferSize)
		{
			// Уменьшаем количество данных, которые надо считать
			readSize = outDataBufferSize-outDataPosition;
		}

		if(readSize > 0)
		{
			assert(currentOffset+readSize <= documentDataSize);
			assert(outDataPosition+readSize <= outDataBufferSize);
			memcpy(&outDataBuffer[outDataPosition], &documentData[currentOffset], readSize);
			outDataPosition += readSize;
		}
		else
		{
			break;
    }
	}

	return outDataPosition;
}

DWORD GetCompoundDocumentInfo(const BYTE *documentData, DWORD dataSize,
		wstring &defaultExtension,
		std::vector<CompoundDocument_DirectoryEntryStruct> &Directory,
		std::map<wstring, ULONGLONG> &NameMap,
		vector<BinaryBlock> &dataStreams
	)
{
	// Предполагается, что в буфере находится КОРРЕКТНЫЙ документ с обратным порядком байтов

	PCompoundDocumentObjectHeaderStruct documentHeader = (PCompoundDocumentObjectHeaderStruct)documentData;

	if(documentHeader->ByteOrder != 0xFFFE) return 0;

	DWORD sectorSize = 1 << documentHeader->SSZ;
	DWORD shortSectorSize = 1 << documentHeader->SSSZ;
	const LONG TestMaxSecId = min(DWORD(documentHeader->SAT_SizeInSectors*sectorSize/sizeof(LONG)-1), DWORD((dataSize-512)/sectorSize));

	LONG *SAT_Sectors = new LONG[documentHeader->SAT_SizeInSectors];
	unsigned int SAT_Sectors_CurrentPosition = 0;

	// Считываем первые 109 записей MSAT, чтобы сформировать SAT

	for(WORD i=0; i<109; i++)
	{
		LONG secId = documentHeader->MSAT[i];
		if(secId > TestMaxSecId) return 0; // В буфер считаны не все данные

		if(secId >= 0)
		{
			SAT_Sectors[SAT_Sectors_CurrentPosition++] = secId;
		}
	}

	if(documentHeader->MSAT_SecId != -2)
	{
		LONG currentSecId = documentHeader->MSAT_SecId;
		WORD numberOfRecords = sectorSize/sizeof(LONG)-1;
		for(DWORD n=0; n<documentHeader->MSAT_SizeInSectors; n++)
		{
			if((currentSecId < -4) || (currentSecId > TestMaxSecId))
			{
				delete[] SAT_Sectors;
				return 0;
			}

			// Считать очередной сектор SAT
			LONG *SAT_Sectors_Part = (LONG*)&documentData[512 + currentSecId*sectorSize];
			for(WORD i=0; (i<numberOfRecords) && (SAT_Sectors_CurrentPosition < documentHeader->SAT_SizeInSectors); i++)
			{
				LONG SAT_SecId = SAT_Sectors_Part[i];

				// Проверка корректности документа
				if((SAT_SecId < -4) || (SAT_SecId > TestMaxSecId))
				{
					delete[] SAT_Sectors;
					return 0;
				}

				SAT_Sectors[SAT_Sectors_CurrentPosition++] = SAT_SecId;
			}

			currentSecId = SAT_Sectors_Part[numberOfRecords];
			if(currentSecId == -2) break;
		}
	}

	// Считываем SAT

	DWORD SAT_SizeInRecords = sectorSize/sizeof(LONG)*documentHeader->SAT_SizeInSectors;
	LONG *SAT = new LONG[SAT_SizeInRecords];
	int SAT_CurrentPosition = 0;

	for(DWORD n=0; n<documentHeader->SAT_SizeInSectors; n++)
	{
		DWORD SAT_PartOffset = 512 + SAT_Sectors[n]*sectorSize;
		LONG *SAT_Part = (LONG*)&documentData[SAT_PartOffset];
		memcpy((BYTE*)&SAT[SAT_CurrentPosition], (BYTE*)SAT_Part, sectorSize);
		SAT_CurrentPosition += sectorSize/sizeof(LONG);
	}

	// Просматриваем SAT от начала до конца и ищем самую дальнюю занятую запись
	// LONG maxSecId = SAT_SizeInRecords-1; // Если надо быстро, но с запасом

	LONG maxSecId = 0; // Медленно и точно (не учитываем свободные записи)

	if(maxSecId == 0)
	{
		for(DWORD i=0; i<SAT_SizeInRecords; i++)
		{
			LONG secId = SAT[i];

			// Проверка корректности документа
			if((secId < -4) || (secId > TestMaxSecId))
			{
				delete[] SAT;
				delete[] SAT_Sectors;
				return 0;
			}

			if(secId >= 0)
			{
				maxSecId = (secId > maxSecId ? secId : maxSecId);
				maxSecId = (LONG(i) > maxSecId ? i : maxSecId);
			}
			else if((secId == -2 || secId == -3 || secId == -4) && LONG(i) > maxSecId)
			{
				maxSecId = i;
			}
		}
	}

	// Считываем SSAT

	LONG *SSAT_Sectors = new LONG[documentHeader->SSAT_SizeInSectors];
	unsigned int SSAT_Sectors_CurrentPosition = 0;

	DWORD SSAT_SizeInRecords = sectorSize/sizeof(LONG)*documentHeader->SSAT_SizeInSectors;
	LONG *SSAT = new LONG[SSAT_SizeInRecords];
	int SSAT_CurrentPosition = 0;
	LONG SSAT_CurrentSector = documentHeader->SSAT_SecId;
	LONG SSAT_MaxSecId = documentHeader->SSAT_SecId+documentHeader->SSAT_SizeInSectors-1;
	if(SSAT_MaxSecId > maxSecId) maxSecId = SSAT_MaxSecId;

	for(DWORD n=0; n<documentHeader->SSAT_SizeInSectors; n++)
	{
		DWORD SSAT_PartOffset = 512 + SSAT_CurrentSector*sectorSize;
		LONG *SSAT_Part = (LONG*)&documentData[SSAT_PartOffset];
		memcpy((BYTE*)&SSAT[SSAT_CurrentPosition], (BYTE*)SSAT_Part, sectorSize);
		SSAT_CurrentPosition += sectorSize/sizeof(LONG);
		SSAT_CurrentSector = SAT[SSAT_CurrentSector];
	}

	// Просматриваем "каталог"

	bool documentTypeDefined = false;
	LONG nextDirectorySector = documentHeader->DirectoryStreamSecId;
	int i = 0;
	bool isError = false;
	BYTE *shortStreamData = NULL;
	DWORD shortStreamDataSize = 0;

	do
	{
		if(nextDirectorySector <= 0)
		{
			// По идее, такой ситуации быть не должно
			isError = true;
			break;
		}

		DWORD directoryStartPos = 512 + nextDirectorySector*sectorSize;
		for(DWORD currentDirectoryPos = 0; currentDirectoryPos < sectorSize; currentDirectoryPos += 128)
		{
			PCompoundDocument_DirectoryEntryStruct directoryEntry = (PCompoundDocument_DirectoryEntryStruct)&documentData[directoryStartPos+currentDirectoryPos];
			if(directoryEntry->Type == 0)
			{
				continue;
			}

			if(directoryEntry->NameLength <= 0)
			{
				isError = true;
				break;
			}
			else if(i == 0 && directoryEntry->Type != 5)
			{
				isError = true;
				break;
			}
			else if(directoryEntry->StreamSize < 0)
			{
				isError = true;
				break;
			}
			else if(directoryEntry->StreamSize >= documentHeader->StdStreamMinSize && directoryEntry->StreamFirstSectorId > maxSecId)
			{
				isError = true;
				break;
			}
			else if(directoryEntry->StreamFirstSectorId < -2)
			{
				isError = true;
				break;
			}
			else if(!documentTypeDefined && wcscmp(directoryEntry->EntryName, L"WordDocument") == 0)
			{
				defaultExtension = L"doc";
				documentTypeDefined = true;
			}
			else if(!documentTypeDefined && wcscmp(directoryEntry->EntryName, L"Workbook") == 0)
			{
				defaultExtension = L"xls";
				documentTypeDefined = true;
			}
			else if(!documentTypeDefined && wcscmp(directoryEntry->EntryName, L"PowerPoint Document") == 0)
			{
				defaultExtension = L"ppt";
				documentTypeDefined = true;
			}
			else if(!documentTypeDefined && wcscmp(directoryEntry->EntryName, L"DestList") == 0)
			{
				defaultExtension = L"automaticDestinations-ms";
				documentTypeDefined = true;
			}

			vector<LONG> streamOffsets;
			DWORD readSize = 0;
			if(directoryEntry->StreamSize > 0)
			{
				if(i == 0) // Root Entry
				{
					assert(directoryEntry->Type == 5);
					shortStreamDataSize = directoryEntry->StreamSize;
					DWORD shortStreamBufferSize = (shortStreamDataSize+sectorSize-1)/sectorSize*sectorSize;
					shortStreamData = new BYTE[shortStreamBufferSize];
					streamOffsets.reserve((shortStreamDataSize+sectorSize-1)/sectorSize);

					bool resultOk = GetStreamOffsets(directoryEntry->StreamFirstSectorId, maxSecId,
							SAT, SAT_SizeInRecords, sectorSize, false, streamOffsets);

					if(resultOk && LONG(streamOffsets.size()*sectorSize) >= directoryEntry->StreamSize)
					{
						readSize = ReadStreamDataByOffsets(
								documentData, dataSize,
								streamOffsets, sectorSize,
								directoryEntry->StreamSize,
								shortStreamData, shortStreamBufferSize
							);

						assert(LONG(readSize) == directoryEntry->StreamSize);
					}

					if(resultOk && LONG(readSize) == directoryEntry->StreamSize)
					{
						Directory.push_back(*directoryEntry);
						NameMap[wstring(directoryEntry->EntryName, directoryEntry->NameLength/sizeof(WCHAR)-1)] = i;

						// Выводим данные потока
						dataStreams.push_back(BinaryBlock());

						i++;
					}
					else
					{
						isError = true;
						break;
          }
				}
				else
				{
					bool resultOk = false;
					DWORD bufferSize = (directoryEntry->StreamSize+sectorSize-1)/sectorSize*sectorSize;
					BinaryBlock streamData;
					if(directoryEntry->StreamSize < documentHeader->StdStreamMinSize)
					{
						LONG maxShortSectorId = shortStreamDataSize/shortSectorSize-1;
						streamOffsets.reserve((directoryEntry->StreamSize+shortSectorSize-1)/shortSectorSize);
						streamData.resize((directoryEntry->StreamSize+shortSectorSize-1)/shortSectorSize*shortSectorSize, 0);

						resultOk = GetStreamOffsets(directoryEntry->StreamFirstSectorId, maxShortSectorId,
								SSAT, SAT_SizeInRecords, shortSectorSize, true, streamOffsets);

						if(resultOk && LONG(streamOffsets.size()*shortSectorSize) >= directoryEntry->StreamSize)
						{
							readSize = ReadStreamDataByOffsets(
									shortStreamData, shortStreamDataSize,
									streamOffsets, shortSectorSize,
									directoryEntry->StreamSize,
									&streamData.front(), streamData.size()
								);

							assert(LONG(readSize) == directoryEntry->StreamSize);
						}
					}
					else
					{
						streamOffsets.reserve((directoryEntry->StreamSize+sectorSize-1)/sectorSize);
						streamData.resize((directoryEntry->StreamSize+sectorSize-1)/sectorSize*sectorSize, 0);

						resultOk = GetStreamOffsets(directoryEntry->StreamFirstSectorId, maxSecId,
								SAT, SAT_SizeInRecords, sectorSize, false, streamOffsets);

						if(resultOk && LONG(streamOffsets.size()*sectorSize) >= directoryEntry->StreamSize)
						{
							readSize = ReadStreamDataByOffsets(
									documentData, dataSize,
									streamOffsets, sectorSize,
									directoryEntry->StreamSize,
									&streamData.front(), streamData.size()
								);

							assert(LONG(readSize) == directoryEntry->StreamSize);
						}
					}

					if(resultOk && LONG(readSize) == directoryEntry->StreamSize)
					{
						Directory.push_back(*directoryEntry);
						NameMap[wstring(directoryEntry->EntryName, directoryEntry->NameLength/sizeof(WCHAR)-1)] = i;

						// Выводим данные потока
						streamData.resize(directoryEntry->StreamSize); // Обрезаем лишнее
						dataStreams.push_back(streamData);

						i++;
					}
				}
			}
		}

		if(isError) break;
		nextDirectorySector = SAT[nextDirectorySector];

	}	while(nextDirectorySector != -2);

	if(shortStreamData) delete[] shortStreamData;

	delete[] SAT;
	delete[] SAT_Sectors;

	delete[] SSAT;
	delete[] SSAT_Sectors;

	if(!isError)
	{
		return 512 + (maxSecId+1)*sectorSize;
	}
	else
	{
		return 0;
	}
}

} // CompoundDocumentObject

/**
Констурктор класса
\param filePath Путь к объекту

Открытие файла презентации в режиме бинарного чтения:
\code
FILE *file = _wfopen(filePath, L"rb")
\endcode


\code
if(file == NULL)
    {
        std::wcout << L"Файл не открылся!!!" << std::endl;
        return;
    }
\endcode

Переходим в конец файла, получаем его размер и переходим в начало файла
\code
    _fseeki64(file, 0, SEEK_END);
    size_t fileSize = _ftelli64(file);
    _fseeki64(file, 0, SEEK_SET);
\endcode

Массив, в которы загружаем весь файл
\code
BYTE *buffer = new BYTE[fileSize];
\endcode

Записываем в буффер каждый элемент. Если ошибка, то выводим сообщение
\code
if(fread(buffer, 1, fileSize, file) != fileSize)
    {
        std::wcout << L"Ошибка чтения файла" << std::endl;
        return;
    }
\endcode

Закрытие файла
\code
fclose (file);
\endcode

Создаем переменную для хранения имени потока и массив для хранения потока, далее сопоставляем массив и имя потока. DataStreams - массив из массивов потоков
\code
std::wstring defaultExtension;
std::vector<CompoundDocumentObject::CompoundDocument_DirectoryEntryStruct> Directory;
std::map<std::wstring, ULONGLONG> NameMap;
std::vector<BinaryBlock> DataStreams;
\endcode

Разбор PPT файла
\code
DWORD realDataSize = CompoundDocumentObject::GetCompoundDocumentInfo(buffer, fileSize, defaultExtension, Directory, NameMap, DataStreams)
\endcode

Переменные для хранения индекса потока
\code
DWORDLONG powerPointDocumentIndex
DWORDLONG currentUserIndex
DWORDLONG picturesIndex
\endcode

Перебор всех потоков
\code
for(std::map<std::wstring, ULONGLONG>::const_iterator i=NameMap.begin(); i!=NameMap.end(); i++)
    {
        if(i->first == L"PowerPoint Document")
        {
            powerPointDocumentIndex = i->second;
        }
        else if(i->first == L"Current User")
        {
            currentUserIndex = i->second;
        }
        else if(i->first == L"Pictures")
        {
            picturesIndex = i->second;
            Pictures = DataStreams[picturesIndex];
        }
    }
\endcode

Взятие из DataStreams массива необходимого потока
\code
BinaryBlock CurrentUser = DataStreams[currentUserIndex]
PowerPointDocument = DataStreams[powerPointDocumentIndex]
\endcode

Применение к потоку CureentUser соответствующую структуру
\code
CurrentUserAtom *user = (CurrentUserAtom*)CurrentUser.data();
\endcode

Присваивание смещения к последнему изменению
\code
offsetToCurrentEdit = user->offsetToCurrentEdit;
\endcode

Очистка памяти
\code
delete[] buffer;
Directory.clear();
NameMap.clear();
DataStreams.clear();
\endcode
*/
PPT::PPT(const wchar_t *filePath)
{
    FILE *file = _wfopen(filePath, L"rb");

    if(file == NULL)
    {
        std::wcout << L"Файл не открылся!!!" << std::endl;
        return;
    }

    _fseeki64(file, 0, SEEK_END);
    size_t fileSize = _ftelli64(file);
    _fseeki64(file, 0, SEEK_SET);

    BYTE *buffer = new BYTE[fileSize];

    if(fread(buffer, 1, fileSize, file) != fileSize)
    {
        std::wcout << L"Ошибка чтения файла" << std::endl;
        return;
    }

    fclose (file);

    std::wstring defaultExtension;
    std::vector<CompoundDocumentObject::CompoundDocument_DirectoryEntryStruct> Directory;
	std::map<std::wstring, ULONGLONG> NameMap;
	std::vector<BinaryBlock> DataStreams;

    // Разобрать файл
	DWORD realDataSize = CompoundDocumentObject::GetCompoundDocumentInfo(buffer, fileSize, defaultExtension, Directory, NameMap, DataStreams);

    ULONGLONG powerPointDocumentIndex = NULL;
    ULONGLONG currentUserIndex = NULL;
    ULONGLONG picturesIndex = NULL;

    for(std::map<std::wstring, ULONGLONG>::const_iterator i=NameMap.begin(); i!=NameMap.end(); i++)
    {
        if(i->first == L"PowerPoint Document")
        {
            powerPointDocumentIndex = i->second;
        }
        else if(i->first == L"Current User")
        {
            currentUserIndex = i->second;
        }
        else if(i->first == L"Pictures")
        {
            picturesIndex = i->second;
            Pictures = DataStreams[picturesIndex];
        }
    }

    if(picturesIndex == NULL || powerPointDocumentIndex == NULL)
    {
        std::wcout << L"программа при разборе файла не нашла поток Current User" << std::endl;
        return;
    }

    BinaryBlock CurrentUser = DataStreams[currentUserIndex];
    PowerPointDocument = DataStreams[powerPointDocumentIndex];

    CurrentUserAtom *user = (CurrentUserAtom*)CurrentUser.data();


    offsetToCurrentEdit = user->offsetToCurrentEdit;


    // Очистить память
    delete[] buffer;
    Directory.clear();
    NameMap.clear();
    DataStreams.clear();
}

/**
Декструктор PPT файла
*/
PPT::~PPT()
{
    PowerPointDocument.clear();
    Pictures.clear();
}

/**
Получение указателя на поток PowerPointDocument
\code
BYTE *buffer = PowerPointDocument.data();
\endcode

Указатель на начало потока
\code
BYTE *powerPointDocumentBegin = buffer;
\endcode

Переход к записи текущего изменения и примененение структуры
\code
buffer = powerPointDocumentBegin + offsetToCurrentEdit;
UserEditAtom *userEdit = (UserEditAtom*)buffer;
\endcode

Переход к DirectoryEntry и применение структуры
\code
buffer = powerPointDocumentBegin + userEdit->offsetPersistDirectory + sizeof(RecordHeader);
persistIdAndcPersist *idAndCount = (persistIdAndcPersist*)buffer;
\endcode

Получение записи cPersist и преобразнование записей в LittleEndian
\code
DWORD cPersist = (*idAndCount >> 12) & 0xFFFFF;
cPersist = (cPersist & 0xFF) << 8 | (cPersist & 0xFF00) >>  8;
\endcode

Составление таблицы смещений к сохраненным объектам
\code
DWORD *directoryEntry = (DWORD*)(buffer + sizeof(persistIdAndcPersist));
\endcode

Переход к DocumentContainer пропуская RecordHeader
\code
buffer = powerPointDocumentBegin + directoryEntry[0] + sizeof(RecordHeader);
\endcode

Применение структуры RecordHeader
\code
RecordHeader *rh = (RecordHeader*)buffer;
\endcode

Поиск типа записи, соотвествующей контейнеру с текстом (в MS PPT 1992-2003)
\code
while(rh->recType != (WORD)RecordTypeEnum::RT_SlideListWithText || rh->recVerAndInstance != 0x000F)
{
    buffer += rh->recLen + sizeof(RecordHeader);
    rh = (RecordHeader*)buffer;
}
\endcode

Вход в SlideListWithText
\code
while(rh->recType != (WORD)RecordTypeEnum::RT_SlideListWithText || rh->recVerAndInstance != 0x000F)
{
    buffer += sizeof(RecordHeader);
    rh = (RecordHeader*)buffer;
}
\endcode



\parma bool isTextHere переменная для проверки на наличие текста

Открытие фалйа в бинарной записи
\code
FILE *file = _wfopen(filePath, L"wb")
\endcode
\parma DWORD slideNumber переменнная для подсчета слайдов

Перебор всех элементов до конца DocumentContainer
\code
while(rh->recType != (WORD)RecordTypeEnum::RT_EndDocumentAtom)
\endcode

Если находится тип, соответствующий SlidePersistAtom, то увеличивается номер слайда и этот номер записывается
\code
if(rh->recType == (WORD)RecordTypeEnum::RT_SlidePersistAtom && rh->recVerAndInstance == 0x0000)
{
    fwprintf(file, L"Слайд номер %u: \n", ++slideNumber);
}
\endcode

Если находится тип, соответствующий тексту, то isTextHere = true, после чего если текст хранится в UTF-16, то делаем указатель на начало текста и записываем это в файл
\code
else if((rh->recType == (WORD)RecordTypeEnum::RT_TextCharsAtom || rh->recType == (WORD)RecordTypeEnum::RT_TextBytesAtom) &&
            rh->recVerAndInstance == 0x0000)
        {
            isTextHere = true;

            if(rh->recType == (WORD)RecordTypeEnum::RT_TextCharsAtom)
            {
                WORD *text = (WORD*)(buffer + sizeof(RecordHeader));

                fwrite(text, 2, (rh->recLen)/2, file);
                fwprintf(file, L"\n\n");
            }
            else
            {
                BYTE *text = buffer + sizeof(RecordHeader);

                WORD *chars = (WORD*)(TextBytesToChars(text, rh->recLen));

                fwrite(chars, 2, rh->recLen, file);
                fwprintf(file, L"\n\n");

                delete[] chars;
            }
        }
\endcode

Если запись не была найдена, то переходим к другой записи
\code
buffer += sizeof(RecordHeader) + rh->recLen;
rh = (RecordHeader*)buffer;
\endcode

Закрытие файла, вывод надписи, если текст был найден. Иначе используем другой алгоритм поиска (для MS PPT 2007 и позже)
\code
fclose(file);
if(isTextHere)
{
    std::wcout << L"Текст из презентации был успешно сохранён в ваш файл!" << std::endl;
}
\endcode

Иначе используем другой алгоритм поиска (для MS PPT 2007 и позже)
Происходит перебор всех элементов из directoryEntry, начиная со второго элемента
Если не первый слайд, то добавляем пропуск строки.
Написание номера слайда,
\code
if(i > 2)
{
    fwprintf(file, L"\n");
}
fwprintf(file, L"Слайд номер %u: \n", (i - 1));
\endcode

Переход к первому сохраненному объекту
\code
buffer = powerPointDocumentBegin + directoryEntry[i];
\encdcode

Применение струкутры RecordHeader. В нем находится размер всего слайда
\code
rh = (RecordHeader*)buffer;
\encdcode

Создать указатель на конец сладйа
\code
BYTE *endOfSlide = buffer + rh->recLen + sizeof(RecordHeader);
\encdcode

Пропуск записи SlideContainer и OfficeArtDgContainer
\code
 buffer += sizeof(SlideContainer) + sizeof(OfficeArtDgContainer);
\encdcode

Повторное применение струкутры RecordHeader
\code
rh = (RecordHeader*)buffer;
\encdcode

С помощью цикла ищем OfficeArtSpContainer
\code
while(rh->recType != (WORD)RecordTypeEnum::RT_OfficeArtSpContainer || rh->recVerAndInstance != 0x000F)
{
    buffer += rh->recLen + sizeof(RecordHeader);
    rh = (RecordHeader*)buffer;
}
\encdcode

Пропуск заголовка
\code
buffer += sizeof(RecordHeader);
rh = (RecordHeader*)buffer;
\encdcode

Поиск OfficeArtClientTextbox
\code
while(rh->recType != (WORD)RecordTypeEnum::RT_OfficeArtClientTextbox || rh->recVerAndInstance != 0x000F)
{
    buffer += rh->recLen + sizeof(RecordHeader);;
    rh = (RecordHeader*)buffer;
}
\encdcode

Пока слайд не закончится, ищем текст в OfficeArtClientTextbox
\code
while(buffer != endOfSlide)
\encdcode

Если текст находится внутри текста, то аналогичным образом перейти к ClientTextBox
\code
if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtSpContainer && rh->recVerAndInstance == 0x000F)
{
    BYTE *endOfSpContainer = buffer + sizeof(RecordHeader) + rh->recLen;

    buffer += sizeof(RecordHeader);
    rh = (RecordHeader*)buffer;            {


                    while(buffer != endOfSpContainer)
                    {
                        buffer += rh->recLen + sizeof(RecordHeader);;
                        rh = (RecordHeader*)buffer;

                        if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtClientTextbox && rh->recVerAndInstance == 0x000F)
                        {
                            buffer += sizeof(RecordHeader);
                            rh = (RecordHeader*)buffer;

                            break;
                        }
                    }
                }

\encdcode

Аналогичные действия, если найден текст
\code
else if((rh->recType == (WORD)RecordTypeEnum::RT_TextCharsAtom || rh->recType == (WORD)RecordTypeEnum::RT_TextBytesAtom) &&
                         rh->recVerAndInstance == 0x0000)
                {
                    if(rh->recType == (WORD)RecordTypeEnum::RT_TextCharsAtom)
                    {
                        WORD *text = (WORD*)(buffer + sizeof(RecordHeader));

                        fwrite(text, 2, (rh->recLen)/2, file);
                        fwprintf(file, L"\n");
                    }
                    else
                    {
                        BYTE *text = buffer + sizeof(RecordHeader);

                        WORD *chars = (WORD*)(TextBytesToChars(text, rh->recLen));

                        fwrite(chars, 2, rh->recLen, file);
                        fwprintf(file, L"\n");

                        delete[] chars;
                    }
                }
                buffer += sizeof(RecordHeader) + rh->recLen;
                rh = (RecordHeader*)buffer;
            }
\endcode

Закрытие файла
\code
fclose(file);
\endcode
*/
void PPT::GetText(const wchar_t *filePath)
{
    BYTE *buffer = PowerPointDocument.data();

    BYTE *powerPointDocumentBegin = buffer;

    buffer = powerPointDocumentBegin + offsetToCurrentEdit;

    UserEditAtom *userEdit = (UserEditAtom*)buffer;

    DWORD directoryEntryCount = userEdit->persistIdSeed;
    DWORD directoryEntry[directoryEntryCount];
    DWORD directoryEntryCurrentElem = directoryEntryCount - 1;
    DWORD lastDocumentContainer = NULL;

    while(userEdit->offsetLastEdit != 0x00000000)
    {
        buffer = powerPointDocumentBegin + userEdit->offsetPersistDirectory + sizeof(RecordHeader);
        persistIdAndcPersist *idAndCount = (persistIdAndcPersist*)buffer;

        DWORD cPersist = (*idAndCount >> 12) & 0xFFFFF;
        cPersist = (cPersist & 0xFF) << 8 | (cPersist & 0xFF00) >>  8;

        DWORD *currentDirectoryEntry = (DWORD*)(buffer + sizeof(persistIdAndcPersist));

        if(lastDocumentContainer == NULL)
        {
            lastDocumentContainer = currentDirectoryEntry[0];
        }

        for(DWORD i = 0; i < cPersist; i++)
        {
            directoryEntry[directoryEntryCurrentElem--] = currentDirectoryEntry[cPersist - i - 1];
        }

        buffer = powerPointDocumentBegin + userEdit->offsetLastEdit;
        userEdit = (UserEditAtom*)buffer;

    }

    buffer = powerPointDocumentBegin + userEdit->offsetPersistDirectory + sizeof(RecordHeader);
    persistIdAndcPersist *idAndCount = (persistIdAndcPersist*)buffer;

    DWORD cPersist = (*idAndCount >> 12) & 0xFFFFF;
    cPersist = (cPersist & 0xFF) << 8 | (cPersist & 0xFF00) >>  8;

    DWORD *currentDirectoryEntry = (DWORD*)(buffer + sizeof(persistIdAndcPersist));

    for(DWORD i = 0; i < cPersist; i++)
    {
        directoryEntry[directoryEntryCurrentElem--] = currentDirectoryEntry[cPersist - i - 1];
    }

    buffer = powerPointDocumentBegin + lastDocumentContainer + sizeof(RecordHeader);

    RecordHeader *rh = (RecordHeader*)buffer;

    while(rh->recType != (WORD)RecordTypeEnum::RT_SlideListWithText || rh->recVerAndInstance != 0x000F)
    {
        buffer += rh->recLen + sizeof(RecordHeader);
        rh = (RecordHeader*)buffer;
    }

    std::wcout << L"Успешно найден SlideListWithText" << std::endl << std::endl;

    buffer += sizeof(RecordHeader);
    rh = (RecordHeader*)buffer;

    bool isTextHere = false;

    FILE *file = _wfopen(filePath, L"wb");

    DWORD slideNumber = 0;

    while(rh->recType != (WORD)RecordTypeEnum::RT_EndDocumentAtom)
    {
        if(rh->recType == (WORD)RecordTypeEnum::RT_SlidePersistAtom && rh->recVerAndInstance == 0x0000)
        {
            fwprintf(file, L"Слайд номер %u: \n", ++slideNumber);
        }

        else if((rh->recType == (WORD)RecordTypeEnum::RT_TextCharsAtom || rh->recType == (WORD)RecordTypeEnum::RT_TextBytesAtom) &&
            rh->recVerAndInstance == 0x0000)
        {
            isTextHere = true;

            if(rh->recType == (WORD)RecordTypeEnum::RT_TextCharsAtom)
            {
                WORD *text = (WORD*)(buffer + sizeof(RecordHeader));

                fwrite(text, 2, (rh->recLen)/2, file);
                fwprintf(file, L"\n\n");
            }
            else
            {
                BYTE *text = buffer + sizeof(RecordHeader);

                WORD *chars = (WORD*)(TextBytesToChars(text, rh->recLen));

                fwrite(chars, 2, rh->recLen, file);
                fwprintf(file, L"\n\n");

                delete[] chars;
            }
        }
        buffer += sizeof(RecordHeader) + rh->recLen;
        rh = (RecordHeader*)buffer;
    }
    std::wcout << L"DocumentContainer успешно просмотрен" << std::endl << std::endl;

    for(DWORD i = 0; i < directoryEntryCount; i++)
    {
        buffer = powerPointDocumentBegin + directoryEntry[i];
        rh = (RecordHeader*)buffer;

        if(rh->recType != (WORD)RecordTypeEnum::RT_Slide || rh->recVerAndInstance != 0x000F)
        {
            continue;
        }

        BYTE *endOfSlide = buffer + rh->recLen + sizeof(RecordHeader);

        buffer += sizeof(SlideContainer);
        rh = (RecordHeader*)buffer;

        while(rh->recType != (WORD)RecordTypeEnum::RT_Drawing || rh->recVerAndInstance != 0x000F)
        {
            buffer += rh->recLen + sizeof(RecordHeader);
            rh = (RecordHeader*)buffer;
        }

        buffer += sizeof(RecordHeader)*2;
        rh = (RecordHeader*)buffer;

        while(buffer != endOfSlide)
        {
            if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtSpgrContainer && rh->recVerAndInstance == 0x000F)
            {
                buffer += sizeof(RecordHeader);
                rh = (RecordHeader*)buffer;
                continue;
            }
            else if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtSpContainer && rh->recVerAndInstance == 0x000F)
            {
                buffer += sizeof(RecordHeader);
                rh = (RecordHeader*)buffer;
                continue;
            }
            else if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtClientTextbox && rh->recVerAndInstance == 0x000F)
            {
                buffer += sizeof(RecordHeader);
                rh = (RecordHeader*)buffer;
                continue;
            }
            else if((rh->recType == (WORD)RecordTypeEnum::RT_TextCharsAtom || rh->recType == (WORD)RecordTypeEnum::RT_TextBytesAtom) &&
                     rh->recVerAndInstance == 0x0000)
            {
                if(rh->recType == (WORD)RecordTypeEnum::RT_TextCharsAtom)
                {
                    WORD *text = (WORD*)(buffer + sizeof(RecordHeader));

                    fwrite(text, 2, (rh->recLen)/2, file);
                    fwprintf(file, L"\n");
                }
                else
                {
                    BYTE *text = buffer + sizeof(RecordHeader);

                    WORD *chars = (WORD*)(TextBytesToChars(text, rh->recLen));

                    fwrite(chars, 2, rh->recLen, file);
                    fwprintf(file, L"\n");

                    delete[] chars;
                }
            }

            buffer += rh->recLen + sizeof(RecordHeader);
            rh = (RecordHeader*)buffer;
        }
    }

    for(DWORD i = 0; i < directoryEntryCount; i++)
    {
        buffer = powerPointDocumentBegin + directoryEntry[i];
        rh = (RecordHeader*)buffer;

        if(rh->recType != (WORD)RecordTypeEnum::RT_Notes || rh->recVerAndInstance != 0x000F)
        {
            continue;
        }

        BYTE *endOfNote = buffer + rh->recLen + sizeof(RecordHeader);

        buffer += sizeof(NotesContainer) + sizeof(RecordHeader)*2;

        rh = (RecordHeader*)buffer;

        while(buffer != endOfNote)
        {
            if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtSpgrContainer && rh->recVerAndInstance == 0x000F)
            {
                buffer += sizeof(RecordHeader);
                rh = (RecordHeader*)buffer;
                continue;
            }
            else if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtSpContainer && rh->recVerAndInstance == 0x000F)
            {
                buffer += sizeof(RecordHeader);
                rh = (RecordHeader*)buffer;
                continue;
            }
            else if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtClientTextbox && rh->recVerAndInstance == 0x000F)
            {
                buffer += sizeof(RecordHeader);
                rh = (RecordHeader*)buffer;
                continue;
            }
            else if((rh->recType == (WORD)RecordTypeEnum::RT_TextCharsAtom || rh->recType == (WORD)RecordTypeEnum::RT_TextBytesAtom) &&
                     rh->recVerAndInstance == 0x0000)
            {
                if(rh->recType == (WORD)RecordTypeEnum::RT_TextCharsAtom)
                {
                    WORD *text = (WORD*)(buffer + sizeof(RecordHeader));

                    fwrite(text, 2, (rh->recLen)/2, file);
                    fwprintf(file, L"\n");
                }
                else
                {
                    BYTE *text = buffer + sizeof(RecordHeader);

                    WORD *chars = (WORD*)(TextBytesToChars(text, rh->recLen));

                    fwrite(chars, 2, rh->recLen, file);
                    fwprintf(file, L"\n");

                    delete[] chars;
                }
            }

            buffer += rh->recLen + sizeof(RecordHeader);
            rh = (RecordHeader*)buffer;
        }
    }

    fclose(file);
    std::wcout << L"Текст из презентации был успешно сохранён в ваш файл!" << std::endl;
}

/**
Функция для преборазования UTF-8 в UTF-16LE
\return chars
\parma BYTE *text текст в UTF-8
\parma DWORD textSize размер текста в UTF-8
*/
BYTE *TextBytesToChars(BYTE *text, DWORD textSize)
{
    BYTE *chars = new BYTE[textSize*2];

    for(DWORD i = 0; i < textSize*2; i++)
    {
        if(i % 2 == 1)
        {
            chars[i] = 0x00;
        }
        else
        {
            chars[i] = text[i/2];
        }
    }

    return chars;
}

void PPT::GetPicture(const wchar_t *filePath)
{
    if(Pictures.size() == 0)
    {
        std::wcout << L"В презентации картинок не обнаружено" << std::endl;
        return;
    }
    if(CreateDirectoryW(filePath, NULL) == 0 && GetLastError() == ERROR_PATH_NOT_FOUND && *filePath != L'\0')
    {
        std::wcout << L"Не удалось создать каталог: один или несколько промежуточных каталогов не существует!" << std::endl;
        return;
    }

    BYTE *buffer = Pictures.data();
    size_t bufferSize = Pictures.size();
    BYTE *endOfBuffer = buffer + bufferSize - 1;

    RecordHeader *rh;

    WORD pictureNumber = 0;
    while(buffer != endOfBuffer)
    {
        if(pictureNumber > 0)
        {
            buffer++;
        }

        rh = (RecordHeader*)buffer;
        DWORD pictureSize;

        if(rh->recVerAndInstance == (WORD)RecordVerAndInstanceEnum::RVAI_recVerAndInstanceDIBOneUUID ||
           rh->recVerAndInstance == (WORD)RecordVerAndInstanceEnum::RVAI_recVerAndInstanceEMFOneUUID ||
           rh->recVerAndInstance == (WORD)RecordVerAndInstanceEnum::RVAI_recVerAndInstanceJPEGInCMYKOneUUID ||
           rh->recVerAndInstance == (WORD)RecordVerAndInstanceEnum::RVAI_recVerAndInstanceJPEGInRGBOneUUID ||
           rh->recVerAndInstance == (WORD)RecordVerAndInstanceEnum::RVAI_recVerAndInstancePICTOneUUID ||
           rh->recVerAndInstance == (WORD)RecordVerAndInstanceEnum::RVAI_recVerAndInstancePNGOneUUID ||
           rh->recVerAndInstance == (WORD)RecordVerAndInstanceEnum::RVAI_recVerAndInstanceTIFFOneUUID ||
           rh->recVerAndInstance == (WORD)RecordVerAndInstanceEnum::RVAI_recVerAndInstanceWMFOneUUID)
        {
            if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtBlipEMF || rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtBlipWMF ||
               rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtBlipPICT)
            {
                buffer += sizeof(OfficeArtBlipMetafileOneUID);
                pictureSize = rh->recLen - sizeof(OfficeArtBlipMetafileOneUID) + sizeof(RecordHeader);
            }
            else
            {
                buffer += sizeof(OfficeArtBlipTagOneUID);
                pictureSize = rh->recLen - sizeof(OfficeArtBlipTagOneUID) + sizeof(RecordHeader);
            }
        }
        else
        {
            if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtBlipEMF || rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtBlipWMF ||
               rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtBlipPICT)
            {
                buffer += sizeof(OfficeArtBlipMetafileOneUID);
                pictureSize = rh->recLen - sizeof(OfficeArtBlipMetafileOneUID) + sizeof(RecordHeader);
            }
            else
            {
                buffer += sizeof(OfficeArtBlipTagOneUID);
                pictureSize = rh->recLen - sizeof(OfficeArtBlipTagOneUID) + sizeof(RecordHeader);
            }
        }

        wchar_t *path;

        if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtBlipDIB)
        {
            path = MakePictureName(filePath, L"dib", ++pictureNumber);
        }
        else if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtBlipEMF)
        {
            path = MakePictureName(filePath, L"emf", ++pictureNumber);
        }
        else if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtBlipJPEG || rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtBlipJPEGException)
        {
            path = MakePictureName(filePath, L"jpeg", ++pictureNumber);
        }
        else if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtBlipPICT)
        {
            path = MakePictureName(filePath, L"pict", ++pictureNumber);
        }
        else if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtBlipPNG)
        {
            path = MakePictureName(filePath, L"png", ++pictureNumber);
        }
        else if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtBlipTIFF)
        {
            path = MakePictureName(filePath, L"tiff", ++pictureNumber);
        }
        else if(rh->recType == (WORD)RecordTypeEnum::RT_OfficeArtBlipWMF)
        {
            path = MakePictureName(filePath, L"wmf", ++pictureNumber);
        }
        else if(rh->recType == 0x0000 && rh->recVerAndInstance == 0x0000)
        {
            std::wcout << L"Поток Pictures содержит пустоты или в конце забит нулями!" << std::endl;
            return;
        }

        std::wcout << path << std::endl;

        FILE *file = _wfopen(path, L"wb");
        fwrite(buffer, 1, pictureSize, file);
        fclose(file);

        buffer += pictureSize - 1;

        delete[] path;
    }
}

wchar_t *MakePictureName(const wchar_t *path, wchar_t *format, WORD number)
{
    BYTE numeralCount = 0;
    WORD num = number;
    WORD mirrorNumber = 0;

    while(num > 0)
    {
        mirrorNumber = mirrorNumber * 10 + num % 10;

        numeralCount++;
        num /= 10;
    }

    BYTE formatLength = 0;
    WORD pathLength = 0;

    while(format[formatLength] != L'\0')
    {
        formatLength++;
    }
    while(path[pathLength] != L'\0')
    {
        pathLength++;
    }

    if(pathLength == 0)
    {
        wchar_t *name = new wchar_t[numeralCount + formatLength + pathLength + 4];

        for(BYTE i = 0; i < numeralCount + formatLength + pathLength + 4; i++)
        {
            if(i == pathLength)
            {
                name[i++] = L'.';
                name[i] = L'\\';
            }
            else if(i > pathLength && i < numeralCount + pathLength + 2)
            {
                name[i] = L'0' + (mirrorNumber % 10);
                mirrorNumber /= 10;
            }
            else if(i == pathLength + numeralCount + 2)
            {
                name[i] = L'.';
            }
            else
            {
                name[i] = format[i - pathLength - numeralCount - 3];
            }
        }
        return name;
    }
    else
    {
        wchar_t *name = new wchar_t[numeralCount + formatLength + pathLength + 3];

        for(BYTE i = 0; i < numeralCount + formatLength + pathLength + 3; i++)
        {
            if(i < pathLength)
            {
                name[i] = path[i];
            }
            else if(i == pathLength)
            {
                name[i] = L'\\';
            }
            else if(i > pathLength && i < numeralCount + pathLength + 1)
            {
                name[i] = L'0' + (mirrorNumber % 10);
                mirrorNumber /= 10;
            }
            else if(i == pathLength + numeralCount + 1)
            {
                name[i] = L'.';
            }
            else
            {
                name[i] = format[i - pathLength - numeralCount - 2];
            }
        }
        return name;
    }
}
