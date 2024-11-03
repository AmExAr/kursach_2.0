#include <iostream>
#include <vector>
#include <map>
#include <cassert>
#include <wtypes.h>

#include "CompDocObj.h"

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
