#ifndef COMPDOCOBJ_H_INCLUDED
#define COMPDOCOBJ_H_INCLUDED

typedef std::vector<BYTE> BinaryBlock;

namespace CompoundDocumentObject {

#pragma pack(push, 1)

typedef struct
{
	ULARGE_INTEGER TimeStamp;
	WORD ClockSerialNumber;
	BYTE MAC[6];

} UUIDStruct, *PUUIDStruct;

typedef struct
{
	BYTE Signature[8]; // {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}
	BYTE Padding1[16];
	WORD Version1; // 0x003E
	WORD Version2; // 0x0003
	WORD ByteOrder; // 0xFFFE - пор€док от старшего к младшему байту, 0xFEFF - пор€док от младшего к старшему байту
	WORD SSZ; // ќпредел€ет размер сектора в файле; SectorSize = 2^SSZ байт
	WORD SSSZ; // ќпредел€ет размер короткого сектора в файле: ShortSectorSize = 2^SSSZ байт
	BYTE Padding2[10];
	DWORD SAT_SizeInSectors; // ќбщее количество секторов, которые используютс€ дл€ составлени€ таблицы распределени€ сектора
	LONG DirectoryStreamSecId; // »дентификатор первого сектора каталога Directory Entries
	BYTE Padding3[4];
	LONG StdStreamMinSize; // ћинимальный размер стандартного потока
	LONG SSAT_SecId; // »дентификатор первого сектора таблицы размещени€ коротких секторов
	DWORD SSAT_SizeInSectors; // ќбщее количество секторов, используемых дл€ таблицы размещени€ коротких секторов
	LONG MSAT_SecId; // »дентификатор первого сектора таблицы распределени€ мастер-сектора
	DWORD MSAT_SizeInSectors; // ќбщее количество секторов, используемых дл€ таблицы распределени€ мастер-сектора
	LONG MSAT[109]; // ѕерва€ часть таблицы распределени€ мастер-сектора, в которой содержитс€ 109 записей SectorId

} CompoundDocumentObjectHeaderStruct, *PCompoundDocumentObjectHeaderStruct;

typedef struct
{
	WCHAR EntryName[32];
	WORD NameLength;
	BYTE Type; // “ип записи: 0x00-Empty, 0x01-User storage, 0x02-User stream, 0x03-Lock Bytes, 0x04-Property, 0x05-Root storage
	BYTE Color; // ÷вет записи: 0x00-красный, 0x01-черный
	LONG LeftChildDirId; // DirId левого дочернего узла внутри красно-черного дерева всех пр€мых членов родительского хранени€
	LONG RightChildDirId; // DirId правого дочернего узла внутри красно-черного дерева всех пр€мых членов родительского хранени€
	LONG RootDirId; // DirId входа в корневой узел красно-черного дерева всех членов хранилища
	UUIDStruct UUID;
	DWORD UserFlags;
	ULARGE_INTEGER TimeCreate; // ¬рем€ создани€ записи
	ULARGE_INTEGER TimeModify; // ¬рем€ последней модификации записи
	LONG StreamFirstSectorId; // SecId первого сектора или короткого сектора, если эта запись относитс€ к потоку
	LONG StreamSize; // ќбщий размер потока в байтах, если запись относитс€ к потоку
	BYTE Padding[4];

} CompoundDocument_DirectoryEntryStruct, *PCompoundDocument_DirectoryEntryStruct;

#pragma pack(pop)

bool GetStreamOffsets(
		LONG firstSecId, LONG maxSecId,
		const LONG *SAT, DWORD SAT_SizeInRecords,
		DWORD sectorSize, bool shortOffsets,
		std::vector<LONG> &streamOffsets
	);

DWORD ReadStreamDataByOffsets(
		const BYTE *documentData,
		DWORD documentDataSize,
		std::vector<LONG> streamOffsets,
		DWORD sectorSize,
		DWORD streamSize,
		BYTE *outDataBuffer,
		DWORD outDataBufferSize
	);

DWORD ReadStreamDataByOffsets(
		const BYTE *documentData,
		DWORD documentDataSize,
		std::vector<LONG> streamOffsets,
		DWORD sectorSize,
		DWORD streamSize,
		BYTE *outDataBuffer,
		DWORD outDataBufferSize
	);

DWORD GetCompoundDocumentInfo(const BYTE *documentData, DWORD dataSize,
		std::wstring &defaultExtension,
		std::vector<CompoundDocument_DirectoryEntryStruct> &Directory,
		std::map<std::wstring, ULONGLONG> &NameMap,
		std::vector<BinaryBlock> &dataStreams
	);

}

#endif // COMPDOCOBJ_H_INCLUDED
