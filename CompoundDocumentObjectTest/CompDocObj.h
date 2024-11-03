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
	WORD ByteOrder; // 0xFFFE - ������� �� �������� � �������� �����, 0xFEFF - ������� �� �������� � �������� �����
	WORD SSZ; // ���������� ������ ������� � �����; SectorSize = 2^SSZ ����
	WORD SSSZ; // ���������� ������ ��������� ������� � �����: ShortSectorSize = 2^SSSZ ����
	BYTE Padding2[10];
	DWORD SAT_SizeInSectors; // ����� ���������� ��������, ������� ������������ ��� ����������� ������� ������������� �������
	LONG DirectoryStreamSecId; // ������������� ������� ������� �������� Directory Entries
	BYTE Padding3[4];
	LONG StdStreamMinSize; // ����������� ������ ������������ ������
	LONG SSAT_SecId; // ������������� ������� ������� ������� ���������� �������� ��������
	DWORD SSAT_SizeInSectors; // ����� ���������� ��������, ������������ ��� ������� ���������� �������� ��������
	LONG MSAT_SecId; // ������������� ������� ������� ������� ������������� ������-�������
	DWORD MSAT_SizeInSectors; // ����� ���������� ��������, ������������ ��� ������� ������������� ������-�������
	LONG MSAT[109]; // ������ ����� ������� ������������� ������-�������, � ������� ���������� 109 ������� SectorId

} CompoundDocumentObjectHeaderStruct, *PCompoundDocumentObjectHeaderStruct;

typedef struct
{
	WCHAR EntryName[32];
	WORD NameLength;
	BYTE Type; // ��� ������: 0x00-Empty, 0x01-User storage, 0x02-User stream, 0x03-Lock Bytes, 0x04-Property, 0x05-Root storage
	BYTE Color; // ���� ������: 0x00-�������, 0x01-������
	LONG LeftChildDirId; // DirId ������ ��������� ���� ������ ������-������� ������ ���� ������ ������ ������������� ��������
	LONG RightChildDirId; // DirId ������� ��������� ���� ������ ������-������� ������ ���� ������ ������ ������������� ��������
	LONG RootDirId; // DirId ����� � �������� ���� ������-������� ������ ���� ������ ���������
	UUIDStruct UUID;
	DWORD UserFlags;
	ULARGE_INTEGER TimeCreate; // ����� �������� ������
	ULARGE_INTEGER TimeModify; // ����� ��������� ����������� ������
	LONG StreamFirstSectorId; // SecId ������� ������� ��� ��������� �������, ���� ��� ������ ��������� � ������
	LONG StreamSize; // ����� ������ ������ � ������, ���� ������ ��������� � ������
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
