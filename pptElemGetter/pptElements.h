#ifndef pptElementsH
#define pptElementsH
#include <iostream>
#include <iostream>
#include <vector>
#include <map>
#include <cassert>
#include <wtypes.h>


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
	WORD ByteOrder; // 0xFFFE - пор¤док от старшего к младшему байту, 0xFEFF - пор¤док от младшего к старшему байту
	WORD SSZ; // ќпредел¤ет размер сектора в файле; SectorSize = 2^SSZ байт
	WORD SSSZ; // ќпредел¤ет размер короткого сектора в файле: ShortSectorSize = 2^SSSZ байт
	BYTE Padding2[10];
	DWORD SAT_SizeInSectors; // ќбщее количество секторов, которые используютс¤ дл¤ составлени¤ таблицы распределени¤ сектора
	LONG DirectoryStreamSecId; // »дентификатор первого сектора каталога Directory Entries
	BYTE Padding3[4];
	LONG StdStreamMinSize; // ћинимальный размер стандартного потока
	LONG SSAT_SecId; // »дентификатор первого сектора таблицы размещени¤ коротких секторов
	DWORD SSAT_SizeInSectors; // ќбщее количество секторов, используемых дл¤ таблицы размещени¤ коротких секторов
	LONG MSAT_SecId; // »дентификатор первого сектора таблицы распределени¤ мастер-сектора
	DWORD MSAT_SizeInSectors; // ќбщее количество секторов, используемых дл¤ таблицы распределени¤ мастер-сектора
	LONG MSAT[109]; // ѕерва¤ часть таблицы распределени¤ мастер-сектора, в которой содержитс¤ 109 записей SectorId

} CompoundDocumentObjectHeaderStruct, *PCompoundDocumentObjectHeaderStruct;

typedef struct
{
	WCHAR EntryName[32];
	WORD NameLength;
	BYTE Type; // “ип записи: 0x00-Empty, 0x01-User storage, 0x02-User stream, 0x03-Lock Bytes, 0x04-Property, 0x05-Root storage
	BYTE Color; // ÷вет записи: 0x00-красный, 0x01-черный
	LONG LeftChildDirId; // DirId левого дочернего узла внутри красно-черного дерева всех пр¤мых членов родительского хранени¤
	LONG RightChildDirId; // DirId правого дочернего узла внутри красно-черного дерева всех пр¤мых членов родительского хранени¤
	LONG RootDirId; // DirId входа в корневой узел красно-черного дерева всех членов хранилища
	UUIDStruct UUID;
	DWORD UserFlags;
	ULARGE_INTEGER TimeCreate; // ¬рем¤ создани¤ записи
	ULARGE_INTEGER TimeModify; // ¬рем¤ последней модификации записи
	LONG StreamFirstSectorId; // SecId первого сектора или короткого сектора, если эта запись относитс¤ к потоку
	LONG StreamSize; // ќбщий размер потока в байтах, если запись относитс¤ к потоку
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

DWORD GetCompoundDocumentInfo(
		const BYTE *documentData, DWORD dataSize,
		std::wstring &defaultExtension,
		std::vector<CompoundDocument_DirectoryEntryStruct> &Directory,
		std::map<std::wstring, ULONGLONG> &NameMap,
		std::vector<BinaryBlock> &dataStreams
	);

}

typedef uint32_t persistIdAndcPersist;

#pragma pack(push, 1)

typedef struct
{
    WORD recVerAndInstance;
    WORD recType;
    DWORD recLen;

} RecordHeader;

typedef struct
{
    RecordHeader rh;
    DWORD size;
    DWORD headerToken;
    DWORD offsetToCurrentEdit;

} CurrentUserAtom;

typedef struct
{
    RecordHeader rh;
    DWORD lastSlideIdRef;
    WORD version;
    BYTE minorVersion;
    BYTE majorVersion;
    DWORD offsetLastEdit;
    DWORD offsetPersistDirectory;

} UserEditAtom;

typedef struct
{
    RecordHeader rh;
    DWORD persistIdRef;

} SlidePersistAtom;

typedef struct
{
    RecordHeader rh;
    BYTE slideAtom[32];
    BYTE slideShowSlideInfoAtom[24];

} SlideContainer;

typedef struct
{
    RecordHeader rh;
    BYTE drawingData[16];

} OfficeArtDgContainer;

typedef struct
{
    RecordHeader rh;
    BYTE shapeGroup[24];
    BYTE shapeProp[16];
    BYTE deletedShape[12];

} OfficeArtSpContainer;

#pragma pack(pop)

enum class RecordTypeEnum : WORD
{
    RT_OfficeArtSpContainer = 0xF004,
    RT_OfficeArtClientTextbox = 0xF00D,

    RT_Document	= 0x03E8,
    RT_DocumentAtom = 0x03E9,
    RT_EndDocumentAtom = 0x03EA,
    RT_Slide = 0x03EE,
    RT_SlideAtom = 0x03EF,
    RT_Notes = 0x03F0,
    RT_NotesAtom = 0x03F1,
    RT_Environment = 0x03F2,
    RT_SlidePersistAtom = 0x03F3,
    RT_MainMaster = 0x03F8,
    RT_SlideShowSlideInfoAtom = 0x03F9,
    RT_SlideViewInfo = 0x03FA,
    RT_GuideAtom = 0x03FB,
    RT_ViewInfoAtom = 0x03FD,
    RT_SlideViewInfoAtom = 0x03FE,
    RT_VbaInfo = 0x03FF,
    RT_VbaInfoAtom = 0x0400,
    RT_SlideShowDocInfoAtom = 0x0401,
    RT_Summary = 0x0402,
    RT_DocRoutingSlipAtom = 0x0406,
    RT_OutlineViewInfo = 0x0407,
    RT_SorterViewInfo = 0x0408,
    RT_ExternalObjectList = 0x0409,
    RT_ExternalObjectListAtom = 0x040A,
    RT_DrawingGroup = 0x040B,
    RT_Drawing = 0x040C,
    RT_GridSpacing10Atom = 0x040D,
    RT_RoundTripTheme12Atom = 0x040E,
    RT_RoundTripColorMapping12Atom = 0x040F,
    RT_NamedShows = 0x0410,
    RT_NamedShow = 0x0411,
    RT_NamedShowSlidesAtom = 0x0412,
    RT_NotesTextViewInfo9 = 0x0413,
    RT_NormalViewSetInfo9 = 0x0414,
    RT_NormalViewSetInfo9Atom = 0x0415,
    RT_RoundTripOriginalMainMasterId12Atom = 0x041C,
    RT_RoundTripCompositeMasterId12Atom = 0x041D,
    RT_RoundTripContentMasterInfo12Atom = 0x041E,
    RT_RoundTripShapeId12Atom = 0x041F,
    RT_RoundTripHFPlaceholder12Atom =0x0420,
    RT_RoundTripContentMasterId12Atom = 0x0422,
    RT_RoundTripOArtTextStyles12Atom = 0x0423,
    RT_RoundTripHeaderFooterDefaults12Atom = 0x0424,
    RT_RoundTripDocFlags12Atom = 0x0425,
    RT_RoundTripShapeCheckSumForCL12Atom = 0x0426,
    RT_RoundTripNotesMasterTextStyles12Atom = 0x0427,
    RT_RoundTripCustomTableStyles12Atom = 0x0428,
    RT_List = 0x07D0,
    RT_FontCollection = 0x07D5,
    RT_FontCollection10 = 0x07D6,
    RT_BookmarkCollection = 0x07E3,
    RT_SoundCollection = 0x07E4,
    RT_SoundCollectionAtom = 0x07E5,
    RT_Sound = 0x07E6,
    RT_SoundDataBlob = 0x07E7,
    RT_BookmarkSeedAtom = 0x07E9,
    RT_ColorSchemeAtom = 0x07F0,
    RT_BlipCollection9 = 0x07F8,
    RT_BlipEntity9Atom = 0x07F9,
    RT_ExternalObjectRefAtom = 0x0BC1,
    RT_PlaceholderAtom = 0x0BC3,
    RT_ShapeAtom = 0x0BDB,
    RT_ShapeFlags10Atom = 0x0BDC,
    RT_RoundTripNewPlaceholderId12Atom = 0x0BDD,
    RT_OutlineTextRefAtom = 0x0F9E,
    RT_TextHeaderAtom = 0x0F9F,
    RT_TextCharsAtom = 0x0FA0,
    RT_StyleTextPropAtom = 0x0FA1,
    RT_MasterTextPropAtom = 0x0FA2,
    RT_TextMasterStyleAtom = 0x0FA3,
    RT_TextCharFormatExceptionAtom = 0x0FA4,
    RT_TextParagraphFormatExceptionAtom = 0x0FA5,
    RT_TextRulerAtom = 0x0FA6,
    RT_TextBookmarkAtom = 0x0FA7,
    RT_TextBytesAtom = 0x0FA8,
    RT_TextSpecialInfoDefaultAtom = 0x0FA9,
    RT_TextSpecialInfoAtom = 0x0FAA,
    RT_DefaultRulerAtom = 0x0FAB,
    RT_StyleTextProp9Atom = 0x0FAC,
    RT_TextMasterStyle9Atom = 0x0FAD,
    RT_OutlineTextProps9 = 0x0FAE,
    RT_OutlineTextPropsHeader9Atom = 0x0FAF,
    RT_TextDefaults9Atom = 0x0FB0,
    RT_StyleTextProp10Atom = 0x0FB1,
    RT_TextMasterStyle10Atom = 0x0FB2,
    RT_OutlineTextProps10 = 0x0FB3,
    RT_TextDefaults10Atom = 0x0FB4,
    RT_OutlineTextProps11 = 0x0FB5,
    RT_StyleTextProp11Atom = 0x0FB6,
    RT_FontEntityAtom = 0x0FB7,
    RT_FontEmbedDataBlob = 0x0FB8,
    RT_CString = 0x0FBA,
    RT_MetaFile = 0x0FC1,
    RT_ExternalOleObjectAtom = 0x0FC3,
    RT_Kinsoku = 0x0FC8,
    RT_Handout = 0x0FC9,
    RT_ExternalOleEmbed = 0x0FCC,
    RT_ExternalOleEmbedAtom = 0x0FCD,
    RT_ExternalOleLink = 0x0FCE,
    RT_BookmarkEntityAtom = 0x0FD0,
    RT_ExternalOleLinkAtom = 0x0FD1,
    RT_KinsokuAtom = 0x0FD2,
    RT_ExternalHyperlinkAtom = 0x0FD3,
    RT_ExternalHyperlink = 0x0FD7,
    RT_SlideNumberMetaCharAtom = 0x0FD8,
    RT_HeadersFooters = 0x0FD9,
    RT_HeadersFootersAtom = 0x0FDA,
    RT_TextInteractiveInfoAtom = 0x0FDF,
    RT_ExternalHyperlink9 = 0x0FE4,
    RT_RecolorInfoAtom = 0x0FE7,
    RT_ExternalOleControl = 0x0FEE,
    RT_SlideListWithText = 0x0FF0,
    RT_AnimationInfoAtom = 0x0FF1,
    RT_InteractiveInfo = 0x0FF2,
    RT_InteractiveInfoAtom = 0x0FF3,
    RT_UserEditAtom = 0x0FF5,
    RT_CurrentUserAtom = 0x0FF6,
    RT_DateTimeMetaCharAtom = 0x0FF7,
    RT_GenericDateMetaCharAtom = 0x0FF8,
    RT_HeaderMetaCharAtom = 0x0FF9,
    RT_FooterMetaCharAtom = 0x0FFA,
    RT_ExternalOleControlAtom = 0x0FFB,
    RT_ExternalMediaAtom = 0x1004,
    RT_ExternalVideo = 0x1005,
    RT_ExternalAviMovie = 0x1006,
    RT_ExternalMciMovie = 0x1007,
    RT_ExternalMidiAudio = 0x100D,
    RT_ExternalCdAudio = 0x100E,
    RT_ExternalWavAudioEmbedded = 0x100F,
    RT_ExternalWavAudioLink = 0x1010,
    RT_ExternalOleObjectStg = 0x1011,
    RT_ExternalCdAudioAtom = 0x1012,
    RT_ExternalWavAudioEmbeddedAtom = 0x1013,
    RT_AnimationInfo = 0x1014,
    RT_RtfDateTimeMetaCharAtom = 0x1015,
    RT_ExternalHyperlinkFlagsAtom = 0x1018,
    RT_ProgTags = 0x1388,
    RT_ProgStringTag = 0x1389,
    RT_ProgBinaryTag = 0x138A,
    RT_BinaryTagDataBlob = 0x138B,
    RT_PrintOptionsAtom = 0x1770,
    RT_PersistDirectoryAtom = 0x1772,
    RT_PresentationAdvisorFlags9Atom = 0x177A,
    RT_HtmlDocInfo9Atom = 0x177B,
    RT_HtmlPublishInfoAtom = 0x177C,
    RT_HtmlPublishInfo9 = 0x177D,
    RT_BroadcastDocInfo9 = 0x177E,
    RT_BroadcastDocInfo9Atom = 0x177F,
    RT_EnvelopeFlags9Atom = 0x1784,
    RT_EnvelopeData9Atom = 0x1785,
    RT_VisualShapeAtom = 0x2AFB,
    RT_HashCodeAtom = 0x2B00,
    RT_VisualPageAtom = 0x2B01,
    RT_BuildList = 0x2B02,
    RT_BuildAtom = 0x2B03,
    RT_ChartBuild = 0x2B04,
    RT_ChartBuildAtom = 0x2B05,
    RT_DiagramBuild = 0x2B06,
    RT_DiagramBuildAtom = 0x2B07,
    RT_ParaBuild = 0x2B08,
    RT_ParaBuildAtom = 0x2B09,
    RT_LevelInfoAtom = 0x2B0A,
    RT_RoundTripAnimationAtom12Atom = 0x2B0B,
    RT_RoundTripAnimationHashAtom12Atom = 0x2B0D,
    RT_Comment10 = 0x2EE0,
    RT_Comment10Atom = 0x2EE1,
    RT_CommentIndex10 = 0x2EE4,
    RT_CommentIndex10Atom = 0x2EE5,
    RT_LinkedShape10Atom = 0x2EE6,
    RT_LinkedSlide10Atom = 0x2EE7,
    RT_SlideFlags10Atom = 0x2EEA,
    RT_SlideTime10Atom = 0x2EEB,
    RT_DiffTree10 = 0x2EEC,
    RT_Diff10 = 0x2EED,
    RT_Diff10Atom = 0x2EEE,
    RT_SlideListTableSize10Atom = 0x2EEF,
    RT_SlideListEntry10Atom = 0x2EF0,
    RT_SlideListTable10 = 0x2EF1,
    RT_CryptSession10Container = 0x2F14,
    RT_FontEmbedFlags10Atom = 0x32C8,
    RT_FilterPrivacyFlags10Atom = 0x36B0,
    RT_DocToolbarStates10Atom = 0x36B1,
    RT_PhotoAlbumInfo10Atom = 0x36B2,
    RT_SmartTagStore11Container = 0x36B3,
    RT_RoundTripSlideSyncInfo12 = 0x3714,
    RT_RoundTripSlideSyncInfoAtom12 = 0x3715,
    RT_TimeConditionContainer = 0xF125,
    RT_TimeNode = 0xF127,
    RT_TimeCondition = 0xF128,
    RT_TimeModifier = 0xF129,
    RT_TimeBehaviorContainer = 0xF12A,
    RT_TimeAnimateBehaviorContainer = 0xF12B,
    RT_TimeColorBehaviorContainer = 0xF12C,
    RT_TimeEffectBehaviorContainer = 0xF12D,
    RT_TimeMotionBehaviorContainer = 0xF12E,
    RT_TimeRotationBehaviorContainer = 0xF12F,
    RT_TimeScaleBehaviorContainer = 0xF130,
    RT_TimeSetBehaviorContainer = 0xF131,
    RT_TimeCommandBehaviorContainer = 0xF132,
    RT_TimeBehavior = 0xF133,
    RT_TimeAnimateBehavior = 0xF134,
    RT_TimeColorBehavior = 0xF135,
    RT_TimeEffectBehavior = 0xF136,
    RT_TimeMotionBehavior = 0xF137,
    RT_TimeRotationBehavior = 0xF138,
    RT_TimeScaleBehavior = 0xF139,
    RT_TimeSetBehavior = 0xF13A,
    RT_TimeCommandBehavior = 0xF13B,
    RT_TimeClientVisualElement = 0xF13C,
    RT_TimePropertyList = 0xF13D,
    RT_TimeVariantList = 0xF13E,
    RT_TimeAnimationValueList = 0xF13F,
    RT_TimeIterateData = 0xF140,
    RT_TimeSequenceData = 0xF141,
    RT_TimeVariant = 0xF142,
    RT_TimeAnimationValue = 0xF143,
    RT_TimeExtTimeNodeContainer = 0xF144,
    RT_TimeSubEffectContainer = 0xF145,
};

class PPT
{
private:
    DWORD offsetToCurrentEdit;
    BinaryBlock PowerPointDocument;
    BinaryBlock Pictures;

public:
    PPT(wchar_t *filePath);
    ~PPT();
    GetText(wchar_t *filePath);
};


BYTE *TextBytesToChars(BYTE *data, DWORD dataSize);

#endif // pptElementsH
