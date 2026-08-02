//! Typed PowerPoint binary record kinds.

// PowerPoint Binary File Format (MS-PPT) constants

/// PPT record types (based on POI RecordTypes enum)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum PptRecordType {
    /// Unknown record type
    Unknown = 0,
    /// Document record
    Document = 1000,
    /// Document atom record
    DocumentAtom = 1001,
    /// End document record
    EndDocument = 1002,
    /// Slide record
    Slide = 1006,
    /// Slide atom record
    SlideAtom = 1007,
    /// Notes record
    Notes = 1008,
    /// Handout record
    Handout = 0x0FC9,
    /// Notes atom record
    NotesAtom = 1009,
    /// Environment record
    Environment = 1010,
    /// Font collection container
    FontCollection = 2005,
    /// PowerPoint 10 international font collection container
    FontCollection10 = 2006,
    /// PowerPoint 9 picture-bullet collection container
    BlipCollection9 = 2040,
    /// PowerPoint 9 picture-bullet atom
    BlipEntity9Atom = 2041,
    /// Slide persist atom record
    SlidePersistAtom = 1011,
    /// Main master record
    MainMaster = 1016,
    /// Slide list with text record
    SlideListWithText = 4080,
    /// Persist pointer holder record
    PersistPtrHolder = 6001,
    /// PowerPoint 10 cryptographic session container
    CryptSession10Container = 12052,
    /// Slide show slide info atom
    SSSlideInfoAtom = 1017,
    /// VBA info record
    VBAInfo = 1023,
    /// VBA info atom record
    VBAInfoAtom = 1024,
    /// External object list record
    ExObjList = 1033,
    /// External object list atom record
    ExObjListAtom = 1034,
    /// PP drawing group record
    PPDrawingGroup = 1035,
    /// PP drawing record
    PPDrawing = 1036,
    /// PowerPoint 10 square-grid spacing atom
    GridSpacing10Atom = 1037,
    /// PowerPoint 12 embedded DrawingML theme package atom
    RoundTripTheme12Atom = 0x040E,
    /// PowerPoint 12 DrawingML color mapping XML atom
    DocRoutingSlipAtom = 0x0406,
    SlideShowDocInfoAtom = 0x0401,
    Summary = 0x0402,
    BookmarkCollection = 0x07E3,
    BookmarkSeedAtom = 0x07E9,
    BookmarkEntityAtom = 0x0FD0,
    TextBookmarkAtom = 0x0FA7,
    PrintOptionsAtom = 0x1770,
    ExternalObjectRefAtom = 3009,
    /// Metafile recolor mapping atom.
    RecolorInfoAtom = 0x0FE7,
    MetaFile = 0x0FC1,
    ExternalOleObjectAtom = 0x0FC3,
    ExternalOleEmbed = 0x0FCC,
    ExternalOleEmbedAtom = 0x0FCD,
    ExternalOleLink = 0x0FCE,
    ExternalOleLinkAtom = 0x0FD1,
    ExternalOleControl = 0x0FEE,
    ExternalOleControlAtom = 0x0FFB,
    ExternalMediaAtom = 0x1004,
    ExternalVideo = 0x1005,
    ExternalAviMovie = 0x1006,
    ExternalMciMovie = 0x1007,
    ExternalMidiAudio = 0x100D,
    ExternalCdAudio = 0x100E,
    ExternalWavAudioEmbedded = 0x100F,
    ExternalWavAudioLink = 0x1010,
    ExternalOleObjectStg = 0x1011,
    ExternalCdAudioAtom = 0x1012,
    ExternalWavAudioEmbeddedAtom = 0x1013,
    RoundTripColorMapping12Atom = 0x040F,
    /// PowerPoint 12 notes-master text styles atom
    RoundTripNotesMasterTextStyles12Atom = 0x0427,
    /// PowerPoint 12 original main-master identifier atom
    RoundTripOriginalMainMasterId12Atom = 0x041C,
    /// PowerPoint 12 composite master identifier atom
    RoundTripCompositeMasterId12Atom = 0x041D,
    /// PowerPoint 12 embedded content-master slide-layout package atom
    RoundTripContentMasterInfo12Atom = 0x041E,
    /// PowerPoint 12 round-trip shape identifier atom
    RoundTripShapeId12Atom = 0x041F,
    /// PowerPoint 12 header/footer placeholder identity atom
    RoundTripHFPlaceholder12Atom = 0x0420,
    /// PowerPoint 12 content master identifier atom
    RoundTripContentMasterId12Atom = 0x0422,
    /// PowerPoint 12 embedded main-master text-styles package atom
    RoundTripOArtTextStyles12Atom = 0x0423,
    /// PowerPoint 12 default header and footer flags atom
    RoundTripHeaderFooterDefaults12Atom = 0x0424,
    /// PowerPoint 12 document round-trip flags atom
    RoundTripDocFlags12Atom = 0x0425,
    /// PowerPoint 12 custom-layout shape and text checksum atom
    RoundTripShapeCheckSumForCustomLayouts12Atom = 0x0426,
    /// PowerPoint 12 embedded custom table-styles package atom
    RoundTripCustomTableStyles12Atom = 0x0428,
    /// OE placeholder atom record (placeholder data)
    OEPlaceholderAtom = 3011,
    /// ShapeFlagsAtom shape-level flags
    ShapeAtom = 0x0BDB,
    /// ShapeFlags10Atom PowerPoint 2002 shape-level flags
    ShapeFlags10Atom = 0x0BDC,
    /// PowerPoint 12 new placeholder identity atom
    RoundTripNewPlaceholderId12Atom = 0x0BDD,
    /// PowerPoint 12 embedded animation package atom
    RoundTripAnimation12Atom = 0x2B0B,
    /// PowerPoint 12 animation checksum atom
    RoundTripAnimationHash12Atom = 0x2B0D,
    /// Text header atom record
    TextHeaderAtom = 3999,
    /// Text characters atom record
    TextCharsAtom = 4000,
    /// Text bytes atom record
    TextBytesAtom = 4008,
    /// Text special info atom record
    TextSpecInfoAtom = 4010,
    /// Default text ruler atom record
    DefaultRulerAtom = 4011,
    /// PowerPoint 9 additional text properties atom record
    StyleTextProp9Atom = 4012,
    /// PowerPoint 9 master text style atom record
    TextMasterStyle9Atom = 4013,
    /// PowerPoint 9 outline text properties container
    OutlineTextProps9 = 4014,
    /// PowerPoint 9 outline text properties header atom
    OutlineTextPropsHeader9Atom = 4015,
    /// PowerPoint 9 default text properties atom
    TextDefaults9Atom = 4016,
    /// PowerPoint 10 additional character properties atom
    StyleTextProp10Atom = 4017,
    /// PowerPoint 10 master text style atom record
    TextMasterStyle10Atom = 4018,
    /// PowerPoint 10 outline text properties container
    OutlineTextProps10 = 4019,
    /// PowerPoint 10 default text properties atom
    TextDefaults10Atom = 4020,
    /// PowerPoint 11 outline text properties container
    OutlineTextProps11 = 4021,
    /// PowerPoint 11 additional text properties atom
    StyleTextProp11Atom = 4022,
    /// Style text prop atom record
    StyleTextPropAtom = 4001,
    /// Master text prop atom record
    MasterTextPropAtom = 4002,
    /// Text master style atom record
    TxMasterStyleAtom = 4003,
    /// Text CF style atom record
    TxCFStyleAtom = 4004,
    /// Text PF style atom record
    TxPFStyleAtom = 4005,
    /// Text ruler atom record
    TextRulerAtom = 4006,
    /// Font entity atom record
    FontEntityAtom = 4023,
    /// Embedded font data atom record
    FontEmbeddedData = 4024,
    /// PowerPoint 10 font embedding flags atom
    FontEmbedFlags10Atom = 0x32C8,
    /// PowerPoint 10 privacy flags atom
    FilterPrivacyFlags10Atom = 0x36B0,
    /// PowerPoint 10 reviewing toolbar and gallery state atom
    DocToolbarStates10Atom = 0x36B1,
    /// PowerPoint 10 photo album settings atom
    PhotoAlbumInfo10Atom = 0x36B2,
    /// PowerPoint 11 smart tag store container
    SmartTagStore11 = 0x36B3,
    /// PowerPoint 12 slide-library synchronization container
    RoundTripSlideSyncInfo12 = 0x3714,
    /// PowerPoint 12 slide-library synchronization timestamps atom
    RoundTripSlideSyncInfoAtom12 = 0x3715,
    /// CString record
    CString = 4026,
    /// East Asian line-breaking settings container
    Kinsoku = 4040,
    /// East Asian line-breaking settings atom
    KinsokuAtom = 4050,
    /// External hyperlink atom or hyperlink reference atom
    ExternalHyperlinkAtom = 4051,
    /// External hyperlink container
    ExternalHyperlink = 4055,
    /// Text-range anchor for the preceding interactive information record
    TextInteractiveInfoAtom = 4063,
    /// PowerPoint 9 external hyperlink extension container
    ExternalHyperlink9 = 4068,
    /// Headers footers container record
    HeadersFooters = 4057,
    /// Headers footers atom record
    HeadersFootersAtom = 4058,
    /// Interactive info record
    InteractiveInfo = 4082,
    /// Interactive info atom record
    InteractiveInfoAtom = 4083,
    /// User edit atom record
    UserEditAtom = 4085,
    /// Current user atom record
    CurrentUserAtom = 4086,
    /// Notes text view info 9 container
    NotesTextViewInfo9 = 1043,
    /// Normal view set info 9 container
    NormalViewSetInfo9 = 1044,
    /// Normal view set info 9 atom record
    NormalViewSetInfo9Atom = 1045,
    /// Outline text reference atom record
    OutlineTextRefAtom = 3998,
    /// Text special info default atom record
    TextSpecialInfoDefaultAtom = 4009,
    /// Slide number metachar atom record
    SlideNumberMCAtom = 4056,
    /// Date time MC atom record
    DateTimeMCAtom = 4087,
    /// Generic date metachar atom record
    GenericDateMCAtom = 4088,
    /// Header metachar atom record
    HeaderMCAtom = 4089,
    /// Footer metachar atom record
    FooterMCAtom = 4090,
    /// RTF date time metachar atom record
    RtfDateTimeMCAtom = 4117,
    /// Programmable tags container
    ProgTags = 0x1388,
    /// Programmable string tag container
    ProgStringTag = 0x1389,
    /// Programmable binary tag container
    ProgBinaryTag = 0x138A,
    /// Programmable binary tag data atom
    BinaryTagData = 0x138B,
    /// Animation info record
    AnimationInfo = 4116,
    /// Animation info atom record
    AnimationInfoAtom = 4081,
    /// PowerPoint 9 external hyperlink flags atom
    ExternalHyperlinkFlagsAtom = 4120,
    /// Build list record
    BuildList = 0x2B02,
    /// Build atom record
    BuildAtom = 0x2B03,
    /// Chart build record
    ChartBuild = 0x2B04,
    /// Chart build atom record
    ChartBuildAtom = 0x2B05,
    /// Diagram build record
    DiagramBuild = 0x2B06,
    /// Diagram build atom record
    DiagramBuildAtom = 0x2B07,
    /// Paragraph build record
    ParaBuild = 0x2B08,
    /// Paragraph build atom record
    ParaBuildAtom = 0x2B09,
    /// Paragraph build level atom record
    LevelInfoAtom = 0x2B0A,
    /// Extended time node container record
    ExtTimeNode = 0xF144,
    /// Subordinate effect time node container record
    TimeSubEffectContainer = 0xF145,
    /// Sound collection container
    SoundCollection = 2020,
    /// Sound collection atom
    SoundCollectionAtom = 2021,
    /// Sound record
    Sound = 2022,
    /// Sound data record
    SoundData = 2023,
    /// Time node record
    TimeNode = 0xF127,
    /// Time condition container
    TimeConditionContainer = 0xF125,
    /// Time condition atom
    TimeCondition = 0xF128,
    /// Time modifier atom
    TimeModifier = 0xF129,
    /// Shared animation behavior container
    TimeBehaviorContainer = 0xF12A,
    /// Generic property animation behavior container
    TimeAnimateBehaviorContainer = 0xF12B,
    /// Color behavior container
    TimeColorBehaviorContainer = 0xF12C,
    /// Image effect behavior container
    TimeEffectBehaviorContainer = 0xF12D,
    /// Motion-path behavior container
    TimeMotionBehaviorContainer = 0xF12E,
    /// Rotation behavior container
    TimeRotationBehaviorContainer = 0xF12F,
    /// Scale behavior container
    TimeScaleBehaviorContainer = 0xF130,
    /// Set-property behavior container
    TimeSetBehaviorContainer = 0xF131,
    /// Command behavior container
    TimeCommandBehaviorContainer = 0xF132,
    /// Shared animation behavior atom
    TimeBehavior = 0xF133,
    /// Generic property animation behavior atom
    TimeAnimateBehavior = 0xF134,
    /// Color behavior atom
    TimeColorBehavior = 0xF135,
    /// Image effect behavior atom
    TimeEffectBehavior = 0xF136,
    /// Motion-path behavior atom
    TimeMotionBehavior = 0xF137,
    /// Rotation behavior atom
    TimeRotationBehavior = 0xF138,
    /// Scale behavior atom
    TimeScaleBehavior = 0xF139,
    /// Set-property behavior atom
    TimeSetBehavior = 0xF13A,
    /// Command behavior atom
    TimeCommandBehavior = 0xF13B,
    /// Animation target container
    TimeClientVisualElement = 0xF13C,
    /// Time property list record
    TimePropertyList = 0xF13D,
    /// Time string-list container
    TimeVariantList = 0xF13E,
    /// Generic animation keyframe list container
    TimeAnimationValueList = 0xF13F,
    /// Time iterate-data atom
    TimeIterateData = 0xF140,
    /// Time sequence-data atom
    TimeSequenceData = 0xF141,
    /// Time variant atom record
    TimeVariant = 0xF142,
    /// Generic animation keyframe time atom
    TimeAnimationValue = 0xF143,
    /// Shape or sound animation target atom
    VisualShapeAtom = 0x2AFB,
    /// PowerPoint 10 animation hash atom
    HashCode10Atom = 0x2B00,
    /// Slide animation target atom
    VisualPageAtom = 0x2B01,
    /// Named shows container (custom slide shows)
    NamedShows = 1040,
    /// Named show container
    NamedShow = 1041,
    /// Named show slides atom
    NamedShowSlides = 1042,
    /// Slide or notes view information container
    SlideViewInfo = 1018,
    /// Slide or notes alignment guide atom
    GuideAtom = 1019,
    /// Zoom view information atom
    ViewInfoAtom = 1021,
    /// Slide view editing-preferences atom
    SlideViewInfoAtom = 1022,
    /// Outline editing-view information container
    OutlineViewInfo = 1031,
    /// Slide-sorter editing-view information container
    SorterViewInfo = 1032,
    /// Document information list container
    DocInfoList = 2000,
    /// Comment 2000 record
    Comment2000 = 12000,
    /// Comment 2000 atom record
    Comment2000Atom = 12001,
    /// PowerPoint 10 comment author container
    CommentIndex10 = 12004,
    /// PowerPoint 10 comment author index atom
    CommentIndex10Atom = 12005,
    /// PowerPoint 10 linked-shape atom
    LinkedShape10Atom = 12006,
    /// PowerPoint 10 linked-slide atom
    LinkedSlide10Atom = 12007,
    /// PowerPoint 10 document-comparison tree container
    DiffTree10 = 0x2EEC,
    /// PowerPoint 10 document-comparison diff container
    Diff10 = 0x2EED,
    /// PowerPoint 10 document-comparison diff atom
    Diff10Atom = 0x2EEE,
    /// PowerPoint 10 document-comparison slide-list size atom
    SlideListTableSize10Atom = 0x2EEF,
    /// PowerPoint 10 document-comparison slide-list entry atom
    SlideListEntry10Atom = 0x2EF0,
    /// PowerPoint 10 document-comparison slide-list table container
    SlideListTable10 = 0x2EF1,
    /// PowerPoint 10 slide flags atom
    SlideFlags10Atom = 12010,
    /// PowerPoint 10 slide creation time atom
    SlideTime10Atom = 12011,
}

impl From<u16> for PptRecordType {
    fn from(value: u16) -> Self {
        match value {
            0 => PptRecordType::Unknown,
            1000 => PptRecordType::Document,
            1001 => PptRecordType::DocumentAtom,
            1002 => PptRecordType::EndDocument,
            1006 => PptRecordType::Slide,
            1007 => PptRecordType::SlideAtom,
            1008 => PptRecordType::Notes,
            1009 => PptRecordType::NotesAtom,
            1010 => PptRecordType::Environment,
            2005 => PptRecordType::FontCollection,
            2006 => PptRecordType::FontCollection10,
            2040 => PptRecordType::BlipCollection9,
            2041 => PptRecordType::BlipEntity9Atom,
            1011 => PptRecordType::SlidePersistAtom,
            1016 => PptRecordType::MainMaster,
            1017 => PptRecordType::SSSlideInfoAtom,
            4080 => PptRecordType::SlideListWithText,
            6001 | 6002 => PptRecordType::PersistPtrHolder, // Both values are used
            12052 => PptRecordType::CryptSession10Container,
            1023 => PptRecordType::VBAInfo,
            1024 => PptRecordType::VBAInfoAtom,
            1033 => PptRecordType::ExObjList,
            1034 => PptRecordType::ExObjListAtom,
            1035 => PptRecordType::PPDrawingGroup,
            1036 => PptRecordType::PPDrawing,
            1037 => PptRecordType::GridSpacing10Atom,
            2020 => PptRecordType::SoundCollection,
            2021 => PptRecordType::SoundCollectionAtom,
            2022 => PptRecordType::Sound,
            2023 => PptRecordType::SoundData,
            0x0FC9 => PptRecordType::Handout,
            0x0427 => PptRecordType::RoundTripNotesMasterTextStyles12Atom,
            0x040E => PptRecordType::RoundTripTheme12Atom,
            0x0406 => PptRecordType::DocRoutingSlipAtom,
            0x0401 => PptRecordType::SlideShowDocInfoAtom,
            0x0402 => PptRecordType::Summary,
            0x07E3 => PptRecordType::BookmarkCollection,
            0x07E9 => PptRecordType::BookmarkSeedAtom,
            0x0FD0 => PptRecordType::BookmarkEntityAtom,
            0x0FA7 => PptRecordType::TextBookmarkAtom,
            0x1770 => PptRecordType::PrintOptionsAtom,
            3009 => PptRecordType::ExternalObjectRefAtom,
            0x0FE7 => PptRecordType::RecolorInfoAtom,
            0x0FC1 => PptRecordType::MetaFile,
            0x0FC3 => PptRecordType::ExternalOleObjectAtom,
            0x0FCC => PptRecordType::ExternalOleEmbed,
            0x0FCD => PptRecordType::ExternalOleEmbedAtom,
            0x0FCE => PptRecordType::ExternalOleLink,
            0x0FD1 => PptRecordType::ExternalOleLinkAtom,
            0x0FEE => PptRecordType::ExternalOleControl,
            0x0FFB => PptRecordType::ExternalOleControlAtom,
            0x1004 => PptRecordType::ExternalMediaAtom,
            0x1005 => PptRecordType::ExternalVideo,
            0x1006 => PptRecordType::ExternalAviMovie,
            0x1007 => PptRecordType::ExternalMciMovie,
            0x100D => PptRecordType::ExternalMidiAudio,
            0x100E => PptRecordType::ExternalCdAudio,
            0x100F => PptRecordType::ExternalWavAudioEmbedded,
            0x1010 => PptRecordType::ExternalWavAudioLink,
            0x1011 => PptRecordType::ExternalOleObjectStg,
            0x1012 => PptRecordType::ExternalCdAudioAtom,
            0x1013 => PptRecordType::ExternalWavAudioEmbeddedAtom,
            0x040F => PptRecordType::RoundTripColorMapping12Atom,
            0x041C => PptRecordType::RoundTripOriginalMainMasterId12Atom,
            0x041D => PptRecordType::RoundTripCompositeMasterId12Atom,
            0x041E => PptRecordType::RoundTripContentMasterInfo12Atom,
            0x041F => PptRecordType::RoundTripShapeId12Atom,
            0x0420 => PptRecordType::RoundTripHFPlaceholder12Atom,
            0x0422 => PptRecordType::RoundTripContentMasterId12Atom,
            0x0423 => PptRecordType::RoundTripOArtTextStyles12Atom,
            0x0424 => PptRecordType::RoundTripHeaderFooterDefaults12Atom,
            0x0425 => PptRecordType::RoundTripDocFlags12Atom,
            0x0426 => PptRecordType::RoundTripShapeCheckSumForCustomLayouts12Atom,
            0x0428 => PptRecordType::RoundTripCustomTableStyles12Atom,
            3011 => PptRecordType::OEPlaceholderAtom,
            0x0BDB => PptRecordType::ShapeAtom,
            0x0BDC => PptRecordType::ShapeFlags10Atom,
            0x0BDD => PptRecordType::RoundTripNewPlaceholderId12Atom,
            0x2B0B => PptRecordType::RoundTripAnimation12Atom,
            0x2B0D => PptRecordType::RoundTripAnimationHash12Atom,
            3999 => PptRecordType::TextHeaderAtom,
            4000 => PptRecordType::TextCharsAtom,
            4008 => PptRecordType::TextBytesAtom,
            4010 => PptRecordType::TextSpecInfoAtom,
            4011 => PptRecordType::DefaultRulerAtom,
            4012 => PptRecordType::StyleTextProp9Atom,
            4013 => PptRecordType::TextMasterStyle9Atom,
            4014 => PptRecordType::OutlineTextProps9,
            4015 => PptRecordType::OutlineTextPropsHeader9Atom,
            4016 => PptRecordType::TextDefaults9Atom,
            4017 => PptRecordType::StyleTextProp10Atom,
            4018 => PptRecordType::TextMasterStyle10Atom,
            4019 => PptRecordType::OutlineTextProps10,
            4020 => PptRecordType::TextDefaults10Atom,
            4021 => PptRecordType::OutlineTextProps11,
            4022 => PptRecordType::StyleTextProp11Atom,
            4001 => PptRecordType::StyleTextPropAtom,
            4002 => PptRecordType::MasterTextPropAtom,
            4003 => PptRecordType::TxMasterStyleAtom,
            4004 => PptRecordType::TxCFStyleAtom,
            4005 => PptRecordType::TxPFStyleAtom,
            4006 => PptRecordType::TextRulerAtom,
            4023 => PptRecordType::FontEntityAtom,
            4024 => PptRecordType::FontEmbeddedData,
            0x32C8 => PptRecordType::FontEmbedFlags10Atom,
            0x36B0 => PptRecordType::FilterPrivacyFlags10Atom,
            0x36B1 => PptRecordType::DocToolbarStates10Atom,
            0x36B2 => PptRecordType::PhotoAlbumInfo10Atom,
            0x36B3 => PptRecordType::SmartTagStore11,
            0x3714 => PptRecordType::RoundTripSlideSyncInfo12,
            0x3715 => PptRecordType::RoundTripSlideSyncInfoAtom12,
            4026 => PptRecordType::CString,
            4040 => PptRecordType::Kinsoku,
            4050 => PptRecordType::KinsokuAtom,
            4051 => PptRecordType::ExternalHyperlinkAtom,
            4055 => PptRecordType::ExternalHyperlink,
            4063 => PptRecordType::TextInteractiveInfoAtom,
            4068 => PptRecordType::ExternalHyperlink9,
            4057 => PptRecordType::HeadersFooters,
            4058 => PptRecordType::HeadersFootersAtom,
            4082 => PptRecordType::InteractiveInfo,
            4083 => PptRecordType::InteractiveInfoAtom,
            4085 => PptRecordType::UserEditAtom,
            4086 => PptRecordType::CurrentUserAtom,
            3998 => PptRecordType::OutlineTextRefAtom,
            4009 => PptRecordType::TextSpecialInfoDefaultAtom,
            1043 => PptRecordType::NotesTextViewInfo9,
            1044 => PptRecordType::NormalViewSetInfo9,
            1045 => PptRecordType::NormalViewSetInfo9Atom,
            4056 => PptRecordType::SlideNumberMCAtom,
            4087 => PptRecordType::DateTimeMCAtom,
            4088 => PptRecordType::GenericDateMCAtom,
            4089 => PptRecordType::HeaderMCAtom,
            4090 => PptRecordType::FooterMCAtom,
            4117 => PptRecordType::RtfDateTimeMCAtom,
            0x1388 => PptRecordType::ProgTags,
            0x1389 => PptRecordType::ProgStringTag,
            0x138A => PptRecordType::ProgBinaryTag,
            0x138B => PptRecordType::BinaryTagData,
            4116 => PptRecordType::AnimationInfo,
            4081 => PptRecordType::AnimationInfoAtom,
            4120 => PptRecordType::ExternalHyperlinkFlagsAtom,
            0x2B02 => PptRecordType::BuildList,
            0x2B03 => PptRecordType::BuildAtom,
            0x2B04 => PptRecordType::ChartBuild,
            0x2B05 => PptRecordType::ChartBuildAtom,
            0x2B06 => PptRecordType::DiagramBuild,
            0x2B07 => PptRecordType::DiagramBuildAtom,
            0x2B08 => PptRecordType::ParaBuild,
            0x2B09 => PptRecordType::ParaBuildAtom,
            0x2B0A => PptRecordType::LevelInfoAtom,
            0xF144 => PptRecordType::ExtTimeNode,
            0xF145 => PptRecordType::TimeSubEffectContainer,
            0xF125 => PptRecordType::TimeConditionContainer,
            0xF127 => PptRecordType::TimeNode,
            0xF128 => PptRecordType::TimeCondition,
            0xF129 => PptRecordType::TimeModifier,
            0xF12A => PptRecordType::TimeBehaviorContainer,
            0xF12B => PptRecordType::TimeAnimateBehaviorContainer,
            0xF12C => PptRecordType::TimeColorBehaviorContainer,
            0xF12D => PptRecordType::TimeEffectBehaviorContainer,
            0xF12E => PptRecordType::TimeMotionBehaviorContainer,
            0xF12F => PptRecordType::TimeRotationBehaviorContainer,
            0xF130 => PptRecordType::TimeScaleBehaviorContainer,
            0xF131 => PptRecordType::TimeSetBehaviorContainer,
            0xF132 => PptRecordType::TimeCommandBehaviorContainer,
            0xF133 => PptRecordType::TimeBehavior,
            0xF134 => PptRecordType::TimeAnimateBehavior,
            0xF135 => PptRecordType::TimeColorBehavior,
            0xF136 => PptRecordType::TimeEffectBehavior,
            0xF137 => PptRecordType::TimeMotionBehavior,
            0xF138 => PptRecordType::TimeRotationBehavior,
            0xF139 => PptRecordType::TimeScaleBehavior,
            0xF13A => PptRecordType::TimeSetBehavior,
            0xF13B => PptRecordType::TimeCommandBehavior,
            0xF13C => PptRecordType::TimeClientVisualElement,
            0xF13D => PptRecordType::TimePropertyList,
            0xF13E => PptRecordType::TimeVariantList,
            0xF13F => PptRecordType::TimeAnimationValueList,
            0xF140 => PptRecordType::TimeIterateData,
            0xF141 => PptRecordType::TimeSequenceData,
            0xF142 => PptRecordType::TimeVariant,
            0xF143 => PptRecordType::TimeAnimationValue,
            0x2AFB => PptRecordType::VisualShapeAtom,
            0x2B00 => PptRecordType::HashCode10Atom,
            0x2B01 => PptRecordType::VisualPageAtom,
            1040 => PptRecordType::NamedShows,
            1041 => PptRecordType::NamedShow,
            1042 => PptRecordType::NamedShowSlides,
            1018 => PptRecordType::SlideViewInfo,
            1019 => PptRecordType::GuideAtom,
            1021 => PptRecordType::ViewInfoAtom,
            1022 => PptRecordType::SlideViewInfoAtom,
            1031 => PptRecordType::OutlineViewInfo,
            1032 => PptRecordType::SorterViewInfo,
            2000 => PptRecordType::DocInfoList,
            12000 => PptRecordType::Comment2000,
            12001 => PptRecordType::Comment2000Atom,
            12004 => PptRecordType::CommentIndex10,
            12005 => PptRecordType::CommentIndex10Atom,
            12006 => PptRecordType::LinkedShape10Atom,
            12007 => PptRecordType::LinkedSlide10Atom,
            0x2EEC => PptRecordType::DiffTree10,
            0x2EED => PptRecordType::Diff10,
            0x2EEE => PptRecordType::Diff10Atom,
            0x2EEF => PptRecordType::SlideListTableSize10Atom,
            0x2EF0 => PptRecordType::SlideListEntry10Atom,
            0x2EF1 => PptRecordType::SlideListTable10,
            12010 => PptRecordType::SlideFlags10Atom,
            12011 => PptRecordType::SlideTime10Atom,
            _ => PptRecordType::Unknown,
        }
    }
}

impl PptRecordType {
    /// Get the u16 value of this record type
    pub fn as_u16(self) -> u16 {
        self as u16
    }
}
