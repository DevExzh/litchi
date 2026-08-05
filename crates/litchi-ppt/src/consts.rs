//! Typed PowerPoint binary record kinds.

// PowerPoint Binary File Format (MS-PPT) constants

/// PPT record types (based on POI RecordTypes enum)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum RecordType {
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

impl From<u16> for RecordType {
    fn from(value: u16) -> Self {
        match value {
            0 => RecordType::Unknown,
            1000 => RecordType::Document,
            1001 => RecordType::DocumentAtom,
            1002 => RecordType::EndDocument,
            1006 => RecordType::Slide,
            1007 => RecordType::SlideAtom,
            1008 => RecordType::Notes,
            1009 => RecordType::NotesAtom,
            1010 => RecordType::Environment,
            2005 => RecordType::FontCollection,
            2006 => RecordType::FontCollection10,
            2040 => RecordType::BlipCollection9,
            2041 => RecordType::BlipEntity9Atom,
            1011 => RecordType::SlidePersistAtom,
            1016 => RecordType::MainMaster,
            1017 => RecordType::SSSlideInfoAtom,
            4080 => RecordType::SlideListWithText,
            6001 | 6002 => RecordType::PersistPtrHolder, // Both values are used
            12052 => RecordType::CryptSession10Container,
            1023 => RecordType::VBAInfo,
            1024 => RecordType::VBAInfoAtom,
            1033 => RecordType::ExObjList,
            1034 => RecordType::ExObjListAtom,
            1035 => RecordType::PPDrawingGroup,
            1036 => RecordType::PPDrawing,
            1037 => RecordType::GridSpacing10Atom,
            2020 => RecordType::SoundCollection,
            2021 => RecordType::SoundCollectionAtom,
            2022 => RecordType::Sound,
            2023 => RecordType::SoundData,
            0x0FC9 => RecordType::Handout,
            0x0427 => RecordType::RoundTripNotesMasterTextStyles12Atom,
            0x040E => RecordType::RoundTripTheme12Atom,
            0x0406 => RecordType::DocRoutingSlipAtom,
            0x0401 => RecordType::SlideShowDocInfoAtom,
            0x0402 => RecordType::Summary,
            0x07E3 => RecordType::BookmarkCollection,
            0x07E9 => RecordType::BookmarkSeedAtom,
            0x0FD0 => RecordType::BookmarkEntityAtom,
            0x0FA7 => RecordType::TextBookmarkAtom,
            0x1770 => RecordType::PrintOptionsAtom,
            3009 => RecordType::ExternalObjectRefAtom,
            0x0FE7 => RecordType::RecolorInfoAtom,
            0x0FC1 => RecordType::MetaFile,
            0x0FC3 => RecordType::ExternalOleObjectAtom,
            0x0FCC => RecordType::ExternalOleEmbed,
            0x0FCD => RecordType::ExternalOleEmbedAtom,
            0x0FCE => RecordType::ExternalOleLink,
            0x0FD1 => RecordType::ExternalOleLinkAtom,
            0x0FEE => RecordType::ExternalOleControl,
            0x0FFB => RecordType::ExternalOleControlAtom,
            0x1004 => RecordType::ExternalMediaAtom,
            0x1005 => RecordType::ExternalVideo,
            0x1006 => RecordType::ExternalAviMovie,
            0x1007 => RecordType::ExternalMciMovie,
            0x100D => RecordType::ExternalMidiAudio,
            0x100E => RecordType::ExternalCdAudio,
            0x100F => RecordType::ExternalWavAudioEmbedded,
            0x1010 => RecordType::ExternalWavAudioLink,
            0x1011 => RecordType::ExternalOleObjectStg,
            0x1012 => RecordType::ExternalCdAudioAtom,
            0x1013 => RecordType::ExternalWavAudioEmbeddedAtom,
            0x040F => RecordType::RoundTripColorMapping12Atom,
            0x041C => RecordType::RoundTripOriginalMainMasterId12Atom,
            0x041D => RecordType::RoundTripCompositeMasterId12Atom,
            0x041E => RecordType::RoundTripContentMasterInfo12Atom,
            0x041F => RecordType::RoundTripShapeId12Atom,
            0x0420 => RecordType::RoundTripHFPlaceholder12Atom,
            0x0422 => RecordType::RoundTripContentMasterId12Atom,
            0x0423 => RecordType::RoundTripOArtTextStyles12Atom,
            0x0424 => RecordType::RoundTripHeaderFooterDefaults12Atom,
            0x0425 => RecordType::RoundTripDocFlags12Atom,
            0x0426 => RecordType::RoundTripShapeCheckSumForCustomLayouts12Atom,
            0x0428 => RecordType::RoundTripCustomTableStyles12Atom,
            3011 => RecordType::OEPlaceholderAtom,
            0x0BDB => RecordType::ShapeAtom,
            0x0BDC => RecordType::ShapeFlags10Atom,
            0x0BDD => RecordType::RoundTripNewPlaceholderId12Atom,
            0x2B0B => RecordType::RoundTripAnimation12Atom,
            0x2B0D => RecordType::RoundTripAnimationHash12Atom,
            3999 => RecordType::TextHeaderAtom,
            4000 => RecordType::TextCharsAtom,
            4008 => RecordType::TextBytesAtom,
            4010 => RecordType::TextSpecInfoAtom,
            4011 => RecordType::DefaultRulerAtom,
            4012 => RecordType::StyleTextProp9Atom,
            4013 => RecordType::TextMasterStyle9Atom,
            4014 => RecordType::OutlineTextProps9,
            4015 => RecordType::OutlineTextPropsHeader9Atom,
            4016 => RecordType::TextDefaults9Atom,
            4017 => RecordType::StyleTextProp10Atom,
            4018 => RecordType::TextMasterStyle10Atom,
            4019 => RecordType::OutlineTextProps10,
            4020 => RecordType::TextDefaults10Atom,
            4021 => RecordType::OutlineTextProps11,
            4022 => RecordType::StyleTextProp11Atom,
            4001 => RecordType::StyleTextPropAtom,
            4002 => RecordType::MasterTextPropAtom,
            4003 => RecordType::TxMasterStyleAtom,
            4004 => RecordType::TxCFStyleAtom,
            4005 => RecordType::TxPFStyleAtom,
            4006 => RecordType::TextRulerAtom,
            4023 => RecordType::FontEntityAtom,
            4024 => RecordType::FontEmbeddedData,
            0x32C8 => RecordType::FontEmbedFlags10Atom,
            0x36B0 => RecordType::FilterPrivacyFlags10Atom,
            0x36B1 => RecordType::DocToolbarStates10Atom,
            0x36B2 => RecordType::PhotoAlbumInfo10Atom,
            0x36B3 => RecordType::SmartTagStore11,
            0x3714 => RecordType::RoundTripSlideSyncInfo12,
            0x3715 => RecordType::RoundTripSlideSyncInfoAtom12,
            4026 => RecordType::CString,
            4040 => RecordType::Kinsoku,
            4050 => RecordType::KinsokuAtom,
            4051 => RecordType::ExternalHyperlinkAtom,
            4055 => RecordType::ExternalHyperlink,
            4063 => RecordType::TextInteractiveInfoAtom,
            4068 => RecordType::ExternalHyperlink9,
            4057 => RecordType::HeadersFooters,
            4058 => RecordType::HeadersFootersAtom,
            4082 => RecordType::InteractiveInfo,
            4083 => RecordType::InteractiveInfoAtom,
            4085 => RecordType::UserEditAtom,
            4086 => RecordType::CurrentUserAtom,
            3998 => RecordType::OutlineTextRefAtom,
            4009 => RecordType::TextSpecialInfoDefaultAtom,
            1043 => RecordType::NotesTextViewInfo9,
            1044 => RecordType::NormalViewSetInfo9,
            1045 => RecordType::NormalViewSetInfo9Atom,
            4056 => RecordType::SlideNumberMCAtom,
            4087 => RecordType::DateTimeMCAtom,
            4088 => RecordType::GenericDateMCAtom,
            4089 => RecordType::HeaderMCAtom,
            4090 => RecordType::FooterMCAtom,
            4117 => RecordType::RtfDateTimeMCAtom,
            0x1388 => RecordType::ProgTags,
            0x1389 => RecordType::ProgStringTag,
            0x138A => RecordType::ProgBinaryTag,
            0x138B => RecordType::BinaryTagData,
            4116 => RecordType::AnimationInfo,
            4081 => RecordType::AnimationInfoAtom,
            4120 => RecordType::ExternalHyperlinkFlagsAtom,
            0x2B02 => RecordType::BuildList,
            0x2B03 => RecordType::BuildAtom,
            0x2B04 => RecordType::ChartBuild,
            0x2B05 => RecordType::ChartBuildAtom,
            0x2B06 => RecordType::DiagramBuild,
            0x2B07 => RecordType::DiagramBuildAtom,
            0x2B08 => RecordType::ParaBuild,
            0x2B09 => RecordType::ParaBuildAtom,
            0x2B0A => RecordType::LevelInfoAtom,
            0xF144 => RecordType::ExtTimeNode,
            0xF145 => RecordType::TimeSubEffectContainer,
            0xF125 => RecordType::TimeConditionContainer,
            0xF127 => RecordType::TimeNode,
            0xF128 => RecordType::TimeCondition,
            0xF129 => RecordType::TimeModifier,
            0xF12A => RecordType::TimeBehaviorContainer,
            0xF12B => RecordType::TimeAnimateBehaviorContainer,
            0xF12C => RecordType::TimeColorBehaviorContainer,
            0xF12D => RecordType::TimeEffectBehaviorContainer,
            0xF12E => RecordType::TimeMotionBehaviorContainer,
            0xF12F => RecordType::TimeRotationBehaviorContainer,
            0xF130 => RecordType::TimeScaleBehaviorContainer,
            0xF131 => RecordType::TimeSetBehaviorContainer,
            0xF132 => RecordType::TimeCommandBehaviorContainer,
            0xF133 => RecordType::TimeBehavior,
            0xF134 => RecordType::TimeAnimateBehavior,
            0xF135 => RecordType::TimeColorBehavior,
            0xF136 => RecordType::TimeEffectBehavior,
            0xF137 => RecordType::TimeMotionBehavior,
            0xF138 => RecordType::TimeRotationBehavior,
            0xF139 => RecordType::TimeScaleBehavior,
            0xF13A => RecordType::TimeSetBehavior,
            0xF13B => RecordType::TimeCommandBehavior,
            0xF13C => RecordType::TimeClientVisualElement,
            0xF13D => RecordType::TimePropertyList,
            0xF13E => RecordType::TimeVariantList,
            0xF13F => RecordType::TimeAnimationValueList,
            0xF140 => RecordType::TimeIterateData,
            0xF141 => RecordType::TimeSequenceData,
            0xF142 => RecordType::TimeVariant,
            0xF143 => RecordType::TimeAnimationValue,
            0x2AFB => RecordType::VisualShapeAtom,
            0x2B00 => RecordType::HashCode10Atom,
            0x2B01 => RecordType::VisualPageAtom,
            1040 => RecordType::NamedShows,
            1041 => RecordType::NamedShow,
            1042 => RecordType::NamedShowSlides,
            1018 => RecordType::SlideViewInfo,
            1019 => RecordType::GuideAtom,
            1021 => RecordType::ViewInfoAtom,
            1022 => RecordType::SlideViewInfoAtom,
            1031 => RecordType::OutlineViewInfo,
            1032 => RecordType::SorterViewInfo,
            2000 => RecordType::DocInfoList,
            12000 => RecordType::Comment2000,
            12001 => RecordType::Comment2000Atom,
            12004 => RecordType::CommentIndex10,
            12005 => RecordType::CommentIndex10Atom,
            12006 => RecordType::LinkedShape10Atom,
            12007 => RecordType::LinkedSlide10Atom,
            0x2EEC => RecordType::DiffTree10,
            0x2EED => RecordType::Diff10,
            0x2EEE => RecordType::Diff10Atom,
            0x2EEF => RecordType::SlideListTableSize10Atom,
            0x2EF0 => RecordType::SlideListEntry10Atom,
            0x2EF1 => RecordType::SlideListTable10,
            12010 => RecordType::SlideFlags10Atom,
            12011 => RecordType::SlideTime10Atom,
            _ => RecordType::Unknown,
        }
    }
}

impl RecordType {
    /// Get the u16 value of this record type
    pub fn as_u16(self) -> u16 {
        self as u16
    }
}
