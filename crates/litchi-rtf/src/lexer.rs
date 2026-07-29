//! RTF lexer/tokenizer.
//!
//! This module implements a high-performance lexer that tokenizes RTF input
//! using arena allocation for temporary data structures.

use super::error::{RtfError, RtfResult};
use bumpalo::Bump;
use std::borrow::Cow;

/// Control word with optional parameter.
///
/// This enum represents all control words defined in the RTF 1.9.1 specification.
/// Some variants may not be actively used by the parser yet but are included
/// for completeness and future extensibility.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ControlWord<'a> {
    // Document structure
    Rtf(i32),
    Ansi,
    AnsiCodePage(i32),
    Mac,
    Pc,
    Pca,

    // Header groups
    FontTable,
    ColorTable,
    StyleSheet,
    Info,
    DefaultCharacterProperties(Option<i32>),
    DefaultParagraphProperties(Option<i32>),
    DefaultFont(Option<i32>),
    AssociatedDefaultFont(Option<i32>),
    StylesheetDefaultBidiFont(Option<i32>),
    StylesheetDefaultDoubleByteFont(Option<i32>),
    StylesheetDefaultHighAnsiFont(Option<i32>),
    StylesheetDefaultLowAnsiFont(Option<i32>),
    LowAnsiCharacter(Option<i32>),
    HighAnsiCharacter(Option<i32>),
    DoubleByteCharacter(Option<i32>),
    FontComplexScript(Option<i32>),
    CharacterGrid(Option<i32>),
    AnimatedText(Option<i32>),
    EmphasisMark(crate::EmphasisMark, Option<i32>),
    NoWidowControl(Option<i32>),
    LegacySectionNumberingLevel(i32),
    LegacyParagraphNumbering(Option<i32>),
    LegacyNumberingLevel(Option<i32>),
    LegacyNumberingLevelBullet(Option<i32>),
    LegacyNumberingLevelBody(Option<i32>),
    LegacyNumberingLevelContinue(Option<i32>),
    LegacyNumberingDecimal(Option<i32>),
    LegacyNumberingUpperRoman(Option<i32>),
    LegacyNumberingLowerRoman(Option<i32>),
    LegacyNumberingUpperLetter(Option<i32>),
    LegacyNumberingLowerLetter(Option<i32>),
    LegacyNumberingFormat(crate::LegacyParagraphNumberingFormat, Option<i32>),
    LegacyNumberingStart(Option<i32>),
    LegacyNumberingIndent(Option<i32>),
    LegacyNumberingSpace(Option<i32>),
    LegacyNumberingHanging(Option<i32>),
    LegacyNumberingPrevious(Option<i32>),
    LegacyNumberingAcross(Option<i32>),
    LegacyNumberingOnce(Option<i32>),
    LegacyNumberingRestart(Option<i32>),
    LegacyNumberingBidiA(Option<i32>),
    LegacyNumberingBidiB(Option<i32>),
    LegacyNumberingAlignLeft(Option<i32>),
    LegacyNumberingAlignCenter(Option<i32>),
    LegacyNumberingAlignRight(Option<i32>),
    LegacyNumberingFont(Option<i32>),
    LegacyNumberingFontSize(Option<i32>),
    LegacyNumberingColor(Option<i32>),
    LegacyNumberingBold(Option<i32>),
    LegacyNumberingItalic(Option<i32>),
    LegacyNumberingCaps(Option<i32>),
    LegacyNumberingSmallCaps(Option<i32>),
    LegacyNumberingStrike(Option<i32>),
    LegacyNumberingUnderlineToggle(Option<i32>),
    LegacyNumberingUnderline(crate::LegacyParagraphNumberingUnderline, Option<i32>),
    LegacyNumberingRevisionAuthor(Option<i32>),
    LegacyNumberingRevisionDate(Option<i32>),
    LegacyNumberingRevisionFormat(Option<i32>),
    LegacyNumberingRevisionNoTrack(Option<i32>),
    LegacyNumberingRevisionParagraph(Option<i32>),
    LegacyNumberingRevisionRgb(Option<i32>),
    LegacyNumberingRevisionStart(Option<i32>),
    LegacyNumberingRevisionStop(Option<i32>),
    LegacyNumberingRevisionTextStart(Option<i32>),
    LegacyNumberingTextBefore,
    LegacyNumberingTextAfter,
    GeneratedListText,
    LegacyGeneratedListText,
    ParagraphGroupTable,
    ParagraphGroup,
    ParagraphGroupParent(i32),
    TableNestingLevel(Option<i32>),

    // Stylesheet entries and metadata
    ParagraphStyle(Option<i32>),
    CharacterStyle(Option<i32>),
    SectionStyle(Option<i32>),
    TableStyle(Option<i32>),
    StyleBasedOn(i32),
    StyleNext(i32),
    StyleLink(i32),
    StyleAdditive(bool),
    StyleAutoUpdate(bool),
    StyleHidden(bool),
    StyleLocked(bool),
    StyleSemiHidden(bool),
    StyleUnhideWhenUsed(bool),
    StyleQuickFormat(bool),
    StylePriority(i32),
    StyleRevisionId(i32),
    StylePersonal(bool),
    StyleCompose(bool),
    StyleReply(bool),

    // Embedded content
    Picture,
    PictureProperties(Option<i32>),
    PictureShapeId(Option<i32>),
    InvalidPictureShapePropertyParameter,
    ShapeBinaryValue(Option<i32>),
    ShapeThemeValue(Option<i32>),
    ShapeThemeColor(crate::ShapeThemeColor, Option<i32>),
    ShapeThemeTint(Option<i32>),
    ShapeThemeShade(Option<i32>),
    ShapeResult(Option<i32>),
    Object,
    InvalidObjectDestinationParameter,
    Result,
    InvalidObjectResultDestinationParameter,
    ObjectClass,
    DocumentVariable,
    UserProperties,
    PropertyName,
    PropertyType(Option<i32>),
    StaticValue,
    LinkValue,
    ObjectName,
    ObjectData,
    ObjectEmbedded,
    ObjectLink,
    ObjectAutoLink,
    ObjectHtml,
    ObjectSubscriber(Option<i32>),
    ObjectPublisher(Option<i32>),
    ObjectInstallableCommand(Option<i32>),
    ObjectOleControl(Option<i32>),
    ObjectLinkSelf(Option<i32>),
    ObjectWidth(i32),
    ObjectHeight(i32),
    ObjectAlignment(Option<i32>),
    ObjectTranslationY(Option<i32>),
    ObjectCropTop(Option<i32>),
    ObjectCropBottom(Option<i32>),
    ObjectCropLeft(Option<i32>),
    ObjectCropRight(Option<i32>),
    ObjectScaleX(Option<i32>),
    ObjectScaleY(Option<i32>),
    ObjectLocked(bool),
    ObjectUpdate(bool),
    ObjectSetSize(bool),
    ObjectResultMerge(Option<i32>),
    ObjectResultRtf(Option<i32>),
    ObjectResultText(Option<i32>),
    ObjectResultPicture(Option<i32>),
    ObjectResultBitmap(Option<i32>),
    ObjectResultHtml(Option<i32>),
    OleClassId(Option<i32>),
    InvalidObjectModifierParameter,

    // Picture properties
    PictureWidth(i32),
    PictureHeight(i32),
    PictureGoalWidth(i32),
    PictureGoalHeight(i32),
    PictureScaleX(i32),
    PictureScaleY(i32),
    PictureScaled(Option<i32>),
    PictureBitmap(Option<i32>),
    PictureBitsPerPixel(Option<i32>),
    PictureCropLeft(Option<i32>),
    PictureCropRight(Option<i32>),
    PictureCropTop(Option<i32>),
    PictureCropBottom(Option<i32>),
    WindowsBitmapBitsPerPixel(Option<i32>),
    WindowsBitmapPlanes(Option<i32>),
    WindowsBitmapWidthBytes(Option<i32>),
    Emfblip,
    Pngblip,
    Jpegblip,
    Macpict,
    Pmmetafile(i32),
    Wmetafile(i32),
    Dibitmap(i32),
    Wbitmap(i32),
    BlipTag(i32),
    BlipUid,
    BlipUnitsPerInch(i32),

    // Field support
    Field,
    FieldInstruction,
    FieldResult,
    FieldLock(Option<i32>),
    FieldDirty(Option<i32>),
    FieldEdit(Option<i32>),
    FieldPrivate(Option<i32>),
    FormField,
    DataField,
    FormFieldType(i32),
    FormFieldMaxLength(i32),
    FormFieldProtected(bool),
    FormFieldRecalculate(bool),
    FormFieldAutomaticSize(bool),
    FormFieldName,
    FormFieldFormat,
    FormFieldDefaultText,
    FormFieldDefaultResult(i32),
    FormFieldResult(i32),
    FormFieldHalfPointSize(i32),
    FormFieldOwnHelp(bool),
    FormFieldOwnStatus(bool),
    FormFieldHelpText,
    FormFieldStatusText,
    FormFieldEntryMacro,
    FormFieldExitMacro,
    FormFieldListEntry,
    FormFieldTextType(i32),
    FormFieldHasListBox(bool),
    Generator,
    RevisionSaveTable,
    RevisionSaveId(i32),
    RevisionSaveRoot(i32),
    InsertRsid(i32),
    DeleteRsid(i32),
    CharStyleRsid(i32),
    ParagraphRsid(i32),
    SectionRsid(i32),
    TableRsid(i32),
    XmlNamespaceTable,
    XmlNamespace(i32),
    XmlOpen,
    XmlClose,
    XmlAttributeName,
    XmlAttributeValue,
    MathZoneInline,
    MathZoneDisplay,
    MathZoneParagraphProperties,
    MathZonePictureFallback,
    MathAccent,
    MathBar,
    MathBorderBox,
    MathBox,
    MathDelimiter,
    MathEquationArray,
    MathFraction,
    MathFunction,
    MathGroupChar,
    MathLimitLower,
    MathLimitUpper,
    MathMatrix,
    MathNary,
    MathPhantom,
    MathRadical,
    MathScriptPre,
    MathScriptSub,
    MathScriptSubSup,
    MathScriptSup,
    MathRun,
    MathElement,
    MathNumerator,
    MathDenominator,
    MathDegree,
    MathSubscript,
    MathSuperscript,
    MathLimit,
    MathFunctionName,
    MathMatrixRow,
    MathAccentProperties,
    MathBarProperties,
    MathBorderBoxProperties,
    MathBoxProperties,
    MathDelimiterProperties,
    MathEquationArrayProperties,
    MathFractionProperties,
    MathFunctionProperties,
    MathGroupCharProperties,
    MathLimitLowerProperties,
    MathLimitUpperProperties,
    MathMatrixProperties,
    MathNaryProperties,
    MathPhantomProperties,
    MathRadicalProperties,
    MathScriptPreProperties,
    MathScriptSubProperties,
    MathScriptSubSupProperties,
    MathScriptSupProperties,
    MathRunProperties,
    MathControlProperties,
    MathMatrixColumns,
    MathMatrixColumn,
    MathMatrixColumnProperties,
    MathArgumentProperties,
    MathPropertyArgumentSize(Option<i32>),
    MathPropertyType(Option<i32>),
    MathPropertyGrow(Option<i32>),
    MathPropertyChar(Option<i32>),
    MathPropertyBeginChar(Option<i32>),
    MathPropertyEndChar(Option<i32>),
    MathPropertySeparatorChar(Option<i32>),
    MathPropertyPosition(Option<i32>),
    MathPropertyVerticalJustify(Option<i32>),
    MathPropertyBaseJustify(Option<i32>),
    MathPropertyJustify(Option<i32>),
    MathPropertyAlign(Option<i32>),
    MathPropertyAlignScript(Option<i32>),
    MathPropertyDegreeHide(Option<i32>),
    MathPropertyDifferential(Option<i32>),
    MathPropertyDifferentialStyle(Option<i32>),
    MathPropertyHideBottom(Option<i32>),
    MathPropertyHideLeft(Option<i32>),
    MathPropertyHideRight(Option<i32>),
    MathPropertyHideTop(Option<i32>),
    MathPropertyLimitLocation(Option<i32>),
    MathPropertyPlaceholderHide(Option<i32>),
    MathPropertySubscriptHide(Option<i32>),
    MathPropertySuperscriptHide(Option<i32>),
    MathPropertyStrikeBltr(Option<i32>),
    MathPropertyStrikeHorizontal(Option<i32>),
    MathPropertyStrikeTlbr(Option<i32>),
    MathPropertyStrikeVertical(Option<i32>),
    MathPropertyStyle(Option<i32>),
    MathPropertyScript(Option<i32>),
    MathPropertyTransparent(Option<i32>),
    MathPropertyShow(Option<i32>),
    MathPropertyShape(Option<i32>),
    MathPropertyZeroAscent(Option<i32>),
    MathPropertyZeroDescent(Option<i32>),
    MathPropertyZeroWidth(Option<i32>),
    MathPropertyOperatorEmulator(Option<i32>),
    MathPropertyNoBreak(Option<i32>),
    MathPropertyNormalText(Option<i32>),
    MathPropertyLiteral(Option<i32>),
    MathPropertyMatrixColumnGap(Option<i32>),
    MathPropertyMatrixColumnGapRule(Option<i32>),
    MathPropertyMatrixColumnSpacing(Option<i32>),
    MathPropertyMatrixCellCount(Option<i32>),
    MathPropertyMatrixCellJustify(Option<i32>),
    MathPropertyRowSpacing(Option<i32>),
    MathPropertyRowSpacingRule(Option<i32>),
    MathPropertyBreak(Option<i32>),
    ThemeData,
    ColorSchemeMapping,
    LatentStyles,
    LatentStyleMax(i32),
    LatentStyleLockedDefault(i32),
    LatentStyleSemiHiddenDefault(i32),
    LatentStyleUnhideUsedDefault(i32),
    LatentStyleQuickFormatDefault(i32),
    LatentStylePriorityDefault(i32),
    LatentStyleExceptions,
    LatentStyleLocked(i32),
    LatentStyleSemiHidden(i32),
    LatentStyleUnhideUsed(i32),
    LatentStyleQuickFormat(i32),
    LatentStylePriority(i32),
    DataStore,
    MailMerge,
    MailMergeConnectString,
    MailMergeConnectStringData,
    MailMergeDataSource,
    MailMergeHeaderSource,
    MailMergeLinkToQuery(bool),
    MailMergeQuery,
    MailMergeDataSourceObject,
    MailMergeActiveRecord(i32),
    MailMergeColumnDelimiter(i32),
    MailMergeColumnCount(i32),
    MailMergeDynamicAddress(bool),
    MailMergeFirstRowHeader(bool),
    MailMergeFilter,
    MailMergeFieldMapData,
    MailMergeFieldMapColumn(i32),
    MailMergeHash(i32),
    MailMergeId(i32),
    MailMergeMappedName,
    MailMergeName,
    MailMergeRecipientData,
    MailMergeSort,
    MailMergeSourceType(i32),
    MailMergeTable,
    MailMergeUdl,
    MailMergeUdlData,
    MailMergeUniqueTag,
    MathProperties,
    MathBreakBinary(i32),
    MathBreakBinarySubtraction(i32),
    MathDefaultJustification(i32),
    MathDisplayDefaults(i32),
    MathInterEquationSpacing(i32),
    MathIntegralLimitPlacement(i32),
    MathIntraEquationSpacing(i32),
    MathLeftMargin(i32),
    MathFont(i32),
    MathNaryLimitPlacement(i32),
    MathPostSpacing(i32),
    MathPreSpacing(i32),
    MathRightMargin(i32),
    MathSmallFractions(i32),
    MathWrapIndent(i32),
    MathWrapRight(i32),
    DefaultTabWidth(Option<i32>),
    DefaultLanguage(i32),
    DefaultLanguageEastAsian(i32),
    DefaultLanguageComplexScript(i32),
    Language(i32),
    LanguageEastAsian(i32),
    LanguageNoProof(i32),
    LanguageEastAsianNoProof(i32),
    NoProof(bool),
    LeftToRightCharacter,
    RightToLeftCharacter,
    LeftToRightParagraph,
    RightToLeftParagraph,
    LeftToRightDocument,
    RightToLeftDocument,
    LeftToRightSection,
    RightToLeftSection,
    LeftToRightRow(Option<i32>),
    RightToLeftRow(Option<i32>),
    TableRightToLeft(bool),
    RightGutter(bool),
    FormProtection(Option<i32>),
    AnnotationProtection(Option<i32>),
    RevisionProtection(Option<i32>),
    ReadOnlyProtection(Option<i32>),
    AllProtection(Option<i32>),
    EnforceProtection(Option<i32>),
    ProtectionLevel(Option<i32>),
    Password,
    ProtectionUserTable,
    HyphenateAutomatically(Option<i32>),
    HyphenateCapitalizedWords(Option<i32>),
    HyphenationConsecutiveLines(Option<i32>),
    HyphenationHotZone(Option<i32>),
    NextFile,
    DocumentTemplate,
    DocumentViewKind(Option<i32>),
    DocumentViewScale(Option<i32>),
    DocumentZoomKind(Option<i32>),
    DocumentViewBackgroundShapes(Option<i32>),
    DocumentViewNoPageBoundaries(Option<i32>),
    HideReviewMarkup(Option<i32>),
    HideReviewComments(Option<i32>),
    HideReviewInsertionsAndDeletions(Option<i32>),
    WindowCaption,
    XslTransform,
    UseXslTransform(Option<i32>),
    StyleListFilter(Option<i32>),
    StyleSortMethod(Option<i32>),
    ReadOnlyRecommended(Option<i32>),
    SavePreviousPicture(Option<i32>),
    WriteReservation(Option<i32>),
    WriteReservationHash(Option<i32>),
    FromText(Option<i32>),
    FromHtml(Option<i32>),
    DocumentType(Option<i32>),
    MakeBackup(Option<i32>),
    DefaultSaveFormat(Option<i32>),
    BoilerplateDocument(Option<i32>),
    Word97CompatibilityMode(Option<i32>),
    PostScriptOverText(Option<i32>),
    HorizontalDocument(Option<i32>),
    VerticalDocument(Option<i32>),
    CompressJustification(Option<i32>),
    ExpandJustification(Option<i32>),
    LineBasedOnGrid(Option<i32>),
    FractionalCharacterWidths(Option<i32>),
    AbstractNumberingCleanup(Option<i32>),
    DocumentEventMask(Option<i32>),
    DrawingGridFollowsMargins(Option<i32>),
    SnapToDrawingGrid(Option<i32>),
    DrawingGridHorizontalSpacing(Option<i32>),
    DrawingGridVerticalSpacing(Option<i32>),
    DrawingGridHorizontalOrigin(Option<i32>),
    DrawingGridVerticalOrigin(Option<i32>),
    DrawingGridHorizontalShow(Option<i32>),
    DrawingGridVerticalShow(Option<i32>),
    ParallelGutter(Option<i32>),
    PrintTwoOnOne(Option<i32>),
    ThemeLanguage(Option<i32>),
    ThemeLanguageEastAsian(Option<i32>),
    ThemeLanguageComplexScript(Option<i32>),
    RelyOnVml(Option<i32>),
    ValidateXml(Option<i32>),
    ShowPlaceholderText(Option<i32>),
    IgnoreMixedContent(Option<i32>),
    SaveInvalidXml(Option<i32>),
    ShowXmlErrors(Option<i32>),
    DoNotEmbedSystemFonts(Option<i32>),
    DoNotEmbedLinguisticData(Option<i32>),
    TrackMoves(Option<i32>),
    TrackFormatting(Option<i32>),
    LockDocumentTheme(Option<i32>),
    LockQuickFormatSet(Option<i32>),
    UseNormalStyleForLists(Option<i32>),
    UpdateStylesFromTemplate(Option<i32>),
    DeclareStyleRestrictions(Option<i32>),
    EnforceStyleRestrictions(Option<i32>),
    StyleRestrictionsBackwardCompatibility(Option<i32>),
    AllowAutoFormatOverride(Option<i32>),
    BookFold(Option<i32>),
    ReverseBookFold(Option<i32>),
    BookFoldSheets(Option<i32>),
    RemovePersonalInformation(Option<i32>),
    RemoveDateTimeInformation(Option<i32>),
    SuppressRaisedLoweredExtraSpacing(Option<i32>),
    SuppressTopPageExtraSpacing(Option<i32>),
    SuppressSpaceBeforeAfterHardBreak(Option<i32>),
    SuppressWordPerfectExtraLineSpacing(Option<i32>),
    SuppressBottomPageExtraSpacing(Option<i32>),
    DoNotBalanceSbcsDbcs(Option<i32>),
    ExpandSpacingAtShiftReturn(Option<i32>),
    DoNotAddSpaceForUnderline(Option<i32>),
    DoNotUnderlineTrailingSpaces(Option<i32>),
    DoNotTranslateBackslashToYen(Option<i32>),
    LegacyAsianLineBreakingRules(Option<i32>),
    CombineLegacyTableBorders(Option<i32>),
    DoNotAlignTableRowsIndependently(Option<i32>),
    DoNotUseRawTableWidth(Option<i32>),
    KeepTableRowsTogether(Option<i32>),
    DoNotAdjustTableLineHeight(Option<i32>),
    DoNotBreakWrappedTablesAcrossPages(Option<i32>),
    PreventAutofitGrowthIntoMargins(Option<i32>),
    UseWord2003TableStyleRules(Option<i32>),
    DoNotUseWord97ShapeLayout(Option<i32>),
    UseLegacyFootnoteLayout(Option<i32>),
    UseHtmlParagraphAutoSpacing(Option<i32>),
    PreserveLastTabAlignment(Option<i32>),
    UseWord95AutoSpacing(Option<i32>),
    ApplyThaiLineBreakingRules(Option<i32>),
    SnapTextToGridInsideTable(Option<i32>),
    AllowHangingPunctuation(Option<i32>),
    UseAsianLineBreakingRules(Option<i32>),
    CompressPunctuationAtLineStart(Option<i32>),
    NoCompatibilityOptions(Option<i32>),
    NoUiCompatibility(Option<i32>),
    NoFeatureThrottle(Option<i32>),
    ForceCompatibilityUpgrade(Option<i32>),
    PreserveAutofitTableWidthAroundShapes(Option<i32>),
    UseHangingIndentAsNumberingTab(Option<i32>),
    UseLegacyKinsokuCharacters(Option<i32>),
    UseLegacyFloatingObjectIndentation(Option<i32>),
    AllowContextualSpacingInTables(Option<i32>),
    IgnoreCellVerticalAlignmentWithFloatingObjects(Option<i32>),
    IgnoreTextBoxVerticalAlignment(Option<i32>),
    SplitPageBreakParagraph(Option<i32>),
    UseFixedWidthHangul(Option<i32>),
    UseLegacyAutofitWidthExpansion(Option<i32>),
    UseCachedColumnBalancing(Option<i32>),
    UnderlineNumberingSuffix(Option<i32>),
    DoNotSplitRowsAroundFloatingTables(Option<i32>),
    UseAnsiKerningPairs(Option<i32>),

    // Index and table-of-contents source marks
    IndexEntry,
    IndexIdentifier(Option<i32>),
    IndexBold(bool),
    IndexItalic(bool),
    IndexReplacementText,
    IndexBookmarkRange,
    IndexYomi,
    IndexPronunciation,
    TableOfContentsEntry,
    TableOfContentsEntryNoPage,
    TableOfContentsTable(i32),
    TableOfContentsLevel(i32),

    // Font properties
    FontNumber(i32),
    FontSize(i32),
    FontCharset(i32),
    FontFamily(&'a str),
    FontTheme(&'a str),
    FontPitch(i32),
    FontCodePage(i32),
    AssociatedFontNumber(Option<i32>),
    AssociatedFontSize(Option<i32>),
    AssociatedLanguage(Option<i32>),
    AssociatedBold(Option<i32>),
    AssociatedAllCaps(Option<i32>),
    AssociatedColor(Option<i32>),
    AssociatedBaselineDown(Option<i32>),
    AssociatedExpansion(Option<i32>),
    AssociatedItalic(Option<i32>),
    AssociatedOutline(Option<i32>),
    AssociatedSmallCaps(Option<i32>),
    AssociatedShadow(Option<i32>),
    AssociatedStrike(Option<i32>),
    AssociatedUnderline(Option<i32>),
    AssociatedUnderlineDotted(Option<i32>),
    AssociatedUnderlineDouble(Option<i32>),
    AssociatedUnderlineNone(Option<i32>),
    AssociatedUnderlineWords(Option<i32>),
    AssociatedBaselineUp(Option<i32>),
    FontAlternateName,
    FontNonTaggedName,
    FontPanose,
    FontEmbedded,
    FontEmbeddedType(&'a str),
    FontFile,
    FileTable,
    FileEntry,
    FileId(i32),
    FileRelative(i32),
    FileOperatingSystem(i32),
    FileValidMac,
    FileValidDos,
    FileValidNtfs,
    FileValidHpfs,
    FileNetwork,
    FileNonFileSystem,
    UnicodeAlternate,
    UnicodeAlternateDestination,

    // Colors
    Red(i32),
    Green(i32),
    Blue(i32),
    ColorForeground(i32),
    ColorBackground(Option<i32>),

    // Character formatting
    Bold(bool),
    Italic(bool),
    Underline(bool),
    UnderlineNone,
    UnderlineDouble,
    UnderlineDotted,
    UnderlineDashed,
    UnderlineDashDot,
    UnderlineDashDotDot,
    UnderlineWords,
    UnderlineThick,
    UnderlineWave,
    UnderlineHairline,
    UnderlineThickDotted,
    UnderlineThickDashed,
    UnderlineThickDashDot,
    UnderlineThickDashDotDot,
    UnderlineThickLongDash,
    UnderlineLongDash,
    UnderlineHeavyWave,
    UnderlineDoubleWave,
    UnderlineColor(i32),
    Strike(bool),
    DoubleStrike(bool),
    Superscript(bool),
    Subscript(bool),
    NoSuperSub,
    BaselineUp(i32),
    BaselineDown(i32),
    SmallCaps(bool),
    AllCaps(bool),
    Hidden(bool),
    Outline(bool),
    Shadow(bool),
    Emboss(bool),
    Imprint(bool),
    CharSpacing(i32),
    CharSpacingTwips(i32),
    CharScale(i32),
    Kerning(i32),
    Highlight(i32),
    CharacterBorder(Option<i32>),
    CharacterShading(Option<i32>),
    CharacterForegroundPattern(Option<i32>),
    CharacterBackgroundPattern(Option<i32>),
    Plain,

    // Paragraph formatting
    Par,
    Page(Option<i32>),
    Column(Option<i32>),
    Pard,
    LeftAlign,
    RightAlign,
    Center,
    Justify,

    // Paragraph spacing and indentation
    SpaceBefore(i32),
    SpaceAfter(i32),
    SpaceBetween(i32),
    LineMultiple(bool),
    SpaceBeforeAuto(Option<i32>),
    SpaceAfterAuto(Option<i32>),
    ListSpaceBefore(Option<i32>),
    ListSpaceAfter(Option<i32>),
    NoSnapLineGrid(Option<i32>),
    ContextualSpacing(Option<i32>),
    LeftIndent(i32),
    RightIndent(i32),
    FirstLineIndent(i32),
    LogicalLeftIndent(Option<i32>),
    LogicalRightIndent(Option<i32>),
    CharacterFirstLineIndent(Option<i32>),
    CharacterLeftIndent(Option<i32>),
    CharacterRightIndent(Option<i32>),
    MirrorIndents(Option<i32>),

    // Paragraph additional properties
    KeepTogether,
    KeepNext,
    SideBySide(bool),
    PageBreakBefore,
    WidowControl,
    DropCapLines(Option<i32>),
    DropCapType(Option<i32>),
    ParagraphHyphenation(Option<i32>),
    AutoSpaceAlphabetic(Option<i32>),
    AutoSpaceNumbers(Option<i32>),
    AdjustRightIndent(Option<i32>),
    WrapDefault(Option<i32>),
    NoCharacterWrap(Option<i32>),
    NoWordWrap(Option<i32>),
    NoOverflow(Option<i32>),
    FontAlignAuto(Option<i32>),
    FontAlignHanging(Option<i32>),
    FontAlignCenter(Option<i32>),
    FontAlignRoman(Option<i32>),
    FontAlignVariable(Option<i32>),
    FontAlignFixed(Option<i32>),

    // Tables
    TableRowDefaults,
    TableRowIndex(Option<i32>),
    TableRowBandIndex(Option<i32>),
    TableLastRow(Option<i32>),
    TableAutoformatFlag(crate::TableAutoformatFlag, Option<i32>),
    TableRow,
    TableCell,
    CellX(i32),
    InTable,
    TableRowGap(Option<i32>),
    TableRowLeft(Option<i32>),
    TableRowHeight(Option<i32>),
    TablePreferredWidthUnit(crate::TableDistanceScope, Option<i32>),
    TablePreferredWidthValue(crate::TableDistanceScope, Option<i32>),
    /// `false` is logical leading (`B`); `true` is logical trailing (`A`).
    TableInvisibleWidthUnit(bool, Option<i32>),
    TableInvisibleWidthValue(bool, Option<i32>),
    TableAutoFit(Option<i32>),
    TableIndentValue(Option<i32>),
    TableIndentUnit(Option<i32>),
    TableDistanceValue(crate::TableDistanceTarget, Option<i32>),
    TableDistanceUnit(crate::TableDistanceTarget, Option<i32>),
    TableHorizontalReference(crate::TableHorizontalReference, Option<i32>),
    TableVerticalReference(crate::TableVerticalReference, Option<i32>),
    TableHorizontalPosition(crate::TableHorizontalPosition, Option<i32>),
    TableVerticalPosition(crate::TableVerticalPosition, Option<i32>),
    TableHorizontalOffset(bool, Option<i32>),
    TableVerticalOffset(bool, Option<i32>),
    TableWrapDistance(crate::TableEdge, Option<i32>),
    TableNoOverlap(Option<i32>),
    TableRowHeader(Option<i32>),
    TableRowKeep(Option<i32>),
    TableRowKeepFollow(Option<i32>),
    TableRowAlignment(crate::TableRowAlignment, Option<i32>),
    TableCellVerticalAlignment(crate::TableCellVerticalAlignment, Option<i32>),
    TableCellTextFlow(crate::TableCellTextFlow, Option<i32>),
    TableCellFitText(Option<i32>),
    TableCellNoWrap(Option<i32>),
    TableCellHideMark(Option<i32>),
    TableCellMerge(
        crate::TableCellMergeAxis,
        crate::TableCellMergeRole,
        Option<i32>,
    ),
    TableBorder(crate::table::TableBorderTarget, Option<i32>),
    TableDefaultDistanceValue(crate::TableDistanceKind, crate::TableEdge, Option<i32>),
    TableDefaultDistanceUnit(crate::TableDistanceKind, crate::TableEdge, Option<i32>),
    TableDefaultCellWidthUnit(Option<i32>),
    TableDefaultCellWidthValue(Option<i32>),
    TableShadingAmount(crate::TableDistanceScope, Option<i32>),
    TableShadingForeground(crate::TableDistanceScope, Option<i32>),
    TableShadingBackground(crate::TableDistanceScope, Option<i32>),
    TableShadingPattern(
        crate::TableDistanceScope,
        crate::ShadingPattern,
        Option<i32>,
    ),
    TableRowShadingPatternIndex(Option<i32>),
    NestedTableCell(Option<i32>),
    NestedTableRow(Option<i32>),
    NestedTableProperties(Option<i32>),
    NoNestedTables(Option<i32>),

    // Borders
    BorderTop,
    BorderBottom,
    BorderLeft,
    BorderRight,
    BorderSingle,
    BorderDotted,
    BorderDashed,
    BorderDouble,
    BorderTriple,
    BorderWave,
    BorderNone,
    BorderThick,
    BorderDashSmall,
    BorderDotDash,
    BorderDotDotDash,
    BorderThinThickSmall,
    BorderThickThinSmall,
    BorderThinThickThinSmall,
    BorderThinThickMedium,
    BorderThickThinMedium,
    BorderThinThickThinMedium,
    BorderThinThickLarge,
    BorderThickThinLarge,
    BorderThinThickThinLarge,
    BorderWavyDouble,
    BorderStriped,
    BorderEmbossed,
    BorderEngraved,
    BorderOutset,
    BorderInset,
    BorderShadow,
    BorderFrame,
    BorderWidth(Option<i32>),
    BorderColor(Option<i32>),
    BorderSpace(Option<i32>),
    PageBorderTop,
    PageBorderLeft,
    PageBorderBottom,
    PageBorderRight,
    PageBorderOptions(Option<i32>),
    PageBorderSurroundHeader,
    PageBorderSurroundFooter,
    PageBorderSnap,
    PageBorderArt(Option<i32>),

    // Shading
    Shading(Option<i32>),
    ForegroundPattern(Option<i32>),
    BackgroundPattern(Option<i32>),

    // Tab stops
    TabLeft(Option<i32>),
    TabRight(Option<i32>),
    TabCenter(Option<i32>),
    TabDecimal(Option<i32>),
    TabBar(Option<i32>),
    TabPosition(Option<i32>),
    TabLeaderDot(Option<i32>),
    TabLeaderMiddleDot(Option<i32>),
    TabLeaderHyphen(Option<i32>),
    TabLeaderUnderscore(Option<i32>),
    TabLeaderThick(Option<i32>),
    TabLeaderEqual(Option<i32>),

    // Lists
    ListTable,
    List,
    ListTemplateId(i32),
    ListSimple(bool),
    ListHybrid(bool),
    ListName,
    ListStyleName,
    ListPicture(Option<i32>),
    ShapePicture(Option<i32>),
    NonShapePicture(Option<i32>),
    ListId(i32),
    ListOverrideTable,
    ListOverride,
    ListOverrideCount(i32),
    ListOverrideIndex(i32),
    ListLevelIndex(i32),
    ListOverrideLevel,
    ListOverrideStartAt(bool),
    ListOverrideFormat(bool),
    ListLevel,
    ListLevelType(i32),
    ListLevelJustification(i32),
    ListLevelFollow(i32),
    ListLevelSpace(i32),
    ListLevelIndent(i32),
    ListNumberText,
    ListLevelStartAt(i32),
    ListLevelNumbers,
    ListLevelPicture(i32),
    ListLevelTentative,
    ListLevelLegal(bool),
    ListLevelNoRestart(bool),
    ListLevelOld(bool),
    ListLevelPrevious(bool),
    ListLevelPreviousSpace(bool),
    ListLevelTemplateId(i32),

    // Sections
    SectionBreak,
    TitlePage,
    SectionEndnoteHere,
    OutlineLevel(i32),
    SectionContinuous,
    SectionColumn,
    SectionPage,
    SectionEvenPage,
    SectionOddPage,
    PageWidth(i32),
    PageHeight(i32),
    MarginLeft(i32),
    MarginRight(i32),
    MarginTop(i32),
    MarginBottom(i32),
    MarginGutter(i32),
    FacingPages(bool),
    MirrorMargins(Option<i32>),
    DocumentGutter(Option<i32>),
    HeaderDistance(i32),
    FooterDistance(i32),
    Landscape,
    Columns(Option<i32>),
    ColumnSpace(Option<i32>),
    ColumnNumber(Option<i32>),
    ColumnWidth(Option<i32>),
    ColumnSpaceRight(Option<i32>),
    ColumnSeparator(bool),
    PageNumberStart(i32),
    PageNumberFormat(crate::PageNumberFormat),
    PageNumberRestart(crate::PageNumberRestart),
    PageNumberOffsetX(i32),
    PageNumberOffsetY(i32),
    PageNumberHeadingLevel(Option<i32>),
    PageNumberHeadingSeparator(crate::PageNumberHeadingSeparator),
    SectionLineGrid(Option<i32>),
    SectionDocumentGrid(crate::SectionDocumentGridType),
    PaperSourceFirst(Option<i32>),
    PaperSourceOther(Option<i32>),
    ParagraphRevisionAuthor(i32),
    ParagraphRevisionDate(i32),
    SectionRevisionAuthor(i32),
    SectionRevisionDate(i32),
    TableRowRevisionAuthor(i32),
    TableRowRevisionDate(i32),
    CellRevisionMark(crate::CellRevisionKind),
    CellRevisionAuthor(crate::CellRevisionKind, i32),
    CellRevisionDate(crate::CellRevisionKind, i32),
    VerticalAlignTop,
    VerticalAlignCenter,
    VerticalAlignJustify,
    VerticalAlignBottom,
    LineNumbering(Option<i32>),
    LineNumberDistance(Option<i32>),
    LineNumberStart(Option<i32>),
    LineNumberRestartSection,
    LineNumberRestartPage,
    LineNumberContinuous,
    Header,
    HeaderFirst,
    HeaderLeft,
    HeaderRight,
    Footer,
    FooterFirst,
    FooterLeft,
    FooterRight,

    // Footnotes and endnotes
    Footnote,
    Endnote,
    FootnoteNumber(i32),
    NoteKinds(i32),
    FootnotePlacement(crate::NotePlacement),
    EndnotePlacement(crate::NotePlacement),
    FootnoteStart(i32),
    EndnoteStart(i32),
    FootnoteRestart(crate::FootnoteRestart),
    EndnoteRestart(crate::EndnoteRestart),
    FootnoteNumbering(crate::NoteNumberingStyle),
    EndnoteNumbering(crate::NoteNumberingStyle),
    SectionFootnotePlacement(crate::SectionFootnotePlacement),
    SectionFootnoteStart(i32),
    SectionEndnoteStart(i32),
    SectionFootnoteRestart(crate::FootnoteRestart),
    SectionEndnoteRestart(crate::EndnoteRestart),
    SectionFootnoteNumbering(crate::NoteNumberingStyle),
    SectionEndnoteNumbering(crate::NoteNumberingStyle),
    FootnoteSeparator,
    FootnoteContinuationSeparator,
    FootnoteContinuationNotice,
    EndnoteSeparator,
    EndnoteContinuationSeparator,
    EndnoteContinuationNotice,
    NoteSeparatorCharacter,
    NoteContinuationSeparatorCharacter,

    // Bookmarks
    BookmarkStart,
    BookmarkEnd,
    BookmarkFirstColumn(i32),
    BookmarkLastColumn(i32),
    BookmarkPublic,

    // Annotations/comments
    Annotation,
    AnnotationDate,
    AnnotationAuthor,
    AnnotationInitials,
    AnnotationReference,
    AnnotationParent,
    AnnotationRangeStart,
    AnnotationRangeEnd,
    AnnotationIcon,
    AnnotationTime,
    AnnotationMark,

    // Tracked revisions
    RevisionTable,
    Revised(bool),
    Deleted(bool),
    RevisionAuthor(i32),
    DeletedRevisionAuthor(i32),
    RevisionDate(i32),
    DeletedRevisionDate(i32),
    RevisionProperty(i32),

    // Shapes
    Shape(Option<i32>),
    ShapeInstance,
    ShapeType(i32),
    ShapeLeft(i32),
    ShapeTop(i32),
    ShapeRight(i32),
    ShapeBottom(i32),
    ShapeWidth(i32),
    ShapeHeight(i32),
    ShapeRotation(i32),
    ShapeZOrder(i32),
    ShapeWrap(i32),
    ShapeBelowText(bool),
    ShapeLockAnchor,
    ShapeGroup(Option<i32>),
    ShapeProperty,
    ShapePropertyName,
    ShapePropertyValue,
    ShapeText(Option<i32>),
    LegacyDrawingObject,
    LegacyTextBox,
    LegacyTextBoxText,
    LegacyAnchorXPage,
    LegacyAnchorXMargin,
    LegacyAnchorXColumn,
    LegacyAnchorYPage,
    LegacyAnchorYMargin,
    LegacyAnchorYParagraph,
    LegacyDrawingHeight(i32),
    LegacyTextBoxMargin(i32),
    LegacyDrawingX(i32),
    LegacyDrawingY(i32),
    LegacyDrawingWidth(i32),
    LegacyDrawingHeightSize(i32),
    LegacyTextLeftRightTopBottom,
    LegacyTextLeftRightTopBottomVertical,
    LegacyTextTopBottomRightLeft,
    LegacyTextTopBottomRightLeftVertical,
    LegacyTextBottomTopLeftRight,
    LegacyDrawingLock,
    LegacyDrawingGroup,
    LegacyDrawingCount(i32),
    LegacyDrawingEndGroup,
    LegacyDrawingArc,
    LegacyDrawingCallout,
    LegacyDrawingEllipse,
    LegacyDrawingLine,
    LegacyDrawingPolygon,
    LegacyDrawingPolyline,
    LegacyDrawingRectangle,
    LegacyDrawingRoundRectangle,
    LegacyDrawingPointX(i32),
    LegacyDrawingPointY(i32),
    LegacyDrawingArcFlipX,
    LegacyDrawingArcFlipY,
    LegacyCalloutType(crate::LegacyCalloutType),
    LegacyCalloutAngle(i32),
    LegacyCalloutAccent,
    LegacyCalloutSmartAttach,
    LegacyCalloutBestFit,
    LegacyCalloutMinusX,
    LegacyCalloutMinusY,
    LegacyCalloutBorder,
    LegacyCalloutAttachment(crate::LegacyCalloutAttachment),
    LegacyCalloutDescent(i32),
    LegacyCalloutOffset(i32),
    LegacyCalloutLength(i32),
    LegacyDrawingLineStyle(crate::LegacyDrawingLineStyle),
    LegacyDrawingLineGray(i32),
    LegacyDrawingLineRed(i32),
    LegacyDrawingLineGreen(i32),
    LegacyDrawingLineBlue(i32),
    LegacyDrawingLinePalette,
    LegacyDrawingLineWidth(i32),
    LegacyDrawingFillForegroundGray(i32),
    LegacyDrawingFillForegroundRed(i32),
    LegacyDrawingFillForegroundGreen(i32),
    LegacyDrawingFillForegroundBlue(i32),
    LegacyDrawingFillForegroundPalette,
    LegacyDrawingFillBackgroundGray(i32),
    LegacyDrawingFillBackgroundRed(i32),
    LegacyDrawingFillBackgroundGreen(i32),
    LegacyDrawingFillBackgroundBlue(i32),
    LegacyDrawingFillBackgroundPalette,
    LegacyDrawingFillPattern(i32),
    LegacyDrawingStartArrowFill(crate::LegacyDrawingArrowFill),
    LegacyDrawingStartArrowLength(i32),
    LegacyDrawingStartArrowWidth(i32),
    LegacyDrawingEndArrowFill(crate::LegacyDrawingArrowFill),
    LegacyDrawingEndArrowLength(i32),
    LegacyDrawingEndArrowWidth(i32),
    LegacyDrawingShadow,
    LegacyDrawingShadowX(i32),
    LegacyDrawingShadowY(i32),
    BackgroundDestination(Option<i32>),

    // Document info
    Title,
    Subject,
    Author,
    Manager,
    Company,
    Operator,
    Category,
    Keywords,
    Comment,
    DocComment,
    HyperlinkBase,
    CreationTime,
    RevisionTime,
    PrintTime,
    BackupTime,
    InfoVersion(i32),
    InfoRevision(i32),
    EditingTime(i32),
    NumberOfPages(i32),
    NumberOfWords(i32),
    NumberOfCharacters(i32),
    NumberOfCharactersWithSpaces(i32),
    DocumentId(i32),
    Year(i32),
    Month(i32),
    Day(i32),
    Hour(i32),
    Minute(i32),
    Second(i32),

    // Unicode
    Unicode(i32),
    UnicodeSkip(i32),

    // Special
    Tab,
    Line,
    Section,
    SectionDefault,
    NonBreakingSpace,
    OptionalHyphen,
    NonBreakingHyphen,
    EmDash,
    EnDash,
    EmSpace,
    EnSpace,
    QuarterEmSpace,
    Bullet,
    LeftToRightMark,
    RightToLeftMark,
    ZeroWidthJoiner,
    ZeroWidthNonJoiner,
    CurrentDate,
    CurrentDateLong,
    CurrentDateAbbreviated,
    CurrentTime,

    // Binary data
    Binary(i32),

    // Ignorable destination
    IgnorableDestination,

    // Unknown control word
    Unknown(&'a str, Option<i32>),
}

/// Token types.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Token<'a> {
    /// Opening brace
    OpenBrace,
    /// Closing brace
    CloseBrace,
    /// Control word
    Control(ControlWord<'a>),
    /// Plain text
    Text(Cow<'a, str>),
    /// Exact payload consumed by a `binN` control word
    Binary(Cow<'a, [u8]>),
}

/// Character set encoding for RTF.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CharacterSet {
    /// ANSI (Windows-1252 / CP1252)
    #[default]
    Ansi,
    /// Mac (Mac Roman)
    Mac,
    /// PC (DOS / CP437)
    Pc,
    /// PC (DOS / CP850)
    Pca,
}

/// RTF Lexer using arena allocation.
pub struct Lexer<'a> {
    /// Source input
    input: &'a str,
    /// Current position in bytes
    pos: usize,
    /// Arena allocator for temporary strings
    arena: &'a Bump,
}

impl<'a> Lexer<'a> {
    /// Create a new lexer.
    #[inline]
    pub fn new(input: &'a str, arena: &'a Bump) -> Self {
        Self {
            input,
            pos: 0,
            arena,
        }
    }

    /// Tokenize the entire input.
    pub fn tokenize(&mut self) -> RtfResult<Vec<Token<'a>>> {
        let mut tokens = Vec::new();

        while self.pos < self.input.len() {
            let token = self.next_token()?;
            tokens.push(token);
        }

        Ok(tokens)
    }

    /// Get the next token.
    fn next_token(&mut self) -> RtfResult<Token<'a>> {
        if self.pos >= self.input.len() {
            return Err(RtfError::UnexpectedEof);
        }

        let ch = self.current_char();
        match ch {
            '{' => {
                self.advance();
                Ok(Token::OpenBrace)
            },
            '}' => {
                self.advance();
                Ok(Token::CloseBrace)
            },
            '\\' => self.parse_control_word(),
            _ => self.parse_text(),
        }
    }

    /// Parse a control word or control symbol.
    fn parse_control_word(&mut self) -> RtfResult<Token<'a>> {
        self.advance(); // Skip '\'

        if self.pos >= self.input.len() {
            return Err(RtfError::UnexpectedEof);
        }

        let ch = self.current_char();

        // Handle special control symbols
        match ch {
            '\\' | '{' | '}' => {
                let text = self.arena.alloc_str(&ch.to_string());
                self.advance();
                return Ok(Token::Text(Cow::Borrowed(text)));
            },
            '\'' => return self.parse_hex_char(),
            '*' => {
                self.advance();
                return Ok(Token::Control(ControlWord::IgnorableDestination));
            },
            '\n' | '\r' => {
                self.advance();
                return Ok(Token::Control(ControlWord::Par));
            },
            '~' => {
                self.advance();
                return Ok(Token::Control(ControlWord::NonBreakingSpace));
            },
            '-' => {
                self.advance();
                return Ok(Token::Control(ControlWord::OptionalHyphen));
            },
            '_' => {
                self.advance();
                return Ok(Token::Control(ControlWord::NonBreakingHyphen));
            },
            _ => {},
        }

        // Parse control word
        let start = self.pos;

        // Read alphabetic characters
        while self.pos < self.input.len() && self.current_char().is_ascii_alphabetic() {
            self.advance();
        }

        if start == self.pos {
            // No alphabetic characters, might be a control symbol
            return Err(RtfError::InvalidControlWord(format!(
                "Invalid control word at position {}",
                self.pos
            )));
        }

        let word = &self.input[start..self.pos];

        // Parse optional numeric parameter
        let param = self.parse_numeric_parameter()?;

        // Skip optional space delimiter after control word
        if self.pos < self.input.len() && self.current_char() == ' ' {
            self.advance();
        }

        // Match control word to enum variant
        let control = self.match_control_word(word, param)?;

        // Handle binary data immediately after \bin. The input uses a one-byte
        // Latin-1 transport mapping, so each scalar maps back to its source byte.
        if let ControlWord::Binary(size) = control {
            let size = usize::try_from(size).map_err(|_| {
                RtfError::MalformedDocument("RTF binary length cannot be negative".to_string())
            })?;
            let mut data = Vec::with_capacity(size);
            for _ in 0..size {
                if self.pos >= self.input.len() {
                    return Err(RtfError::UnexpectedEof);
                }
                let value = u8::try_from(self.current_char() as u32).map_err(|_| {
                    RtfError::MalformedDocument(
                        "RTF binary payload is not byte-preserving".to_string(),
                    )
                })?;
                data.push(value);
                self.advance();
            }
            let allocated = self.arena.alloc_slice_copy(&data);
            return Ok(Token::Binary(Cow::Borrowed(allocated)));
        }

        Ok(Token::Control(control))
    }

    /// Parse numeric parameter after control word.
    fn parse_numeric_parameter(&mut self) -> RtfResult<Option<i32>> {
        if self.pos >= self.input.len() {
            return Ok(None);
        }

        let ch = self.current_char();
        if !ch.is_ascii_digit() && ch != '-' {
            return Ok(None);
        }

        let start = self.pos;
        if ch == '-' {
            self.advance();
        }

        while self.pos < self.input.len() && self.current_char().is_ascii_digit() {
            self.advance();
        }

        let num_str = &self.input[start..self.pos];
        let num = num_str.parse::<i32>()?;
        Ok(Some(num))
    }

    /// Match control word string to enum variant.
    fn match_control_word(&self, word: &'a str, param: Option<i32>) -> RtfResult<ControlWord<'a>> {
        let param_value = param.unwrap_or(1);
        let param_bool = param.unwrap_or(1) != 0;

        #[allow(clippy::match_same_arms)]
        let control = match word {
            // Document
            "rtf" => ControlWord::Rtf(param_value),
            "ansi" => ControlWord::Ansi,
            "ansicpg" => ControlWord::AnsiCodePage(param_value),
            "mac" => ControlWord::Mac,
            "pc" => ControlWord::Pc,
            "pca" => ControlWord::Pca,

            // Headers
            "fonttbl" => ControlWord::FontTable,
            "colortbl" => ControlWord::ColorTable,
            "stylesheet" => ControlWord::StyleSheet,
            "info" => ControlWord::Info,
            "defchp" => ControlWord::DefaultCharacterProperties(param),
            "defpap" => ControlWord::DefaultParagraphProperties(param),
            "deff" => ControlWord::DefaultFont(param),
            "adeff" => ControlWord::AssociatedDefaultFont(param),
            "stshfbi" => ControlWord::StylesheetDefaultBidiFont(param),
            "stshfdbch" => ControlWord::StylesheetDefaultDoubleByteFont(param),
            "stshfhich" => ControlWord::StylesheetDefaultHighAnsiFont(param),
            "stshfloch" => ControlWord::StylesheetDefaultLowAnsiFont(param),
            "pn" => ControlWord::LegacyParagraphNumbering(param),
            "pnlvl" => ControlWord::LegacyNumberingLevel(param),
            "pnlvlblt" => ControlWord::LegacyNumberingLevelBullet(param),
            "pnlvlbody" => ControlWord::LegacyNumberingLevelBody(param),
            "pnlvlcont" => ControlWord::LegacyNumberingLevelContinue(param),
            "pnseclvl" => ControlWord::LegacySectionNumberingLevel(param_value),
            "pndec" => ControlWord::LegacyNumberingDecimal(param),
            "pnucrm" => ControlWord::LegacyNumberingUpperRoman(param),
            "pnlcrm" => ControlWord::LegacyNumberingLowerRoman(param),
            "pnucltr" => ControlWord::LegacyNumberingUpperLetter(param),
            "pnlcltr" => ControlWord::LegacyNumberingLowerLetter(param),
            "pnaiu" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::Aiueo,
                param,
            ),
            "pnaiud" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::AiueoDbChar,
                param,
            ),
            "pnaiueo" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::AiueoExtended,
                param,
            ),
            "pnaiueod" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::AiueoExtendedDbChar,
                param,
            ),
            "pnchosung" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::Chosung,
                param,
            ),
            "pncard" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::CardinalText,
                param,
            ),
            "pndecd" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::DecimalWithPeriod,
                param,
            ),
            "pnord" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::Ordinal,
                param,
            ),
            "pnordt" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::OrdinalText,
                param,
            ),
            "pncnum" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::ChineseCounting,
                param,
            ),
            "pndbnum" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::ChineseCountingDbChar,
                param,
            ),
            "pndbnumd" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::ChineseCountingKorean,
                param,
            ),
            "pndbnumk" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::ChineseCountingLegal,
                param,
            ),
            "pndbnuml" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::ChineseCountingThousand,
                param,
            ),
            "pndbnumt" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::ChineseCountingTraditional,
                param,
            ),
            "pnganada" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::Ganada,
                param,
            ),
            "pngbnum" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::GbCounting,
                param,
            ),
            "pngbnumd" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::GbCountingDbChar,
                param,
            ),
            "pngbnumk" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::GbCountingKorean,
                param,
            ),
            "pngbnuml" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::GbCountingLegal,
                param,
            ),
            "pniroha" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::Iroha,
                param,
            ),
            "pnirohad" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::IrohaDbChar,
                param,
            ),
            "pnzodiac" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::Zodiac,
                param,
            ),
            "pnzodiacd" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::ZodiacDbChar,
                param,
            ),
            "pnzodiacl" => ControlWord::LegacyNumberingFormat(
                crate::LegacyParagraphNumberingFormat::ZodiacLegal,
                param,
            ),
            "pnstart" => ControlWord::LegacyNumberingStart(param),
            "pnindent" => ControlWord::LegacyNumberingIndent(param),
            "pnsp" => ControlWord::LegacyNumberingSpace(param),
            "pnhang" => ControlWord::LegacyNumberingHanging(param),
            "pnprev" => ControlWord::LegacyNumberingPrevious(param),
            "pnacross" => ControlWord::LegacyNumberingAcross(param),
            "pnnumonce" => ControlWord::LegacyNumberingOnce(param),
            "pnrestart" => ControlWord::LegacyNumberingRestart(param),
            "pnbidia" => ControlWord::LegacyNumberingBidiA(param),
            "pnbidib" => ControlWord::LegacyNumberingBidiB(param),
            "pnql" => ControlWord::LegacyNumberingAlignLeft(param),
            "pnqc" => ControlWord::LegacyNumberingAlignCenter(param),
            "pnqr" => ControlWord::LegacyNumberingAlignRight(param),
            "pnf" => ControlWord::LegacyNumberingFont(param),
            "pnfs" => ControlWord::LegacyNumberingFontSize(param),
            "pncf" => ControlWord::LegacyNumberingColor(param),
            "pnb" => ControlWord::LegacyNumberingBold(param),
            "pni" => ControlWord::LegacyNumberingItalic(param),
            "pncaps" => ControlWord::LegacyNumberingCaps(param),
            "pnscaps" => ControlWord::LegacyNumberingSmallCaps(param),
            "pnstrike" => ControlWord::LegacyNumberingStrike(param),
            "pnul" => ControlWord::LegacyNumberingUnderlineToggle(param),
            "pnuld" => ControlWord::LegacyNumberingUnderline(
                crate::LegacyParagraphNumberingUnderline::Dotted,
                param,
            ),
            "pnuldash" => ControlWord::LegacyNumberingUnderline(
                crate::LegacyParagraphNumberingUnderline::Dashed,
                param,
            ),
            "pnuldashd" => ControlWord::LegacyNumberingUnderline(
                crate::LegacyParagraphNumberingUnderline::DashDot,
                param,
            ),
            "pnuldashdd" => ControlWord::LegacyNumberingUnderline(
                crate::LegacyParagraphNumberingUnderline::DashDotDot,
                param,
            ),
            "pnuldb" => ControlWord::LegacyNumberingUnderline(
                crate::LegacyParagraphNumberingUnderline::Double,
                param,
            ),
            "pnulhair" => ControlWord::LegacyNumberingUnderline(
                crate::LegacyParagraphNumberingUnderline::Hairline,
                param,
            ),
            "pnulnone" => ControlWord::LegacyNumberingUnderline(
                crate::LegacyParagraphNumberingUnderline::None,
                param,
            ),
            "pnulth" => ControlWord::LegacyNumberingUnderline(
                crate::LegacyParagraphNumberingUnderline::Thick,
                param,
            ),
            "pnulw" => ControlWord::LegacyNumberingUnderline(
                crate::LegacyParagraphNumberingUnderline::Words,
                param,
            ),
            "pnulwave" => ControlWord::LegacyNumberingUnderline(
                crate::LegacyParagraphNumberingUnderline::Wave,
                param,
            ),
            "pnrauth" => ControlWord::LegacyNumberingRevisionAuthor(param),
            "pnrdate" => ControlWord::LegacyNumberingRevisionDate(param),
            "pnrnfc" => ControlWord::LegacyNumberingRevisionFormat(param),
            "pnrnot" => ControlWord::LegacyNumberingRevisionNoTrack(param),
            "pnrpnbr" => ControlWord::LegacyNumberingRevisionParagraph(param),
            "pnrrgb" => ControlWord::LegacyNumberingRevisionRgb(param),
            "pnrstart" => ControlWord::LegacyNumberingRevisionStart(param),
            "pnrstop" => ControlWord::LegacyNumberingRevisionStop(param),
            "pnrxst" => ControlWord::LegacyNumberingRevisionTextStart(param),
            "pntxtb" => ControlWord::LegacyNumberingTextBefore,
            "pntxta" => ControlWord::LegacyNumberingTextAfter,
            "pntext" => ControlWord::LegacyGeneratedListText,
            "pgptbl" => ControlWord::ParagraphGroupTable,
            "pgp" => ControlWord::ParagraphGroup,
            "ipgp" => ControlWord::ParagraphGroupParent(param_value),
            "itap" => ControlWord::TableNestingLevel(param),

            // Stylesheet entries and metadata
            "s" => ControlWord::ParagraphStyle(param),
            "cs" => ControlWord::CharacterStyle(param),
            "ds" => ControlWord::SectionStyle(param),
            "ts" => ControlWord::TableStyle(param),
            "sbasedon" => ControlWord::StyleBasedOn(param_value),
            "snext" => ControlWord::StyleNext(param_value),
            "slink" => ControlWord::StyleLink(param_value),
            "additive" => ControlWord::StyleAdditive(param_bool),
            "sautoupd" => ControlWord::StyleAutoUpdate(param_bool),
            "shidden" => ControlWord::StyleHidden(param_bool),
            "slocked" => ControlWord::StyleLocked(param_bool),
            "ssemihidden" => ControlWord::StyleSemiHidden(param_bool),
            "sunhideused" => ControlWord::StyleUnhideWhenUsed(param_bool),
            "sqformat" => ControlWord::StyleQuickFormat(param_bool),
            "spriority" => ControlWord::StylePriority(param_value),
            "styrsid" => ControlWord::StyleRevisionId(param_value),
            "spersonal" => ControlWord::StylePersonal(param_bool),
            "scompose" => ControlWord::StyleCompose(param_bool),
            "sreply" => ControlWord::StyleReply(param_bool),

            // Embedded content
            "pict" => ControlWord::Picture,
            "picprop" => ControlWord::PictureProperties(param),
            "shplid" => ControlWord::PictureShapeId(param),
            "object" => {
                if param.is_some() {
                    ControlWord::InvalidObjectDestinationParameter
                } else {
                    ControlWord::Object
                }
            },
            "result" => {
                if param.is_some() {
                    ControlWord::InvalidObjectResultDestinationParameter
                } else {
                    ControlWord::Result
                }
            },
            "objclass" => ControlWord::ObjectClass,
            "docvar" => ControlWord::DocumentVariable,
            "userprops" => ControlWord::UserProperties,
            "propname" => ControlWord::PropertyName,
            "proptype" => ControlWord::PropertyType(param),
            "staticval" => ControlWord::StaticValue,
            "linkval" => ControlWord::LinkValue,
            "objname" => ControlWord::ObjectName,
            "objdata" => ControlWord::ObjectData,
            "objemb" => {
                if param.is_none() {
                    ControlWord::ObjectEmbedded
                } else {
                    ControlWord::InvalidObjectModifierParameter
                }
            },
            "objlink" => {
                if param.is_none() {
                    ControlWord::ObjectLink
                } else {
                    ControlWord::InvalidObjectModifierParameter
                }
            },
            "objautlink" => {
                if param.is_none() {
                    ControlWord::ObjectAutoLink
                } else {
                    ControlWord::InvalidObjectModifierParameter
                }
            },
            "objhtml" => {
                if param.is_none() {
                    ControlWord::ObjectHtml
                } else {
                    ControlWord::InvalidObjectModifierParameter
                }
            },
            "objsub" => ControlWord::ObjectSubscriber(param),
            "objpub" => ControlWord::ObjectPublisher(param),
            "objicemb" => ControlWord::ObjectInstallableCommand(param),
            "objocx" => ControlWord::ObjectOleControl(param),
            "linkself" => ControlWord::ObjectLinkSelf(param),
            "objw" => param.map_or(
                ControlWord::InvalidObjectModifierParameter,
                ControlWord::ObjectWidth,
            ),
            "objh" => param.map_or(
                ControlWord::InvalidObjectModifierParameter,
                ControlWord::ObjectHeight,
            ),
            "objalign" => ControlWord::ObjectAlignment(param),
            "objtransy" => ControlWord::ObjectTranslationY(param),
            "objcropt" => ControlWord::ObjectCropTop(param),
            "objcropb" => ControlWord::ObjectCropBottom(param),
            "objcropl" => ControlWord::ObjectCropLeft(param),
            "objcropr" => ControlWord::ObjectCropRight(param),
            "objscalex" => ControlWord::ObjectScaleX(param),
            "objscaley" => ControlWord::ObjectScaleY(param),
            "objlock" => {
                if param.is_none() {
                    ControlWord::ObjectLocked(true)
                } else {
                    ControlWord::InvalidObjectModifierParameter
                }
            },
            "objupdate" => {
                if param.is_none() {
                    ControlWord::ObjectUpdate(true)
                } else {
                    ControlWord::InvalidObjectModifierParameter
                }
            },
            "objsetsize" => {
                if param.is_none() {
                    ControlWord::ObjectSetSize(true)
                } else {
                    ControlWord::InvalidObjectModifierParameter
                }
            },
            "rsltmerge" => ControlWord::ObjectResultMerge(param),
            "rsltrtf" => ControlWord::ObjectResultRtf(param),
            "rslttxt" => ControlWord::ObjectResultText(param),
            "rsltpict" => ControlWord::ObjectResultPicture(param),
            "rsltbmp" => ControlWord::ObjectResultBitmap(param),
            "rslthtml" => ControlWord::ObjectResultHtml(param),
            "oleclsid" => ControlWord::OleClassId(param),

            // Picture properties
            "picw" => ControlWord::PictureWidth(param_value),
            "pich" => ControlWord::PictureHeight(param_value),
            "picwgoal" => ControlWord::PictureGoalWidth(param_value),
            "pichgoal" => ControlWord::PictureGoalHeight(param_value),
            "picscalex" => ControlWord::PictureScaleX(param_value),
            "picscaley" => ControlWord::PictureScaleY(param_value),
            "picscaled" => ControlWord::PictureScaled(param),
            "picbmp" => ControlWord::PictureBitmap(param),
            "picbpp" => ControlWord::PictureBitsPerPixel(param),
            "piccropl" => ControlWord::PictureCropLeft(param),
            "piccropr" => ControlWord::PictureCropRight(param),
            "piccropt" => ControlWord::PictureCropTop(param),
            "piccropb" => ControlWord::PictureCropBottom(param),
            "wbmbitspixel" => ControlWord::WindowsBitmapBitsPerPixel(param),
            "wbmplanes" => ControlWord::WindowsBitmapPlanes(param),
            "wbmwidthbytes" => ControlWord::WindowsBitmapWidthBytes(param),
            "emfblip" => ControlWord::Emfblip,
            "pngblip" => ControlWord::Pngblip,
            "jpegblip" => ControlWord::Jpegblip,
            "macpict" => ControlWord::Macpict,
            "pmmetafile" => ControlWord::Pmmetafile(param_value),
            "wmetafile" => ControlWord::Wmetafile(param_value),
            "dibitmap" => ControlWord::Dibitmap(param_value),
            "wbitmap" => ControlWord::Wbitmap(param_value),
            "bliptag" => ControlWord::BlipTag(param_value),
            "blipuid" => ControlWord::BlipUid,
            "blipupi" => ControlWord::BlipUnitsPerInch(param_value),

            // Field support
            "field" => ControlWord::Field,
            "fldinst" => ControlWord::FieldInstruction,
            "fldrslt" => ControlWord::FieldResult,
            "fldlock" => ControlWord::FieldLock(param),
            "flddirty" => ControlWord::FieldDirty(param),
            "fldedit" => ControlWord::FieldEdit(param),
            "fldpriv" => ControlWord::FieldPrivate(param),
            "formfield" => ControlWord::FormField,
            "datafield" => ControlWord::DataField,
            "fftype" => ControlWord::FormFieldType(param_value),
            "ffmaxlen" => ControlWord::FormFieldMaxLength(param_value),
            "ffprot" => ControlWord::FormFieldProtected(param_bool),
            "ffrecalc" => ControlWord::FormFieldRecalculate(param_bool),
            "ffsize" => ControlWord::FormFieldAutomaticSize(param_bool),
            "ffname" => ControlWord::FormFieldName,
            "ffformat" => ControlWord::FormFieldFormat,
            "ffdeftext" => ControlWord::FormFieldDefaultText,
            "ffdefres" => ControlWord::FormFieldDefaultResult(param_value),
            "ffres" => ControlWord::FormFieldResult(param_value),
            "ffhps" => ControlWord::FormFieldHalfPointSize(param_value),
            "ffownhelp" => ControlWord::FormFieldOwnHelp(param_bool),
            "ffownstat" => ControlWord::FormFieldOwnStatus(param_bool),
            "ffhelptext" => ControlWord::FormFieldHelpText,
            "ffstattext" => ControlWord::FormFieldStatusText,
            "ffentrymcr" => ControlWord::FormFieldEntryMacro,
            "ffexitmcr" => ControlWord::FormFieldExitMacro,
            "ffl" => ControlWord::FormFieldListEntry,
            "fftypetxt" => ControlWord::FormFieldTextType(param_value),
            "ffhaslistbox" => ControlWord::FormFieldHasListBox(param_bool),
            "generator" => ControlWord::Generator,
            "rsidtbl" => ControlWord::RevisionSaveTable,
            "rsid" => ControlWord::RevisionSaveId(param_value),
            "rsidroot" => ControlWord::RevisionSaveRoot(param_value),
            "insrsid" => ControlWord::InsertRsid(param_value),
            "delrsid" => ControlWord::DeleteRsid(param_value),
            "charrsid" => ControlWord::CharStyleRsid(param_value),
            "pararsid" => ControlWord::ParagraphRsid(param_value),
            "sectrsid" => ControlWord::SectionRsid(param_value),
            "tblrsid" => ControlWord::TableRsid(param_value),
            "xmlnstbl" => ControlWord::XmlNamespaceTable,
            "xmlns" => ControlWord::XmlNamespace(param_value),
            "xmlopen" => ControlWord::XmlOpen,
            "xmlclose" => ControlWord::XmlClose,
            "xmlattrname" => ControlWord::XmlAttributeName,
            "xmlattrvalue" => ControlWord::XmlAttributeValue,
            "mmath" | "moMath" => ControlWord::MathZoneInline,
            "mmathPara" | "moMathPara" => ControlWord::MathZoneDisplay,
            "mmathParaPr" | "moMathParaPr" => ControlWord::MathZoneParagraphProperties,
            "mmathPict" => ControlWord::MathZonePictureFallback,
            "macc" => ControlWord::MathAccent,
            "mbar" => ControlWord::MathBar,
            "mborderBox" => ControlWord::MathBorderBox,
            "mbox" => ControlWord::MathBox,
            "md" => ControlWord::MathDelimiter,
            "meqArr" => ControlWord::MathEquationArray,
            "mf" => ControlWord::MathFraction,
            "mfunc" => ControlWord::MathFunction,
            "mgroupChr" => ControlWord::MathGroupChar,
            "mlimlow" | "mlimLow" => ControlWord::MathLimitLower,
            "mlimupp" | "mlimUpp" => ControlWord::MathLimitUpper,
            "mm" => ControlWord::MathMatrix,
            "mnary" => ControlWord::MathNary,
            "mphant" => ControlWord::MathPhantom,
            "mrad" => ControlWord::MathRadical,
            "msPre" => ControlWord::MathScriptPre,
            "msSub" => ControlWord::MathScriptSub,
            "msSubSup" => ControlWord::MathScriptSubSup,
            "msSup" => ControlWord::MathScriptSup,
            "mr" => ControlWord::MathRun,
            "me" => ControlWord::MathElement,
            "mnum" => ControlWord::MathNumerator,
            "mden" => ControlWord::MathDenominator,
            "mdeg" => ControlWord::MathDegree,
            "msub" => ControlWord::MathSubscript,
            "msup" => ControlWord::MathSuperscript,
            "mlim" => ControlWord::MathLimit,
            "mfName" => ControlWord::MathFunctionName,
            "mmr" => ControlWord::MathMatrixRow,
            "maccPr" => ControlWord::MathAccentProperties,
            "mbarPr" => ControlWord::MathBarProperties,
            "mborderBoxPr" => ControlWord::MathBorderBoxProperties,
            "mboxPr" => ControlWord::MathBoxProperties,
            "mdPr" => ControlWord::MathDelimiterProperties,
            "meqArrPr" => ControlWord::MathEquationArrayProperties,
            "mfPr" => ControlWord::MathFractionProperties,
            "mfuncPr" => ControlWord::MathFunctionProperties,
            "mgroupChrPr" => ControlWord::MathGroupCharProperties,
            "mlimlowPr" | "mlimLowPr" => ControlWord::MathLimitLowerProperties,
            "mlimuppPr" | "mlimUppPr" => ControlWord::MathLimitUpperProperties,
            "mmPr" => ControlWord::MathMatrixProperties,
            "mnaryPr" => ControlWord::MathNaryProperties,
            "mphantPr" => ControlWord::MathPhantomProperties,
            "mradPr" => ControlWord::MathRadicalProperties,
            "msPrePr" => ControlWord::MathScriptPreProperties,
            "msSubPr" => ControlWord::MathScriptSubProperties,
            "msSubSupPr" => ControlWord::MathScriptSubSupProperties,
            "msSupPr" => ControlWord::MathScriptSupProperties,
            "mrPr" => ControlWord::MathRunProperties,
            "mctrlPr" => ControlWord::MathControlProperties,
            "mmcs" => ControlWord::MathMatrixColumns,
            "mmc" => ControlWord::MathMatrixColumn,
            "mmcPr" => ControlWord::MathMatrixColumnProperties,
            "margPr" => ControlWord::MathArgumentProperties,
            "margSz" => ControlWord::MathPropertyArgumentSize(param),
            "mtype" => ControlWord::MathPropertyType(param),
            "mgrow" => ControlWord::MathPropertyGrow(param),
            "mchr" => ControlWord::MathPropertyChar(param),
            "mbegChr" => ControlWord::MathPropertyBeginChar(param),
            "mendChr" => ControlWord::MathPropertyEndChar(param),
            "msepChr" => ControlWord::MathPropertySeparatorChar(param),
            "mpos" => ControlWord::MathPropertyPosition(param),
            "mvertJc" => ControlWord::MathPropertyVerticalJustify(param),
            "mbaseJc" => ControlWord::MathPropertyBaseJustify(param),
            "mjc" => ControlWord::MathPropertyJustify(param),
            "maln" => ControlWord::MathPropertyAlign(param),
            "malnScr" => ControlWord::MathPropertyAlignScript(param),
            "mdegHide" => ControlWord::MathPropertyDegreeHide(param),
            "mdiff" => ControlWord::MathPropertyDifferential(param),
            "mdiffSty" => ControlWord::MathPropertyDifferentialStyle(param),
            "mhideBot" => ControlWord::MathPropertyHideBottom(param),
            "mhideLeft" => ControlWord::MathPropertyHideLeft(param),
            "mhideRight" => ControlWord::MathPropertyHideRight(param),
            "mhideTop" => ControlWord::MathPropertyHideTop(param),
            "mlimLoc" | "mlimloc" => ControlWord::MathPropertyLimitLocation(param),
            "mplcHide" => ControlWord::MathPropertyPlaceholderHide(param),
            "msubHide" => ControlWord::MathPropertySubscriptHide(param),
            "msupHide" => ControlWord::MathPropertySuperscriptHide(param),
            "mstrikeBLTR" => ControlWord::MathPropertyStrikeBltr(param),
            "mstrikeH" => ControlWord::MathPropertyStrikeHorizontal(param),
            "mstrikeTLBR" => ControlWord::MathPropertyStrikeTlbr(param),
            "mstrikeV" => ControlWord::MathPropertyStrikeVertical(param),
            "msty" => ControlWord::MathPropertyStyle(param),
            "mscr" => ControlWord::MathPropertyScript(param),
            "mtransp" => ControlWord::MathPropertyTransparent(param),
            "mshow" => ControlWord::MathPropertyShow(param),
            "mshp" => ControlWord::MathPropertyShape(param),
            "mzeroAsc" => ControlWord::MathPropertyZeroAscent(param),
            "mzeroDesc" => ControlWord::MathPropertyZeroDescent(param),
            "mzeroWid" => ControlWord::MathPropertyZeroWidth(param),
            "mopEmu" => ControlWord::MathPropertyOperatorEmulator(param),
            "mnoBreak" => ControlWord::MathPropertyNoBreak(param),
            "mnor" => ControlWord::MathPropertyNormalText(param),
            "mlit" => ControlWord::MathPropertyLiteral(param),
            "mcGp" => ControlWord::MathPropertyMatrixColumnGap(param),
            "mcGpRule" => ControlWord::MathPropertyMatrixColumnGapRule(param),
            "mcSp" => ControlWord::MathPropertyMatrixColumnSpacing(param),
            "mcount" => ControlWord::MathPropertyMatrixCellCount(param),
            "mmcJc" => ControlWord::MathPropertyMatrixCellJustify(param),
            "mrSp" => ControlWord::MathPropertyRowSpacing(param),
            "mrSpRule" => ControlWord::MathPropertyRowSpacingRule(param),
            "mbrk" => ControlWord::MathPropertyBreak(param),
            "themedata" => ControlWord::ThemeData,
            "colorschememapping" => ControlWord::ColorSchemeMapping,
            "latentstyles" => ControlWord::LatentStyles,
            "lsdstimax" => ControlWord::LatentStyleMax(param_value),
            "lsdlockeddef" => ControlWord::LatentStyleLockedDefault(param_value),
            "lsdsemihiddendef" => ControlWord::LatentStyleSemiHiddenDefault(param_value),
            "lsdunhideuseddef" => ControlWord::LatentStyleUnhideUsedDefault(param_value),
            "lsdqformatdef" => ControlWord::LatentStyleQuickFormatDefault(param_value),
            "lsdprioritydef" => ControlWord::LatentStylePriorityDefault(param_value),
            "lsdlockedexcept" => ControlWord::LatentStyleExceptions,
            "lsdlocked" => ControlWord::LatentStyleLocked(param_value),
            "lsdsemihidden" => ControlWord::LatentStyleSemiHidden(param_value),
            "lsdunhideused" => ControlWord::LatentStyleUnhideUsed(param_value),
            "lsdqformat" => ControlWord::LatentStyleQuickFormat(param_value),
            "lsdpriority" => ControlWord::LatentStylePriority(param_value),
            "datastore" => ControlWord::DataStore,
            "mailmerge" => ControlWord::MailMerge,
            "mmconnectstr" => ControlWord::MailMergeConnectString,
            "mmconnectstrdata" => ControlWord::MailMergeConnectStringData,
            "mmdatasource" => ControlWord::MailMergeDataSource,
            "mmheadersource" => ControlWord::MailMergeHeaderSource,
            "mmlinktoquery" => ControlWord::MailMergeLinkToQuery(param_bool),
            "mmquery" => ControlWord::MailMergeQuery,
            "mmodso" => ControlWord::MailMergeDataSourceObject,
            "mmodsoactive" => ControlWord::MailMergeActiveRecord(param_value),
            "mmodsocoldelim" => ControlWord::MailMergeColumnDelimiter(param_value),
            "mmodsocolumn" => ControlWord::MailMergeColumnCount(param_value),
            "mmodsodynaddr" => ControlWord::MailMergeDynamicAddress(param_bool),
            "mmodsofhdr" => ControlWord::MailMergeFirstRowHeader(param_bool),
            "mmodsofilter" => ControlWord::MailMergeFilter,
            "mmodsofldmpdata" => ControlWord::MailMergeFieldMapData,
            "mmodsofmcolumn" => ControlWord::MailMergeFieldMapColumn(param_value),
            "mmodsohash" => ControlWord::MailMergeHash(param_value),
            "mmodsolid" => ControlWord::MailMergeId(param_value),
            "mmodsomappedname" => ControlWord::MailMergeMappedName,
            "mmodsoname" => ControlWord::MailMergeName,
            "mmodsorecipdata" => ControlWord::MailMergeRecipientData,
            "mmodsosort" => ControlWord::MailMergeSort,
            "mmodsosrc" => ControlWord::MailMergeSourceType(param_value),
            "mmodsotable" => ControlWord::MailMergeTable,
            "mmodsoudl" => ControlWord::MailMergeUdl,
            "mmodsoudldata" => ControlWord::MailMergeUdlData,
            "mmodsouniquetag" => ControlWord::MailMergeUniqueTag,
            "mmathPr" => ControlWord::MathProperties,
            "mbrkBin" => ControlWord::MathBreakBinary(param.unwrap_or(0)),
            "mbrkBinSub" => ControlWord::MathBreakBinarySubtraction(param.unwrap_or(0)),
            "mdefJc" => ControlWord::MathDefaultJustification(param.unwrap_or(1)),
            "mdispDef" => ControlWord::MathDisplayDefaults(param.unwrap_or(1)),
            "minterSp" => ControlWord::MathInterEquationSpacing(param.unwrap_or(0)),
            "mintLim" => ControlWord::MathIntegralLimitPlacement(param.unwrap_or(0)),
            "mintraSp" => ControlWord::MathIntraEquationSpacing(param.unwrap_or(0)),
            "mlMargin" => ControlWord::MathLeftMargin(param.unwrap_or(0)),
            "mmathFont" => ControlWord::MathFont(param.unwrap_or(0)),
            "mnaryLim" => ControlWord::MathNaryLimitPlacement(param.unwrap_or(1)),
            "mpostSp" => ControlWord::MathPostSpacing(param.unwrap_or(0)),
            "mpreSp" => ControlWord::MathPreSpacing(param.unwrap_or(0)),
            "mrMargin" => ControlWord::MathRightMargin(param.unwrap_or(0)),
            "msmallFrac" => ControlWord::MathSmallFractions(param.unwrap_or(0)),
            "mwrapIndent" => ControlWord::MathWrapIndent(param.unwrap_or(1440)),
            "mwrapRight" => ControlWord::MathWrapRight(param.unwrap_or(0)),
            "deftab" => ControlWord::DefaultTabWidth(param),
            "deflang" => ControlWord::DefaultLanguage(param_value),
            "deflangfe" => ControlWord::DefaultLanguageEastAsian(param_value),
            "adeflang" => ControlWord::DefaultLanguageComplexScript(param_value),
            "lang" => ControlWord::Language(param_value),
            "langfe" => ControlWord::LanguageEastAsian(param_value),
            "langnp" => ControlWord::LanguageNoProof(param_value),
            "langfenp" => ControlWord::LanguageEastAsianNoProof(param_value),
            "noproof" => ControlWord::NoProof(param_bool),
            "ltrch" => ControlWord::LeftToRightCharacter,
            "rtlch" => ControlWord::RightToLeftCharacter,
            "ltrpar" => ControlWord::LeftToRightParagraph,
            "rtlpar" => ControlWord::RightToLeftParagraph,
            "ltrdoc" => ControlWord::LeftToRightDocument,
            "rtldoc" => ControlWord::RightToLeftDocument,
            "ltrsect" => ControlWord::LeftToRightSection,
            "rtlsect" => ControlWord::RightToLeftSection,
            "ltrrow" => ControlWord::LeftToRightRow(param),
            "rtlrow" => ControlWord::RightToLeftRow(param),
            "taprtl" => ControlWord::TableRightToLeft(param_bool),
            "rtlgutter" => ControlWord::RightGutter(param_bool),
            "formprot" => ControlWord::FormProtection(param),
            "annotprot" => ControlWord::AnnotationProtection(param),
            "revprot" => ControlWord::RevisionProtection(param),
            "readprot" => ControlWord::ReadOnlyProtection(param),
            "allprot" => ControlWord::AllProtection(param),
            "enforceprot" => ControlWord::EnforceProtection(param),
            "protlevel" => ControlWord::ProtectionLevel(param),
            "password" => ControlWord::Password,
            "protusertbl" => ControlWord::ProtectionUserTable,
            "hyphauto" => ControlWord::HyphenateAutomatically(param),
            "hyphcaps" => ControlWord::HyphenateCapitalizedWords(param),
            "hyphconsec" => ControlWord::HyphenationConsecutiveLines(param),
            "hyphhotz" => ControlWord::HyphenationHotZone(param),
            "nextfile" => ControlWord::NextFile,
            "template" => ControlWord::DocumentTemplate,
            "viewkind" => ControlWord::DocumentViewKind(param),
            "viewscale" => ControlWord::DocumentViewScale(param),
            "viewzk" => ControlWord::DocumentZoomKind(param),
            "viewbksp" => ControlWord::DocumentViewBackgroundShapes(param),
            "viewnobound" => ControlWord::DocumentViewNoPageBoundaries(param),
            "donotshowmarkup" => ControlWord::HideReviewMarkup(param),
            "donotshowcomments" => ControlWord::HideReviewComments(param),
            "donotshowinsdel" => ControlWord::HideReviewInsertionsAndDeletions(param),
            "windowcaption" => ControlWord::WindowCaption,
            "xform" => ControlWord::XslTransform,
            "usexform" => ControlWord::UseXslTransform(param),
            "wgrffmtfilter" => ControlWord::StyleListFilter(param),
            "stylesortmethod" => ControlWord::StyleSortMethod(param),
            "readonlyrecommended" => ControlWord::ReadOnlyRecommended(param),
            "saveprevpict" => ControlWord::SavePreviousPicture(param),
            "writereservation" => ControlWord::WriteReservation(param),
            "writereservhash" => ControlWord::WriteReservationHash(param),
            "fromtext" => ControlWord::FromText(param),
            "fromhtml" => ControlWord::FromHtml(param),
            "doctype" => ControlWord::DocumentType(param),
            "makebackup" => ControlWord::MakeBackup(param),
            "defformat" => ControlWord::DefaultSaveFormat(param),
            "doctemp" => ControlWord::BoilerplateDocument(param),
            "muser" => ControlWord::Word97CompatibilityMode(param),
            "psover" => ControlWord::PostScriptOverText(param),
            "horzdoc" => ControlWord::HorizontalDocument(param),
            "vertdoc" => ControlWord::VerticalDocument(param),
            "jcompress" => ControlWord::CompressJustification(param),
            "jexpand" => ControlWord::ExpandJustification(param),
            "lnongrid" => ControlWord::LineBasedOnGrid(param),
            "fracwidth" => ControlWord::FractionalCharacterWidths(param),
            "ilfomacatclnup" => ControlWord::AbstractNumberingCleanup(param),
            "grfdocevents" => ControlWord::DocumentEventMask(param),
            "dgmargin" => ControlWord::DrawingGridFollowsMargins(param),
            "dgsnap" => ControlWord::SnapToDrawingGrid(param),
            "dghspace" => ControlWord::DrawingGridHorizontalSpacing(param),
            "dgvspace" => ControlWord::DrawingGridVerticalSpacing(param),
            "dghorigin" => ControlWord::DrawingGridHorizontalOrigin(param),
            "dgvorigin" => ControlWord::DrawingGridVerticalOrigin(param),
            "dghshow" => ControlWord::DrawingGridHorizontalShow(param),
            "dgvshow" => ControlWord::DrawingGridVerticalShow(param),
            "gutterprl" => ControlWord::ParallelGutter(param),
            "twoonone" => ControlWord::PrintTwoOnOne(param),
            "themelang" => ControlWord::ThemeLanguage(param),
            "themelangfe" => ControlWord::ThemeLanguageEastAsian(param),
            "themelangcs" => ControlWord::ThemeLanguageComplexScript(param),
            "relyonvml" => ControlWord::RelyOnVml(param),
            "validatexml" => ControlWord::ValidateXml(param),
            "showplaceholdtext" => ControlWord::ShowPlaceholderText(param),
            "ignoremixedcontent" => ControlWord::IgnoreMixedContent(param),
            "saveinvalidxml" => ControlWord::SaveInvalidXml(param),
            "showxmlerrors" => ControlWord::ShowXmlErrors(param),
            "donotembedsysfont" => ControlWord::DoNotEmbedSystemFonts(param),
            "donotembedlingdata" => ControlWord::DoNotEmbedLinguisticData(param),
            "trackmoves" => ControlWord::TrackMoves(param),
            "trackformatting" => ControlWord::TrackFormatting(param),
            "stylelocktheme" => ControlWord::LockDocumentTheme(param),
            "stylelockqfset" => ControlWord::LockQuickFormatSet(param),
            "usenormstyforlist" => ControlWord::UseNormalStyleForLists(param),
            "linkstyles" => ControlWord::UpdateStylesFromTemplate(param),
            "stylelock" => ControlWord::DeclareStyleRestrictions(param),
            "stylelockenforced" => ControlWord::EnforceStyleRestrictions(param),
            "stylelockbackcomp" => ControlWord::StyleRestrictionsBackwardCompatibility(param),
            "autofmtoverride" => ControlWord::AllowAutoFormatOverride(param),
            "bookfold" => ControlWord::BookFold(param),
            "bookfoldrev" => ControlWord::ReverseBookFold(param),
            "bookfoldsheets" => ControlWord::BookFoldSheets(param),
            "rempersonalinfo" => ControlWord::RemovePersonalInformation(param),
            "remdttm" => ControlWord::RemoveDateTimeInformation(param),
            "noextrasprl" => ControlWord::SuppressRaisedLoweredExtraSpacing(param),
            "sprstsp" => ControlWord::SuppressTopPageExtraSpacing(param),
            "sprsspbf" => ControlWord::SuppressSpaceBeforeAfterHardBreak(param),
            "sprslnsp" => ControlWord::SuppressWordPerfectExtraLineSpacing(param),
            "sprsbsp" => ControlWord::SuppressBottomPageExtraSpacing(param),
            "dntblnsbdb" => ControlWord::DoNotBalanceSbcsDbcs(param),
            "expshrtn" => ControlWord::ExpandSpacingAtShiftReturn(param),
            "nospaceforul" => ControlWord::DoNotAddSpaceForUnderline(param),
            "noultrlspc" => ControlWord::DoNotUnderlineTrailingSpaces(param),
            "noxlattoyen" => ControlWord::DoNotTranslateBackslashToYen(param),
            "lnbrkrule" => ControlWord::LegacyAsianLineBreakingRules(param),
            "otblrul" => ControlWord::CombineLegacyTableBorders(param),
            "alntblind" => ControlWord::DoNotAlignTableRowsIndependently(param),
            "lytcalctblwd" => ControlWord::DoNotUseRawTableWidth(param),
            "lyttblrtgr" => ControlWord::KeepTableRowsTogether(param),
            "nolnhtadjtbl" => ControlWord::DoNotAdjustTableLineHeight(param),
            "nobrkwrptbl" => ControlWord::DoNotBreakWrappedTablesAcrossPages(param),
            "nogrowautofit" => ControlWord::PreventAutofitGrowthIntoMargins(param),
            "newtblstyruls" => ControlWord::UseWord2003TableStyleRules(param),
            "splytwnine" => ControlWord::DoNotUseWord97ShapeLayout(param),
            "ftnlytwnine" => ControlWord::UseLegacyFootnoteLayout(param),
            "htmautsp" => ControlWord::UseHtmlParagraphAutoSpacing(param),
            "useltbaln" => ControlWord::PreserveLastTabAlignment(param),
            "oldas" => ControlWord::UseWord95AutoSpacing(param),
            "ApplyBrkRules" => ControlWord::ApplyThaiLineBreakingRules(param),
            "snaptogridincell" => ControlWord::SnapTextToGridInsideTable(param),
            "wrppunct" => ControlWord::AllowHangingPunctuation(param),
            "asianbrkrule" => ControlWord::UseAsianLineBreakingRules(param),
            "toplinepunct" => ControlWord::CompressPunctuationAtLineStart(param),
            "nocompatoptions" => ControlWord::NoCompatibilityOptions(param),
            "nouicompat" => ControlWord::NoUiCompatibility(param),
            "nofeaturethrottle" => ControlWord::NoFeatureThrottle(param),
            "forceupgrade" => ControlWord::ForceCompatibilityUpgrade(param),
            "noafcnsttbl" => ControlWord::PreserveAutofitTableWidthAroundShapes(param),
            "noindnmbrts" => ControlWord::UseHangingIndentAsNumberingTab(param),
            "felnbrelev" => ControlWord::UseLegacyKinsokuCharacters(param),
            "indrlsweleven" => ControlWord::UseLegacyFloatingObjectIndentation(param),
            "nocxsptable" => ControlWord::AllowContextualSpacingInTables(param),
            "notcvasp" => ControlWord::IgnoreCellVerticalAlignmentWithFloatingObjects(param),
            "notvatxbx" => ControlWord::IgnoreTextBoxVerticalAlignment(param),
            "spltpgpar" => ControlWord::SplitPageBreakParagraph(param),
            "hwelev" => ControlWord::UseFixedWidthHangul(param),
            "afelev" => ControlWord::UseLegacyAutofitWidthExpansion(param),
            "cachedcolbal" => ControlWord::UseCachedColumnBalancing(param),
            "utinl" => ControlWord::UnderlineNumberingSuffix(param),
            "notbrkcnstfrctbl" => ControlWord::DoNotSplitRowsAroundFloatingTables(param),
            "krnprsnet" => ControlWord::UseAnsiKerningPairs(param),

            // Index and table-of-contents source marks
            "xe" => ControlWord::IndexEntry,
            "xef" => ControlWord::IndexIdentifier(param),
            "bxe" => ControlWord::IndexBold(param_bool),
            "ixe" => ControlWord::IndexItalic(param_bool),
            "txe" => ControlWord::IndexReplacementText,
            "rxe" => ControlWord::IndexBookmarkRange,
            "yxe" => ControlWord::IndexYomi,
            "pxe" => ControlWord::IndexPronunciation,
            "tc" => ControlWord::TableOfContentsEntry,
            "tcn" => ControlWord::TableOfContentsEntryNoPage,
            "tcf" => ControlWord::TableOfContentsTable(param.unwrap_or(67)),
            "tcl" => ControlWord::TableOfContentsLevel(param.unwrap_or(1)),

            // Fonts
            "f" => ControlWord::FontNumber(param_value),
            "fs" => ControlWord::FontSize(param_value),
            "af" => ControlWord::AssociatedFontNumber(param),
            "afs" => ControlWord::AssociatedFontSize(param),
            "alang" => ControlWord::AssociatedLanguage(param),
            "ab" => ControlWord::AssociatedBold(param),
            "acaps" => ControlWord::AssociatedAllCaps(param),
            "acf" => ControlWord::AssociatedColor(param),
            "adn" => ControlWord::AssociatedBaselineDown(param),
            "aexpnd" => ControlWord::AssociatedExpansion(param),
            "ai" => ControlWord::AssociatedItalic(param),
            "aoutl" => ControlWord::AssociatedOutline(param),
            "ascaps" => ControlWord::AssociatedSmallCaps(param),
            "ashad" => ControlWord::AssociatedShadow(param),
            "astrike" => ControlWord::AssociatedStrike(param),
            "aul" => ControlWord::AssociatedUnderline(param),
            "auld" => ControlWord::AssociatedUnderlineDotted(param),
            "auldb" => ControlWord::AssociatedUnderlineDouble(param),
            "aulnone" => ControlWord::AssociatedUnderlineNone(param),
            "aulw" => ControlWord::AssociatedUnderlineWords(param),
            "aup" => ControlWord::AssociatedBaselineUp(param),
            "fcharset" => ControlWord::FontCharset(param_value),
            "fnil" => ControlWord::FontFamily("nil"),
            "froman" => ControlWord::FontFamily("roman"),
            "fswiss" => ControlWord::FontFamily("swiss"),
            "fmodern" => ControlWord::FontFamily("modern"),
            "fscript" => ControlWord::FontFamily("script"),
            "fdecor" => ControlWord::FontFamily("decor"),
            "flomajor" => ControlWord::FontTheme("flomajor"),
            "fhimajor" => ControlWord::FontTheme("fhimajor"),
            "fdbmajor" => ControlWord::FontTheme("fdbmajor"),
            "fbimajor" => ControlWord::FontTheme("fbimajor"),
            "flominor" => ControlWord::FontTheme("flominor"),
            "fhiminor" => ControlWord::FontTheme("fhiminor"),
            "fdbminor" => ControlWord::FontTheme("fdbminor"),
            "fbiminor" => ControlWord::FontTheme("fbiminor"),
            "ftech" => ControlWord::FontFamily("tech"),
            "fprq" => ControlWord::FontPitch(param_value),
            "cpg" => ControlWord::FontCodePage(param_value),
            "falt" => ControlWord::FontAlternateName,
            "fname" => ControlWord::FontNonTaggedName,
            "panose" => ControlWord::FontPanose,
            "fontemb" => ControlWord::FontEmbedded,
            "ftnil" => ControlWord::FontEmbeddedType("nil"),
            "fttruetype" => ControlWord::FontEmbeddedType("truetype"),
            "fontfile" => ControlWord::FontFile,
            "filetbl" => ControlWord::FileTable,
            "file" => ControlWord::FileEntry,
            "fid" => ControlWord::FileId(param_value),
            "frelative" => ControlWord::FileRelative(param_value),
            "fosnum" => ControlWord::FileOperatingSystem(param_value),
            "fvalidmac" => ControlWord::FileValidMac,
            "fvaliddos" => ControlWord::FileValidDos,
            "fvalidntfs" => ControlWord::FileValidNtfs,
            "fvalidhpfs" => ControlWord::FileValidHpfs,
            "fnetwork" => ControlWord::FileNetwork,
            "fnonfilesys" => ControlWord::FileNonFileSystem,
            "upr" => ControlWord::UnicodeAlternate,
            "ud" => ControlWord::UnicodeAlternateDestination,

            // Colors
            "red" => ControlWord::Red(param_value),
            "green" => ControlWord::Green(param_value),
            "blue" => ControlWord::Blue(param_value),
            "cf" => ControlWord::ColorForeground(param_value),
            "cb" => ControlWord::ColorBackground(param),

            // Character formatting
            "b" => ControlWord::Bold(param_bool),
            "i" => ControlWord::Italic(param_bool),
            "ul" => ControlWord::Underline(param_bool),
            "ulnone" => ControlWord::UnderlineNone,
            "uldb" => ControlWord::UnderlineDouble,
            "uld" => ControlWord::UnderlineDotted,
            "uldash" => ControlWord::UnderlineDashed,
            "uldashd" => ControlWord::UnderlineDashDot,
            "uldashdd" => ControlWord::UnderlineDashDotDot,
            "ulw" => ControlWord::UnderlineWords,
            "ulth" => ControlWord::UnderlineThick,
            "ulwave" => ControlWord::UnderlineWave,
            "ulhair" => ControlWord::UnderlineHairline,
            "ulthd" => ControlWord::UnderlineThickDotted,
            "ulthdash" => ControlWord::UnderlineThickDashed,
            "ulthdashd" => ControlWord::UnderlineThickDashDot,
            "ulthdashdd" => ControlWord::UnderlineThickDashDotDot,
            "ulthldash" => ControlWord::UnderlineThickLongDash,
            "ulldash" => ControlWord::UnderlineLongDash,
            "ulhwave" => ControlWord::UnderlineHeavyWave,
            "ululdbwave" => ControlWord::UnderlineDoubleWave,
            "ulc" => ControlWord::UnderlineColor(param_value),
            "strike" => ControlWord::Strike(param_bool),
            "striked" => ControlWord::DoubleStrike(param_bool),
            "super" => ControlWord::Superscript(param_bool),
            "sub" => ControlWord::Subscript(param_bool),
            "nosupersub" => ControlWord::NoSuperSub,
            "up" => ControlWord::BaselineUp(param.unwrap_or(6)),
            "dn" => ControlWord::BaselineDown(param.unwrap_or(6)),
            "scaps" => ControlWord::SmallCaps(param_bool),
            "caps" => ControlWord::AllCaps(param_bool),
            "v" => ControlWord::Hidden(param_bool),
            "outl" => ControlWord::Outline(param_bool),
            "shad" => ControlWord::Shadow(param_bool),
            "embo" => ControlWord::Emboss(param_bool),
            "impr" => ControlWord::Imprint(param_bool),
            "expnd" => ControlWord::CharSpacing(param_value),
            "expndtw" => ControlWord::CharSpacingTwips(param_value),
            "charscalex" => ControlWord::CharScale(param_value),
            "kerning" => ControlWord::Kerning(param_value),
            "highlight" => ControlWord::Highlight(param_value),
            "chbrdr" => ControlWord::CharacterBorder(param),
            "chshdng" => ControlWord::CharacterShading(param),
            "chcfpat" => ControlWord::CharacterForegroundPattern(param),
            "chcbpat" => ControlWord::CharacterBackgroundPattern(param),
            "plain" => ControlWord::Plain,
            "loch" => ControlWord::LowAnsiCharacter(param),
            "hich" => ControlWord::HighAnsiCharacter(param),
            "dbch" => ControlWord::DoubleByteCharacter(param),
            "fcs" => ControlWord::FontComplexScript(param),
            "cgrid" => ControlWord::CharacterGrid(param),
            "animtext" => ControlWord::AnimatedText(param),
            "accnone" => ControlWord::EmphasisMark(crate::EmphasisMark::None, param),
            "accdot" => ControlWord::EmphasisMark(crate::EmphasisMark::Dot, param),
            "acccomma" => ControlWord::EmphasisMark(crate::EmphasisMark::Comma, param),
            "accunderdot" => ControlWord::EmphasisMark(crate::EmphasisMark::UnderDot, param),
            "acccircle" => ControlWord::EmphasisMark(crate::EmphasisMark::Circle, param),

            // Paragraph
            "par" => ControlWord::Par,
            "page" => ControlWord::Page(param),
            "column" => ControlWord::Column(param),
            "pard" => ControlWord::Pard,
            "ql" => ControlWord::LeftAlign,
            "qr" => ControlWord::RightAlign,
            "qc" => ControlWord::Center,
            "qj" => ControlWord::Justify,

            // Paragraph spacing/indent
            "sb" => ControlWord::SpaceBefore(param_value),
            "sa" => ControlWord::SpaceAfter(param_value),
            "sl" => ControlWord::SpaceBetween(param_value),
            "slmult" => ControlWord::LineMultiple(param_bool),
            "sbauto" => ControlWord::SpaceBeforeAuto(param),
            "saauto" => ControlWord::SpaceAfterAuto(param),
            "lisb" => ControlWord::ListSpaceBefore(param),
            "lisa" => ControlWord::ListSpaceAfter(param),
            "nosnaplinegrid" => ControlWord::NoSnapLineGrid(param),
            "contextualspace" => ControlWord::ContextualSpacing(param),
            "li" => ControlWord::LeftIndent(param_value),
            "ri" => ControlWord::RightIndent(param_value),
            "fi" => ControlWord::FirstLineIndent(param_value),
            "lin" => ControlWord::LogicalLeftIndent(param),
            "rin" => ControlWord::LogicalRightIndent(param),
            "cufi" => ControlWord::CharacterFirstLineIndent(param),
            "culi" => ControlWord::CharacterLeftIndent(param),
            "curi" => ControlWord::CharacterRightIndent(param),
            "indmirror" => ControlWord::MirrorIndents(param),

            // Paragraph additional properties
            "keep" => ControlWord::KeepTogether,
            "keepn" => ControlWord::KeepNext,
            "sbys" => ControlWord::SideBySide(param_bool),
            "pagebb" => ControlWord::PageBreakBefore,
            "widctlpar" => ControlWord::WidowControl,
            "nowidctlpar" => ControlWord::NoWidowControl(param),
            "dropcapli" => ControlWord::DropCapLines(param),
            "dropcapt" => ControlWord::DropCapType(param),
            "hyphpar" => ControlWord::ParagraphHyphenation(param),
            "aspalpha" => ControlWord::AutoSpaceAlphabetic(param),
            "aspnum" => ControlWord::AutoSpaceNumbers(param),
            "adjustright" => ControlWord::AdjustRightIndent(param),
            "wrapdefault" => ControlWord::WrapDefault(param),
            "nocwrap" => ControlWord::NoCharacterWrap(param),
            "nowwrap" => ControlWord::NoWordWrap(param),
            "nooverflow" => ControlWord::NoOverflow(param),
            "faauto" => ControlWord::FontAlignAuto(param),
            "fahang" => ControlWord::FontAlignHanging(param),
            "facenter" => ControlWord::FontAlignCenter(param),
            "faroman" => ControlWord::FontAlignRoman(param),
            "favar" => ControlWord::FontAlignVariable(param),
            "fafixed" => ControlWord::FontAlignFixed(param),

            // Tables
            "trowd" => ControlWord::TableRowDefaults,
            "irow" => ControlWord::TableRowIndex(param),
            "irowband" => ControlWord::TableRowBandIndex(param),
            "lastrow" => ControlWord::TableLastRow(param),
            "tbllkborder" => {
                ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::Border, param)
            },
            "tbllkshading" => {
                ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::Shading, param)
            },
            "tbllkfont" => {
                ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::Font, param)
            },
            "tbllkcolor" => {
                ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::Color, param)
            },
            "tbllkbestfit" => {
                ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::BestFit, param)
            },
            "tbllkhdrrows" => {
                ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::HeaderRows, param)
            },
            "tbllklastrow" => {
                ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::LastRow, param)
            },
            "tbllkhdrcols" => {
                ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::HeaderColumns, param)
            },
            "tbllklastcol" => {
                ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::LastColumn, param)
            },
            "tbllknorowband" => {
                ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::NoRowBanding, param)
            },
            "tbllknocolband" => {
                ControlWord::TableAutoformatFlag(crate::TableAutoformatFlag::NoColumnBanding, param)
            },
            "row" => ControlWord::TableRow,
            "cell" => ControlWord::TableCell,
            "cellx" => ControlWord::CellX(param_value),
            "intbl" => ControlWord::InTable,
            "trgaph" => ControlWord::TableRowGap(param),
            "trleft" => ControlWord::TableRowLeft(param),
            "trrh" => ControlWord::TableRowHeight(param),
            "trftsWidth" => {
                ControlWord::TablePreferredWidthUnit(crate::TableDistanceScope::Row, param)
            },
            "trwWidth" => {
                ControlWord::TablePreferredWidthValue(crate::TableDistanceScope::Row, param)
            },
            "trftsWidthB" => ControlWord::TableInvisibleWidthUnit(false, param),
            "trwWidthB" => ControlWord::TableInvisibleWidthValue(false, param),
            "trftsWidthA" => ControlWord::TableInvisibleWidthUnit(true, param),
            "trwWidthA" => ControlWord::TableInvisibleWidthValue(true, param),
            "clftsWidth" => {
                ControlWord::TablePreferredWidthUnit(crate::TableDistanceScope::Cell, param)
            },
            "clwWidth" => {
                ControlWord::TablePreferredWidthValue(crate::TableDistanceScope::Cell, param)
            },
            "trautofit" => ControlWord::TableAutoFit(param),
            "tblind" => ControlWord::TableIndentValue(param),
            "tblindtype" => ControlWord::TableIndentUnit(param),
            "trpaddl" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Left,
                },
                param,
            ),
            "trpaddr" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Right,
                },
                param,
            ),
            "trpaddt" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Top,
                },
                param,
            ),
            "trpaddb" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Bottom,
                },
                param,
            ),
            "trpaddfl" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Left,
                },
                param,
            ),
            "trpaddfr" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Right,
                },
                param,
            ),
            "trpaddft" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Top,
                },
                param,
            ),
            "trpaddfb" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Bottom,
                },
                param,
            ),
            "trspdl" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Left,
                },
                param,
            ),
            "trspdr" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Right,
                },
                param,
            ),
            "trspdt" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Top,
                },
                param,
            ),
            "trspdb" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Bottom,
                },
                param,
            ),
            "trspdfl" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Left,
                },
                param,
            ),
            "trspdfr" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Right,
                },
                param,
            ),
            "trspdft" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Top,
                },
                param,
            ),
            "trspdfb" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Row,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Bottom,
                },
                param,
            ),
            "clpadl" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Left,
                },
                param,
            ),
            "clpadr" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Right,
                },
                param,
            ),
            "clpadt" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Top,
                },
                param,
            ),
            "clpadb" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Bottom,
                },
                param,
            ),
            "clpadfl" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Left,
                },
                param,
            ),
            "clpadfr" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Right,
                },
                param,
            ),
            "clpadft" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Top,
                },
                param,
            ),
            "clpadfb" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Padding,
                    edge: crate::TableEdge::Bottom,
                },
                param,
            ),
            "clspdl" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Left,
                },
                param,
            ),
            "clspdr" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Right,
                },
                param,
            ),
            "clspdt" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Top,
                },
                param,
            ),
            "clspdb" => ControlWord::TableDistanceValue(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Bottom,
                },
                param,
            ),
            "clspdfl" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Left,
                },
                param,
            ),
            "clspdfr" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Right,
                },
                param,
            ),
            "clspdft" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Top,
                },
                param,
            ),
            "clspdfb" => ControlWord::TableDistanceUnit(
                crate::TableDistanceTarget {
                    scope: crate::TableDistanceScope::Cell,
                    kind: crate::TableDistanceKind::Spacing,
                    edge: crate::TableEdge::Bottom,
                },
                param,
            ),
            "tphcol" => ControlWord::TableHorizontalReference(
                crate::TableHorizontalReference::Column,
                param,
            ),
            "tphmrg" => ControlWord::TableHorizontalReference(
                crate::TableHorizontalReference::Margin,
                param,
            ),
            "tphpg" => {
                ControlWord::TableHorizontalReference(crate::TableHorizontalReference::Page, param)
            },
            "tpvmrg" => {
                ControlWord::TableVerticalReference(crate::TableVerticalReference::Margin, param)
            },
            "tpvpara" => {
                ControlWord::TableVerticalReference(crate::TableVerticalReference::Paragraph, param)
            },
            "tpvpg" => {
                ControlWord::TableVerticalReference(crate::TableVerticalReference::Page, param)
            },
            "tposx" => ControlWord::TableHorizontalOffset(false, param),
            "tposnegx" => ControlWord::TableHorizontalOffset(true, param),
            "tposxc" => {
                ControlWord::TableHorizontalPosition(crate::TableHorizontalPosition::Center, param)
            },
            "tposxi" => {
                ControlWord::TableHorizontalPosition(crate::TableHorizontalPosition::Inside, param)
            },
            "tposxl" => {
                ControlWord::TableHorizontalPosition(crate::TableHorizontalPosition::Left, param)
            },
            "tposxo" => {
                ControlWord::TableHorizontalPosition(crate::TableHorizontalPosition::Outside, param)
            },
            "tposxr" => {
                ControlWord::TableHorizontalPosition(crate::TableHorizontalPosition::Right, param)
            },
            "tposy" => ControlWord::TableVerticalOffset(false, param),
            "tposnegy" => ControlWord::TableVerticalOffset(true, param),
            "tposyb" => {
                ControlWord::TableVerticalPosition(crate::TableVerticalPosition::Bottom, param)
            },
            "tposyc" => {
                ControlWord::TableVerticalPosition(crate::TableVerticalPosition::Center, param)
            },
            "tposyil" => {
                ControlWord::TableVerticalPosition(crate::TableVerticalPosition::Inline, param)
            },
            "tposyin" => {
                ControlWord::TableVerticalPosition(crate::TableVerticalPosition::Inside, param)
            },
            "tposyout" => {
                ControlWord::TableVerticalPosition(crate::TableVerticalPosition::Outside, param)
            },
            "tposyt" => {
                ControlWord::TableVerticalPosition(crate::TableVerticalPosition::Top, param)
            },
            "tdfrmtxtLeft" => ControlWord::TableWrapDistance(crate::TableEdge::Left, param),
            "tdfrmtxtRight" => ControlWord::TableWrapDistance(crate::TableEdge::Right, param),
            "tdfrmtxtTop" => ControlWord::TableWrapDistance(crate::TableEdge::Top, param),
            "tdfrmtxtBottom" => ControlWord::TableWrapDistance(crate::TableEdge::Bottom, param),
            "tabsnoovrlp" => ControlWord::TableNoOverlap(param),
            "trhdr" => ControlWord::TableRowHeader(param),
            "trkeep" => ControlWord::TableRowKeep(param),
            "trkeepfollow" => ControlWord::TableRowKeepFollow(param),
            "trql" => ControlWord::TableRowAlignment(crate::TableRowAlignment::Left, param),
            "trqc" => ControlWord::TableRowAlignment(crate::TableRowAlignment::Center, param),
            "trqr" => ControlWord::TableRowAlignment(crate::TableRowAlignment::Right, param),
            "clvertalt" => ControlWord::TableCellVerticalAlignment(
                crate::TableCellVerticalAlignment::Top,
                param,
            ),
            "clvertalc" => ControlWord::TableCellVerticalAlignment(
                crate::TableCellVerticalAlignment::Center,
                param,
            ),
            "clvertalb" => ControlWord::TableCellVerticalAlignment(
                crate::TableCellVerticalAlignment::Bottom,
                param,
            ),
            "cltxlrtb" => ControlWord::TableCellTextFlow(
                crate::TableCellTextFlow::LeftToRightTopToBottom,
                param,
            ),
            "cltxtbrl" => ControlWord::TableCellTextFlow(
                crate::TableCellTextFlow::RightToLeftTopToBottom,
                param,
            ),
            "cltxbtlr" => ControlWord::TableCellTextFlow(
                crate::TableCellTextFlow::LeftToRightBottomToTop,
                param,
            ),
            "cltxlrtbv" => ControlWord::TableCellTextFlow(
                crate::TableCellTextFlow::LeftToRightTopToBottomVertical,
                param,
            ),
            "cltxtbrlv" => ControlWord::TableCellTextFlow(
                crate::TableCellTextFlow::TopToBottomRightToLeftVertical,
                param,
            ),
            "clFitText" => ControlWord::TableCellFitText(param),
            "clNoWrap" => ControlWord::TableCellNoWrap(param),
            "clhidemark" => ControlWord::TableCellHideMark(param),
            "clmgf" => ControlWord::TableCellMerge(
                crate::TableCellMergeAxis::Horizontal,
                crate::TableCellMergeRole::First,
                param,
            ),
            "clmrg" => ControlWord::TableCellMerge(
                crate::TableCellMergeAxis::Horizontal,
                crate::TableCellMergeRole::Continuation,
                param,
            ),
            "clvmgf" => ControlWord::TableCellMerge(
                crate::TableCellMergeAxis::Vertical,
                crate::TableCellMergeRole::First,
                param,
            ),
            "clvmrg" => ControlWord::TableCellMerge(
                crate::TableCellMergeAxis::Vertical,
                crate::TableCellMergeRole::Continuation,
                param,
            ),
            "trbrdrt" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Top),
                param,
            ),
            "trbrdrl" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Left),
                param,
            ),
            "trbrdrb" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Bottom),
                param,
            ),
            "trbrdrr" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Right),
                param,
            ),
            "trbrdrh" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Horizontal),
                param,
            ),
            "trbrdrv" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::Row(crate::TableRowBorderSide::Vertical),
                param,
            ),
            "clbrdrt" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::Cell(crate::TableCellBorderSide::Top),
                param,
            ),
            "clbrdrl" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::Cell(crate::TableCellBorderSide::Left),
                param,
            ),
            "clbrdrb" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::Cell(crate::TableCellBorderSide::Bottom),
                param,
            ),
            "clbrdrr" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::Cell(crate::TableCellBorderSide::Right),
                param,
            ),
            "cldglu" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::Cell(
                    crate::TableCellBorderSide::UpperLeftToLowerRight,
                ),
                param,
            ),
            "cldgll" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::Cell(
                    crate::TableCellBorderSide::UpperRightToLowerLeft,
                ),
                param,
            ),
            "tsbrdrt" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::StyleDefault(crate::TableStyleBorderSide::Top),
                param,
            ),
            "tsbrdrl" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::StyleDefault(crate::TableStyleBorderSide::Left),
                param,
            ),
            "tsbrdrb" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::StyleDefault(crate::TableStyleBorderSide::Bottom),
                param,
            ),
            "tsbrdrr" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::StyleDefault(crate::TableStyleBorderSide::Right),
                param,
            ),
            "tsbrdrh" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::StyleDefault(
                    crate::TableStyleBorderSide::HorizontalInside,
                ),
                param,
            ),
            "tsbrdrv" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::StyleDefault(
                    crate::TableStyleBorderSide::VerticalInside,
                ),
                param,
            ),
            "tsbrdrdgl" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::StyleDefault(
                    crate::TableStyleBorderSide::DiagonalUpperLeftToLowerRight,
                ),
                param,
            ),
            "tsbrdrdg" => ControlWord::TableBorder(
                crate::table::TableBorderTarget::StyleDefault(
                    crate::TableStyleBorderSide::DiagonalUpperRightToLowerLeft,
                ),
                param,
            ),
            "tscellpaddl" => ControlWord::TableDefaultDistanceValue(
                crate::TableDistanceKind::Padding,
                crate::TableEdge::Left,
                param,
            ),
            "tscellpaddr" => ControlWord::TableDefaultDistanceValue(
                crate::TableDistanceKind::Padding,
                crate::TableEdge::Right,
                param,
            ),
            "tscellpaddt" => ControlWord::TableDefaultDistanceValue(
                crate::TableDistanceKind::Padding,
                crate::TableEdge::Top,
                param,
            ),
            "tscellpaddb" => ControlWord::TableDefaultDistanceValue(
                crate::TableDistanceKind::Padding,
                crate::TableEdge::Bottom,
                param,
            ),
            "tscellpaddfl" => ControlWord::TableDefaultDistanceUnit(
                crate::TableDistanceKind::Padding,
                crate::TableEdge::Left,
                param,
            ),
            "tscellpaddfr" => ControlWord::TableDefaultDistanceUnit(
                crate::TableDistanceKind::Padding,
                crate::TableEdge::Right,
                param,
            ),
            "tscellpaddft" => ControlWord::TableDefaultDistanceUnit(
                crate::TableDistanceKind::Padding,
                crate::TableEdge::Top,
                param,
            ),
            "tscellpaddfb" => ControlWord::TableDefaultDistanceUnit(
                crate::TableDistanceKind::Padding,
                crate::TableEdge::Bottom,
                param,
            ),
            "tscellspcl" => ControlWord::TableDefaultDistanceValue(
                crate::TableDistanceKind::Spacing,
                crate::TableEdge::Left,
                param,
            ),
            "tscellspcr" => ControlWord::TableDefaultDistanceValue(
                crate::TableDistanceKind::Spacing,
                crate::TableEdge::Right,
                param,
            ),
            "tscellspct" => ControlWord::TableDefaultDistanceValue(
                crate::TableDistanceKind::Spacing,
                crate::TableEdge::Top,
                param,
            ),
            "tscellspcb" => ControlWord::TableDefaultDistanceValue(
                crate::TableDistanceKind::Spacing,
                crate::TableEdge::Bottom,
                param,
            ),
            "tscellspcfl" => ControlWord::TableDefaultDistanceUnit(
                crate::TableDistanceKind::Spacing,
                crate::TableEdge::Left,
                param,
            ),
            "tscellspcfr" => ControlWord::TableDefaultDistanceUnit(
                crate::TableDistanceKind::Spacing,
                crate::TableEdge::Right,
                param,
            ),
            "tscellspcft" => ControlWord::TableDefaultDistanceUnit(
                crate::TableDistanceKind::Spacing,
                crate::TableEdge::Top,
                param,
            ),
            "tscellspcfb" => ControlWord::TableDefaultDistanceUnit(
                crate::TableDistanceKind::Spacing,
                crate::TableEdge::Bottom,
                param,
            ),
            "tscellwidthfts" => ControlWord::TableDefaultCellWidthUnit(param),
            "tscellwidth" => ControlWord::TableDefaultCellWidthValue(param),
            "trshdng" => ControlWord::TableShadingAmount(crate::TableDistanceScope::Row, param),
            "trcfpat" => ControlWord::TableShadingForeground(crate::TableDistanceScope::Row, param),
            "trcbpat" => ControlWord::TableShadingBackground(crate::TableDistanceScope::Row, param),
            "trpat" => ControlWord::TableRowShadingPatternIndex(param),
            "clshdng" => ControlWord::TableShadingAmount(crate::TableDistanceScope::Cell, param),
            "clcfpat" => {
                ControlWord::TableShadingForeground(crate::TableDistanceScope::Cell, param)
            },
            "clcbpat" => {
                ControlWord::TableShadingBackground(crate::TableDistanceScope::Cell, param)
            },
            "trbghoriz" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Row,
                crate::ShadingPattern::Horizontal,
                param,
            ),
            "trbgvert" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Row,
                crate::ShadingPattern::Vertical,
                param,
            ),
            "trbgfdiag" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Row,
                crate::ShadingPattern::ForwardDiagonal,
                param,
            ),
            "trbgbdiag" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Row,
                crate::ShadingPattern::BackwardDiagonal,
                param,
            ),
            "trbgcross" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Row,
                crate::ShadingPattern::Cross,
                param,
            ),
            "trbgdcross" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Row,
                crate::ShadingPattern::DiagonalCross,
                param,
            ),
            "trbgdkhor" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Row,
                crate::ShadingPattern::DarkHorizontal,
                param,
            ),
            "trbgdkvert" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Row,
                crate::ShadingPattern::DarkVertical,
                param,
            ),
            "trbgdkfdiag" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Row,
                crate::ShadingPattern::DarkForwardDiagonal,
                param,
            ),
            "trbgdkbdiag" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Row,
                crate::ShadingPattern::DarkBackwardDiagonal,
                param,
            ),
            "trbgdkcross" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Row,
                crate::ShadingPattern::DarkCross,
                param,
            ),
            "trbgdkdcross" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Row,
                crate::ShadingPattern::DarkDiagonalCross,
                param,
            ),
            "clbghoriz" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Cell,
                crate::ShadingPattern::Horizontal,
                param,
            ),
            "clbgvert" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Cell,
                crate::ShadingPattern::Vertical,
                param,
            ),
            "clbgfdiag" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Cell,
                crate::ShadingPattern::ForwardDiagonal,
                param,
            ),
            "clbgbdiag" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Cell,
                crate::ShadingPattern::BackwardDiagonal,
                param,
            ),
            "clbgcross" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Cell,
                crate::ShadingPattern::Cross,
                param,
            ),
            "clbgdcross" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Cell,
                crate::ShadingPattern::DiagonalCross,
                param,
            ),
            "clbgdkhor" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Cell,
                crate::ShadingPattern::DarkHorizontal,
                param,
            ),
            "clbgdkvert" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Cell,
                crate::ShadingPattern::DarkVertical,
                param,
            ),
            "clbgdkfdiag" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Cell,
                crate::ShadingPattern::DarkForwardDiagonal,
                param,
            ),
            "clbgdkbdiag" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Cell,
                crate::ShadingPattern::DarkBackwardDiagonal,
                param,
            ),
            "clbgdkcross" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Cell,
                crate::ShadingPattern::DarkCross,
                param,
            ),
            "clbgdkdcross" => ControlWord::TableShadingPattern(
                crate::TableDistanceScope::Cell,
                crate::ShadingPattern::DarkDiagonalCross,
                param,
            ),
            "nestcell" => ControlWord::NestedTableCell(param),
            "nestrow" => ControlWord::NestedTableRow(param),
            "nesttableprops" => ControlWord::NestedTableProperties(param),
            "nonesttables" => ControlWord::NoNestedTables(param),

            // Borders
            "brdrt" => ControlWord::BorderTop,
            "brdrb" => ControlWord::BorderBottom,
            "brdrl" => ControlWord::BorderLeft,
            "brdrr" => ControlWord::BorderRight,
            "brdrs" => ControlWord::BorderSingle,
            "brdrdot" => ControlWord::BorderDotted,
            "brdrdash" => ControlWord::BorderDashed,
            "brdrdb" => ControlWord::BorderDouble,
            "brdrtriple" => ControlWord::BorderTriple,
            "brdrwavy" => ControlWord::BorderWave,
            "brdrnone" | "brdrnil" => ControlWord::BorderNone,
            "brdrth" => ControlWord::BorderThick,
            "brdrdashsm" => ControlWord::BorderDashSmall,
            "brdrdashd" => ControlWord::BorderDotDash,
            "brdrdashdd" => ControlWord::BorderDotDotDash,
            "brdrtnthsg" => ControlWord::BorderThinThickSmall,
            "brdrthtnsg" => ControlWord::BorderThickThinSmall,
            "brdrtnthtnsg" => ControlWord::BorderThinThickThinSmall,
            "brdrtnthmg" => ControlWord::BorderThinThickMedium,
            "brdrthtnmg" => ControlWord::BorderThickThinMedium,
            "brdrtnthtnmg" => ControlWord::BorderThinThickThinMedium,
            "brdrtnthlg" => ControlWord::BorderThinThickLarge,
            "brdrthtnlg" => ControlWord::BorderThickThinLarge,
            "brdrtnthtnlg" => ControlWord::BorderThinThickThinLarge,
            "brdrwavydb" => ControlWord::BorderWavyDouble,
            "brdrdashdotstr" => ControlWord::BorderStriped,
            "brdremboss" => ControlWord::BorderEmbossed,
            "brdrengrave" => ControlWord::BorderEngraved,
            "brdroutset" => ControlWord::BorderOutset,
            "brdrinset" => ControlWord::BorderInset,
            "brdrsh" => ControlWord::BorderShadow,
            "brdrframe" => ControlWord::BorderFrame,
            "brdrw" => ControlWord::BorderWidth(param),
            "brdrcf" => ControlWord::BorderColor(param),
            "brsp" => ControlWord::BorderSpace(param),
            "pgbrdrt" => ControlWord::PageBorderTop,
            "pgbrdrl" => ControlWord::PageBorderLeft,
            "pgbrdrb" => ControlWord::PageBorderBottom,
            "pgbrdrr" => ControlWord::PageBorderRight,
            "pgbrdropt" => ControlWord::PageBorderOptions(param),
            "pgbrdrhead" => ControlWord::PageBorderSurroundHeader,
            "pgbrdrfoot" => ControlWord::PageBorderSurroundFooter,
            "pgbrdrsnap" => ControlWord::PageBorderSnap,
            "brdrart" => ControlWord::PageBorderArt(param),

            // Shading
            "shading" => ControlWord::Shading(param),
            "cfpat" => ControlWord::ForegroundPattern(param),
            "cbpat" => ControlWord::BackgroundPattern(param),

            // Tab stops
            "tql" => ControlWord::TabLeft(param),
            "tqr" => ControlWord::TabRight(param),
            "tqc" => ControlWord::TabCenter(param),
            "tqdec" => ControlWord::TabDecimal(param),
            "tb" => ControlWord::TabBar(param),
            "tx" => ControlWord::TabPosition(param),
            "tldot" => ControlWord::TabLeaderDot(param),
            "tlmdot" => ControlWord::TabLeaderMiddleDot(param),
            "tlhyph" => ControlWord::TabLeaderHyphen(param),
            "tlul" => ControlWord::TabLeaderUnderscore(param),
            "tlth" => ControlWord::TabLeaderThick(param),
            "tleq" => ControlWord::TabLeaderEqual(param),

            // Lists
            "listtable" => ControlWord::ListTable,
            "list" => ControlWord::List,
            "listtemplateid" => ControlWord::ListTemplateId(param_value),
            "listsimple" => ControlWord::ListSimple(param_bool),
            "listhybrid" => ControlWord::ListHybrid(param_bool),
            "listname" => ControlWord::ListName,
            "liststylename" => ControlWord::ListStyleName,
            "listtext" => ControlWord::GeneratedListText,
            "listpicture" | "listpict" => ControlWord::ListPicture(param),
            "shppict" => ControlWord::ShapePicture(param),
            "nonshppict" => ControlWord::NonShapePicture(param),
            "listid" => ControlWord::ListId(param_value),
            "listoverridetable" => ControlWord::ListOverrideTable,
            "listoverride" => ControlWord::ListOverride,
            "listoverridecount" => ControlWord::ListOverrideCount(param_value),
            "ls" => ControlWord::ListOverrideIndex(param_value),
            "ilvl" => ControlWord::ListLevelIndex(param_value),
            "lfolevel" => ControlWord::ListOverrideLevel,
            "listoverridestartat" => ControlWord::ListOverrideStartAt(param_bool),
            "listoverrideformat" => ControlWord::ListOverrideFormat(param_bool),
            "listlevel" => ControlWord::ListLevel,
            "levelnfc" | "levelnfcn" => ControlWord::ListLevelType(param_value),
            "leveljc" | "leveljcn" => ControlWord::ListLevelJustification(param_value),
            "levelfollow" => ControlWord::ListLevelFollow(param_value),
            "levelspace" => ControlWord::ListLevelSpace(param_value),
            "levelindent" => ControlWord::ListLevelIndent(param_value),
            "leveltext" => ControlWord::ListNumberText,
            "levelstartat" => ControlWord::ListLevelStartAt(param_value),
            "levelnumbers" => ControlWord::ListLevelNumbers,
            "levelpicture" => ControlWord::ListLevelPicture(param_value),
            "lvltentative" => ControlWord::ListLevelTentative,
            "levellegal" => ControlWord::ListLevelLegal(param_bool),
            "levelnorestart" => ControlWord::ListLevelNoRestart(param_bool),
            "levelold" => ControlWord::ListLevelOld(param_bool),
            "levelprev" => ControlWord::ListLevelPrevious(param_bool),
            "levelprevspace" => ControlWord::ListLevelPreviousSpace(param_bool),
            "leveltemplateid" => ControlWord::ListLevelTemplateId(param_value),

            // Sections
            "titlepg" => ControlWord::TitlePage,
            "endnhere" => ControlWord::SectionEndnoteHere,
            "outlinelevel" => ControlWord::OutlineLevel(param_value),
            "sectd" => ControlWord::SectionDefault,
            "sect" => ControlWord::Section,
            "sbk" => ControlWord::SectionBreak,
            "sbknone" => ControlWord::SectionContinuous,
            "sbkcol" => ControlWord::SectionColumn,
            "sbkpage" => ControlWord::SectionPage,
            "sbkeven" => ControlWord::SectionEvenPage,
            "sbkodd" => ControlWord::SectionOddPage,
            "paperw" | "pgwsxn" => ControlWord::PageWidth(param_value),
            "paperh" | "pghsxn" => ControlWord::PageHeight(param_value),
            "margl" | "marglsxn" => ControlWord::MarginLeft(param_value),
            "margr" | "margrsxn" => ControlWord::MarginRight(param_value),
            "margt" | "margtsxn" => ControlWord::MarginTop(param_value),
            "margb" | "margbsxn" => ControlWord::MarginBottom(param_value),
            "guttersxn" => ControlWord::MarginGutter(param_value),
            "binfsxn" => ControlWord::PaperSourceFirst(param),
            "binsxn" => ControlWord::PaperSourceOther(param),
            "facingp" => ControlWord::FacingPages(param_bool),
            "margmirror" => ControlWord::MirrorMargins(param),
            "gutter" => ControlWord::DocumentGutter(param),
            "headery" => ControlWord::HeaderDistance(param_value),
            "footery" => ControlWord::FooterDistance(param_value),
            "landscape" | "lndscpsxn" => ControlWord::Landscape,
            "cols" => ControlWord::Columns(param),
            "colsx" => ControlWord::ColumnSpace(param),
            "colno" => ControlWord::ColumnNumber(param),
            "colw" => ControlWord::ColumnWidth(param),
            "colsr" => ControlWord::ColumnSpaceRight(param),
            "linebetcol" => ControlWord::ColumnSeparator(param_bool),
            "pgnstarts" => ControlWord::PageNumberStart(param_value),
            "pgndec" => ControlWord::PageNumberFormat(crate::PageNumberFormat::Decimal),
            "pgnucrm" => ControlWord::PageNumberFormat(crate::PageNumberFormat::UpperRoman),
            "pgnlcrm" => ControlWord::PageNumberFormat(crate::PageNumberFormat::LowerRoman),
            "pgnucltr" => ControlWord::PageNumberFormat(crate::PageNumberFormat::UpperLetter),
            "pgnlcltr" => ControlWord::PageNumberFormat(crate::PageNumberFormat::LowerLetter),
            "pgnbidia" => ControlWord::PageNumberFormat(crate::PageNumberFormat::BidiAlphabetic),
            "pgnbidib" => ControlWord::PageNumberFormat(crate::PageNumberFormat::BidiAbjad),
            "pgnchosung" => ControlWord::PageNumberFormat(crate::PageNumberFormat::KoreanChosung),
            "pgncnum" => ControlWord::PageNumberFormat(crate::PageNumberFormat::Circle),
            "pgndbnum" => ControlWord::PageNumberFormat(crate::PageNumberFormat::KanjiDigitless),
            "pgndbnumd" => ControlWord::PageNumberFormat(crate::PageNumberFormat::KanjiWithDigit),
            "pgndbnumt" => ControlWord::PageNumberFormat(crate::PageNumberFormat::KanjiThree),
            "pgndbnumk" => ControlWord::PageNumberFormat(crate::PageNumberFormat::KanjiFour),
            "pgndecd" => ControlWord::PageNumberFormat(crate::PageNumberFormat::DoubleDecimal),
            "pgnganada" => ControlWord::PageNumberFormat(crate::PageNumberFormat::KoreanGanada),
            "pgngbnum" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ChineseOne),
            "pgngbnumd" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ChineseTwo),
            "pgngbnuml" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ChineseThree),
            "pgngbnumk" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ChineseFour),
            "pgnhindia" => ControlWord::PageNumberFormat(crate::PageNumberFormat::HindiVowels),
            "pgnhindib" => ControlWord::PageNumberFormat(crate::PageNumberFormat::HindiConsonants),
            "pgnhindic" => ControlWord::PageNumberFormat(crate::PageNumberFormat::HindiNumbers),
            "pgnhindid" => ControlWord::PageNumberFormat(crate::PageNumberFormat::HindiDescriptive),
            "pgnthaia" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ThaiLetters),
            "pgnthaib" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ThaiNumbers),
            "pgnthaic" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ThaiDescriptive),
            "pgnvieta" => {
                ControlWord::PageNumberFormat(crate::PageNumberFormat::VietnameseCardinal)
            },
            "pgnzodiac" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ZodiacOne),
            "pgnzodiacd" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ZodiacTwo),
            "pgnzodiacl" => ControlWord::PageNumberFormat(crate::PageNumberFormat::ZodiacThree),
            "pgnrestart" => ControlWord::PageNumberRestart(crate::PageNumberRestart::Restart),
            "pgncont" => ControlWord::PageNumberRestart(crate::PageNumberRestart::Continuous),
            "pgnx" => ControlWord::PageNumberOffsetX(param_value),
            "pgny" => ControlWord::PageNumberOffsetY(param_value),
            "pgnhn" => ControlWord::PageNumberHeadingLevel(param),
            "pgnhnsh" => {
                ControlWord::PageNumberHeadingSeparator(crate::PageNumberHeadingSeparator::Hyphen)
            },
            "pgnhnsp" => {
                ControlWord::PageNumberHeadingSeparator(crate::PageNumberHeadingSeparator::Period)
            },
            "pgnhnsc" => {
                ControlWord::PageNumberHeadingSeparator(crate::PageNumberHeadingSeparator::Colon)
            },
            "pgnhnsm" => {
                ControlWord::PageNumberHeadingSeparator(crate::PageNumberHeadingSeparator::EmDash)
            },
            "pgnhnsn" => {
                ControlWord::PageNumberHeadingSeparator(crate::PageNumberHeadingSeparator::EnDash)
            },
            "sectlinegrid" => ControlWord::SectionLineGrid(param),
            "sectspecifyl" => {
                ControlWord::SectionDocumentGrid(crate::SectionDocumentGridType::LinesAndCharacters)
            },
            "sectspecifycl" => {
                ControlWord::SectionDocumentGrid(crate::SectionDocumentGridType::CharactersOnly)
            },
            "sectspecifygen" => {
                ControlWord::SectionDocumentGrid(crate::SectionDocumentGridType::Default)
            },
            "prauth" => ControlWord::ParagraphRevisionAuthor(param_value),
            "prdate" => ControlWord::ParagraphRevisionDate(param_value),
            "srauth" => ControlWord::SectionRevisionAuthor(param_value),
            "srdate" => ControlWord::SectionRevisionDate(param_value),
            "trauth" => ControlWord::TableRowRevisionAuthor(param_value),
            "trdate" => ControlWord::TableRowRevisionDate(param_value),
            "clins" => ControlWord::CellRevisionMark(crate::CellRevisionKind::Inserted),
            "clinsauth" => {
                ControlWord::CellRevisionAuthor(crate::CellRevisionKind::Inserted, param_value)
            },
            "clinsdttm" => {
                ControlWord::CellRevisionDate(crate::CellRevisionKind::Inserted, param_value)
            },
            "cldel" => ControlWord::CellRevisionMark(crate::CellRevisionKind::Deleted),
            "cldelauth" => {
                ControlWord::CellRevisionAuthor(crate::CellRevisionKind::Deleted, param_value)
            },
            "cldeldttm" => {
                ControlWord::CellRevisionDate(crate::CellRevisionKind::Deleted, param_value)
            },
            "clmrgd" => ControlWord::CellRevisionMark(crate::CellRevisionKind::MergeDeleted),
            "clmrgdauth" => {
                ControlWord::CellRevisionAuthor(crate::CellRevisionKind::MergeDeleted, param_value)
            },
            "clmrgddttm" => {
                ControlWord::CellRevisionDate(crate::CellRevisionKind::MergeDeleted, param_value)
            },
            "vertalt" => ControlWord::VerticalAlignTop,
            "vertalc" => ControlWord::VerticalAlignCenter,
            "vertalj" => ControlWord::VerticalAlignJustify,
            "vertalb" => ControlWord::VerticalAlignBottom,
            "linemod" => ControlWord::LineNumbering(param),
            "linex" => ControlWord::LineNumberDistance(param),
            "linestarts" => ControlWord::LineNumberStart(param),
            "linerestart" => ControlWord::LineNumberRestartSection,
            "lineppage" => ControlWord::LineNumberRestartPage,
            "linecont" => ControlWord::LineNumberContinuous,
            "header" => ControlWord::Header,
            "headerf" => ControlWord::HeaderFirst,
            "headerl" => ControlWord::HeaderLeft,
            "headerr" => ControlWord::HeaderRight,
            "footer" => ControlWord::Footer,
            "footerf" => ControlWord::FooterFirst,
            "footerl" => ControlWord::FooterLeft,
            "footerr" => ControlWord::FooterRight,

            // Footnotes and endnotes
            "footnote" => ControlWord::Footnote,
            "endnote" => ControlWord::Endnote,
            "chftn" => ControlWord::FootnoteNumber(param_value),
            "fet" => ControlWord::NoteKinds(param_value),
            "endnotes" => ControlWord::FootnotePlacement(crate::NotePlacement::EndOfSection),
            "enddoc" => ControlWord::FootnotePlacement(crate::NotePlacement::EndOfDocument),
            "ftntj" => ControlWord::FootnotePlacement(crate::NotePlacement::BeneathText),
            "ftnbj" => ControlWord::FootnotePlacement(crate::NotePlacement::BottomOfPage),
            "aendnotes" => ControlWord::EndnotePlacement(crate::NotePlacement::EndOfSection),
            "aenddoc" => ControlWord::EndnotePlacement(crate::NotePlacement::EndOfDocument),
            "aftntj" => ControlWord::EndnotePlacement(crate::NotePlacement::BeneathText),
            "aftnbj" => ControlWord::EndnotePlacement(crate::NotePlacement::BottomOfPage),
            "ftnstart" => ControlWord::FootnoteStart(param_value),
            "aftnstart" => ControlWord::EndnoteStart(param_value),
            "ftnrstcont" => ControlWord::FootnoteRestart(crate::FootnoteRestart::Continuous),
            "ftnrestart" => ControlWord::FootnoteRestart(crate::FootnoteRestart::EachSection),
            "ftnrstpg" => ControlWord::FootnoteRestart(crate::FootnoteRestart::EachPage),
            "aftnrstcont" => ControlWord::EndnoteRestart(crate::EndnoteRestart::Continuous),
            "aftnrestart" => ControlWord::EndnoteRestart(crate::EndnoteRestart::EachSection),
            "ftnnar" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::Arabic),
            "ftnnalc" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::LowercaseLetter),
            "ftnnauc" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::UppercaseLetter),
            "ftnnrlc" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::LowercaseRoman),
            "ftnnruc" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::UppercaseRoman),
            "ftnnchi" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::Chicago),
            "ftnnchosung" => {
                ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::KoreanChosung)
            },
            "ftnncnum" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::Circle),
            "ftnndbnum" => {
                ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::KanjiDigitless)
            },
            "ftnndbnumd" => {
                ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::KanjiWithDigit)
            },
            "ftnndbnumt" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::KanjiThree),
            "ftnndbnumk" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::KanjiFour),
            "ftnndbar" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::DoubleByte),
            "ftnnganada" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::KoreanGanada),
            "ftnngbnum" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::ChineseOne),
            "ftnngbnumd" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::ChineseTwo),
            "ftnngbnuml" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::ChineseThree),
            "ftnngbnumk" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::ChineseFour),
            "ftnnzodiac" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::ZodiacOne),
            "ftnnzodiacd" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::ZodiacTwo),
            "ftnnzodiacl" => ControlWord::FootnoteNumbering(crate::NoteNumberingStyle::ZodiacThree),
            "aftnnar" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::Arabic),
            "aftnnalc" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::LowercaseLetter),
            "aftnnauc" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::UppercaseLetter),
            "aftnnrlc" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::LowercaseRoman),
            "aftnnruc" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::UppercaseRoman),
            "aftnnchi" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::Chicago),
            "aftnnchosung" => {
                ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::KoreanChosung)
            },
            "aftnncnum" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::Circle),
            "aftnndbnum" => {
                ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::KanjiDigitless)
            },
            "aftnndbnumd" => {
                ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::KanjiWithDigit)
            },
            "aftnndbnumt" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::KanjiThree),
            "aftnndbnumk" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::KanjiFour),
            "aftnndbar" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::DoubleByte),
            "aftnnganada" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::KoreanGanada),
            "aftnngbnum" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::ChineseOne),
            "aftnngbnumd" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::ChineseTwo),
            "aftnngbnuml" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::ChineseThree),
            "aftnngbnumk" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::ChineseFour),
            "aftnnzodiac" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::ZodiacOne),
            "aftnnzodiacd" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::ZodiacTwo),
            "aftnnzodiacl" => ControlWord::EndnoteNumbering(crate::NoteNumberingStyle::ZodiacThree),
            "sftntj" => {
                ControlWord::SectionFootnotePlacement(crate::SectionFootnotePlacement::BeneathText)
            },
            "sftnbj" => {
                ControlWord::SectionFootnotePlacement(crate::SectionFootnotePlacement::BottomOfPage)
            },
            "sftnstart" => ControlWord::SectionFootnoteStart(param_value),
            "saftnstart" => ControlWord::SectionEndnoteStart(param_value),
            "sftnrstcont" => {
                ControlWord::SectionFootnoteRestart(crate::FootnoteRestart::Continuous)
            },
            "sftnrestart" => {
                ControlWord::SectionFootnoteRestart(crate::FootnoteRestart::EachSection)
            },
            "sftnrstpg" => ControlWord::SectionFootnoteRestart(crate::FootnoteRestart::EachPage),
            "saftnrstcont" => ControlWord::SectionEndnoteRestart(crate::EndnoteRestart::Continuous),
            "saftnrestart" => {
                ControlWord::SectionEndnoteRestart(crate::EndnoteRestart::EachSection)
            },
            "sftnnar" => ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::Arabic),
            "sftnnalc" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::LowercaseLetter)
            },
            "sftnnauc" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::UppercaseLetter)
            },
            "sftnnrlc" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::LowercaseRoman)
            },
            "sftnnruc" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::UppercaseRoman)
            },
            "sftnnchi" => ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::Chicago),
            "sftnnchosung" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::KoreanChosung)
            },
            "sftnncnum" => ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::Circle),
            "sftnndbnum" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::KanjiDigitless)
            },
            "sftnndbnumd" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::KanjiWithDigit)
            },
            "sftnndbnumt" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::KanjiThree)
            },
            "sftnndbnumk" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::KanjiFour)
            },
            "sftnndbar" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::DoubleByte)
            },
            "sftnnganada" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::KoreanGanada)
            },
            "sftnngbnum" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::ChineseOne)
            },
            "sftnngbnumd" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::ChineseTwo)
            },
            "sftnngbnuml" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::ChineseThree)
            },
            "sftnngbnumk" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::ChineseFour)
            },
            "sftnnzodiac" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::ZodiacOne)
            },
            "sftnnzodiacd" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::ZodiacTwo)
            },
            "sftnnzodiacl" => {
                ControlWord::SectionFootnoteNumbering(crate::NoteNumberingStyle::ZodiacThree)
            },
            "saftnnar" => ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::Arabic),
            "saftnnalc" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::LowercaseLetter)
            },
            "saftnnauc" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::UppercaseLetter)
            },
            "saftnnrlc" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::LowercaseRoman)
            },
            "saftnnruc" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::UppercaseRoman)
            },
            "saftnnchi" => ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::Chicago),
            "saftnnchosung" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::KoreanChosung)
            },
            "saftnncnum" => ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::Circle),
            "saftnndbnum" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::KanjiDigitless)
            },
            "saftnndbnumd" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::KanjiWithDigit)
            },
            "saftnndbnumt" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::KanjiThree)
            },
            "saftnndbnumk" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::KanjiFour)
            },
            "saftnndbar" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::DoubleByte)
            },
            "saftnnganada" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::KoreanGanada)
            },
            "saftnngbnum" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::ChineseOne)
            },
            "saftnngbnumd" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::ChineseTwo)
            },
            "saftnngbnuml" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::ChineseThree)
            },
            "saftnngbnumk" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::ChineseFour)
            },
            "saftnnzodiac" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::ZodiacOne)
            },
            "saftnnzodiacd" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::ZodiacTwo)
            },
            "saftnnzodiacl" => {
                ControlWord::SectionEndnoteNumbering(crate::NoteNumberingStyle::ZodiacThree)
            },
            "ftnsep" => ControlWord::FootnoteSeparator,
            "ftnsepc" => ControlWord::FootnoteContinuationSeparator,
            "ftncn" => ControlWord::FootnoteContinuationNotice,
            "aftnsep" => ControlWord::EndnoteSeparator,
            "aftnsepc" => ControlWord::EndnoteContinuationSeparator,
            "aftncn" => ControlWord::EndnoteContinuationNotice,
            "chftnsep" => ControlWord::NoteSeparatorCharacter,
            "chftnsepc" => ControlWord::NoteContinuationSeparatorCharacter,

            // Bookmarks
            "bkmkstart" => ControlWord::BookmarkStart,
            "bkmkend" => ControlWord::BookmarkEnd,
            "bkmkcolf" => ControlWord::BookmarkFirstColumn(param_value),
            "bkmkcoll" => ControlWord::BookmarkLastColumn(param_value),
            "bkmkpub" => ControlWord::BookmarkPublic,

            // Annotations
            "atn" | "annotation" => ControlWord::Annotation,
            "atndate" => ControlWord::AnnotationDate,
            "atnauthor" => ControlWord::AnnotationAuthor,
            "atnid" => ControlWord::AnnotationInitials,
            "atnref" => ControlWord::AnnotationReference,
            "atnparent" => ControlWord::AnnotationParent,
            "atrfstart" => ControlWord::AnnotationRangeStart,
            "atrfend" => ControlWord::AnnotationRangeEnd,
            "atnicn" => ControlWord::AnnotationIcon,
            "atntime" => ControlWord::AnnotationTime,
            "chatn" => ControlWord::AnnotationMark,

            // Tracked revisions
            "revtbl" => ControlWord::RevisionTable,
            "revised" => ControlWord::Revised(param_bool),
            "deleted" => ControlWord::Deleted(param_bool),
            "revauth" => ControlWord::RevisionAuthor(param_value),
            "revauthdel" => ControlWord::DeletedRevisionAuthor(param_value),
            "revdttm" => ControlWord::RevisionDate(param_value),
            "revdttmdel" => ControlWord::DeletedRevisionDate(param_value),
            "revprop" => ControlWord::RevisionProperty(param_value),

            // Shapes
            "shp" => ControlWord::Shape(param),
            "shpgrp" => ControlWord::ShapeGroup(param),
            "shpinst" => param.map_or(ControlWord::ShapeInstance, ControlWord::ShapeType),
            "shpleft" => ControlWord::ShapeLeft(param_value),
            "shptop" => ControlWord::ShapeTop(param_value),
            "shpright" => ControlWord::ShapeRight(param_value),
            "shpbottom" => ControlWord::ShapeBottom(param_value),
            "shpwidth" => ControlWord::ShapeWidth(param_value),
            "shpheight" => ControlWord::ShapeHeight(param_value),
            "shprotation" => ControlWord::ShapeRotation(param_value),
            "shpz" => ControlWord::ShapeZOrder(param_value),
            "shpwr" => ControlWord::ShapeWrap(param_value),
            "shpfblwtxt" => ControlWord::ShapeBelowText(param_bool),
            "shplockanchor" => ControlWord::ShapeLockAnchor,
            "sp" => {
                if param.is_none() {
                    ControlWord::ShapeProperty
                } else {
                    ControlWord::InvalidPictureShapePropertyParameter
                }
            },
            "sn" => {
                if param.is_none() {
                    ControlWord::ShapePropertyName
                } else {
                    ControlWord::InvalidPictureShapePropertyParameter
                }
            },
            "sv" => {
                if param.is_none() {
                    ControlWord::ShapePropertyValue
                } else {
                    ControlWord::InvalidPictureShapePropertyParameter
                }
            },
            "svb" => ControlWord::ShapeBinaryValue(param),
            "hsv" => ControlWord::ShapeThemeValue(param),
            "caccentone" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Accent1, param),
            "caccenttwo" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Accent2, param),
            "caccentthree" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Accent3, param),
            "caccentfour" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Accent4, param),
            "caccentfive" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Accent5, param),
            "caccentsix" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Accent6, param),
            "cbackgroundone" => {
                ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Background1, param)
            },
            "cbackgroundtwo" => {
                ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Background2, param)
            },
            "ctextone" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Text1, param),
            "ctexttwo" => ControlWord::ShapeThemeColor(crate::ShapeThemeColor::Text2, param),
            "ctint" => ControlWord::ShapeThemeTint(param),
            "cshade" => ControlWord::ShapeThemeShade(param),
            "shprslt" => ControlWord::ShapeResult(param),
            "shptxt" => ControlWord::ShapeText(param),
            "do" => ControlWord::LegacyDrawingObject,
            "dptxbx" => ControlWord::LegacyTextBox,
            "dptxbxtext" => ControlWord::LegacyTextBoxText,
            "dobxpage" => ControlWord::LegacyAnchorXPage,
            "dobxmargin" => ControlWord::LegacyAnchorXMargin,
            "dobxcolumn" => ControlWord::LegacyAnchorXColumn,
            "dobypage" => ControlWord::LegacyAnchorYPage,
            "dobymargin" => ControlWord::LegacyAnchorYMargin,
            "dobypara" => ControlWord::LegacyAnchorYParagraph,
            "dodhgt" | "dolhgt" => ControlWord::LegacyDrawingHeight(param_value),
            "dptxbxmar" => ControlWord::LegacyTextBoxMargin(param_value),
            "dpx" => ControlWord::LegacyDrawingX(param_value),
            "dpy" => ControlWord::LegacyDrawingY(param_value),
            "dpxsize" => ControlWord::LegacyDrawingWidth(param_value),
            "dpysize" => ControlWord::LegacyDrawingHeightSize(param_value),
            "dptxlrtb" => ControlWord::LegacyTextLeftRightTopBottom,
            "dptxlrtbv" => ControlWord::LegacyTextLeftRightTopBottomVertical,
            "dptxtbrl" => ControlWord::LegacyTextTopBottomRightLeft,
            "dptxtbrlv" => ControlWord::LegacyTextTopBottomRightLeftVertical,
            "dptxbtlr" => ControlWord::LegacyTextBottomTopLeftRight,
            "dolock" => ControlWord::LegacyDrawingLock,
            "dpgroup" => ControlWord::LegacyDrawingGroup,
            "dpcount" => ControlWord::LegacyDrawingCount(param_value),
            "dppolycount" => ControlWord::LegacyDrawingCount(param_value),
            "dpendgroup" => ControlWord::LegacyDrawingEndGroup,
            "dparc" => ControlWord::LegacyDrawingArc,
            "dpcallout" => ControlWord::LegacyDrawingCallout,
            "dpellipse" => ControlWord::LegacyDrawingEllipse,
            "dpline" => ControlWord::LegacyDrawingLine,
            "dppolygon" => ControlWord::LegacyDrawingPolygon,
            "dppolyline" => ControlWord::LegacyDrawingPolyline,
            "dprect" => ControlWord::LegacyDrawingRectangle,
            "dproundr" => ControlWord::LegacyDrawingRoundRectangle,
            "dpptx" => ControlWord::LegacyDrawingPointX(param_value),
            "dppty" => ControlWord::LegacyDrawingPointY(param_value),
            "dparcflipx" => ControlWord::LegacyDrawingArcFlipX,
            "dparcflipy" => ControlWord::LegacyDrawingArcFlipY,
            "dpcotright" => ControlWord::LegacyCalloutType(crate::LegacyCalloutType::RightAngle),
            "dpcotsingle" => ControlWord::LegacyCalloutType(crate::LegacyCalloutType::Single),
            "dpcotdouble" => ControlWord::LegacyCalloutType(crate::LegacyCalloutType::Double),
            "dpcottriple" => ControlWord::LegacyCalloutType(crate::LegacyCalloutType::Triple),
            "dpcoa" => ControlWord::LegacyCalloutAngle(param_value),
            "dpcoaccent" => ControlWord::LegacyCalloutAccent,
            "dpcosmarta" => ControlWord::LegacyCalloutSmartAttach,
            "dpcobestfit" => ControlWord::LegacyCalloutBestFit,
            "dpcominusx" => ControlWord::LegacyCalloutMinusX,
            "dpcominusy" => ControlWord::LegacyCalloutMinusY,
            "dpcoborder" => ControlWord::LegacyCalloutBorder,
            "dpcodtop" => ControlWord::LegacyCalloutAttachment(crate::LegacyCalloutAttachment::Top),
            "dpcodcenter" => {
                ControlWord::LegacyCalloutAttachment(crate::LegacyCalloutAttachment::Center)
            },
            "dpcodbottom" => {
                ControlWord::LegacyCalloutAttachment(crate::LegacyCalloutAttachment::Bottom)
            },
            "dpcodabs" => {
                ControlWord::LegacyCalloutAttachment(crate::LegacyCalloutAttachment::Absolute)
            },
            "dpcodescent" => ControlWord::LegacyCalloutDescent(param_value),
            "dpcooffset" => ControlWord::LegacyCalloutOffset(param_value),
            "dpcolength" => ControlWord::LegacyCalloutLength(param_value),
            "dplinesolid" => {
                ControlWord::LegacyDrawingLineStyle(crate::LegacyDrawingLineStyle::Solid)
            },
            "dplinehollow" => {
                ControlWord::LegacyDrawingLineStyle(crate::LegacyDrawingLineStyle::Hollow)
            },
            "dplinedash" => {
                ControlWord::LegacyDrawingLineStyle(crate::LegacyDrawingLineStyle::Dashed)
            },
            "dplinedot" => {
                ControlWord::LegacyDrawingLineStyle(crate::LegacyDrawingLineStyle::Dotted)
            },
            "dplinedado" => {
                ControlWord::LegacyDrawingLineStyle(crate::LegacyDrawingLineStyle::DashDot)
            },
            "dplinedadodo" => {
                ControlWord::LegacyDrawingLineStyle(crate::LegacyDrawingLineStyle::DashDotDot)
            },
            "dplinegray" => ControlWord::LegacyDrawingLineGray(param_value),
            "dplinecor" => ControlWord::LegacyDrawingLineRed(param_value),
            "dplinecog" => ControlWord::LegacyDrawingLineGreen(param_value),
            "dplinecob" => ControlWord::LegacyDrawingLineBlue(param_value),
            "dplinepal" => ControlWord::LegacyDrawingLinePalette,
            "dplinew" => ControlWord::LegacyDrawingLineWidth(param_value),
            "dpfillfggray" => ControlWord::LegacyDrawingFillForegroundGray(param_value),
            "dpfillfgcr" => ControlWord::LegacyDrawingFillForegroundRed(param_value),
            "dpfillfgcg" => ControlWord::LegacyDrawingFillForegroundGreen(param_value),
            "dpfillfgcb" => ControlWord::LegacyDrawingFillForegroundBlue(param_value),
            "dpfillfgpal" => ControlWord::LegacyDrawingFillForegroundPalette,
            "dpfillbggray" => ControlWord::LegacyDrawingFillBackgroundGray(param_value),
            "dpfillbgcr" => ControlWord::LegacyDrawingFillBackgroundRed(param_value),
            "dpfillbgcg" => ControlWord::LegacyDrawingFillBackgroundGreen(param_value),
            "dpfillbgcb" => ControlWord::LegacyDrawingFillBackgroundBlue(param_value),
            "dpfillbgpal" => ControlWord::LegacyDrawingFillBackgroundPalette,
            "dpfillpat" => ControlWord::LegacyDrawingFillPattern(param_value),
            "dpastartsol" => {
                ControlWord::LegacyDrawingStartArrowFill(crate::LegacyDrawingArrowFill::Solid)
            },
            "dpastarthol" => {
                ControlWord::LegacyDrawingStartArrowFill(crate::LegacyDrawingArrowFill::Hollow)
            },
            "dpastartl" => ControlWord::LegacyDrawingStartArrowLength(param_value),
            "dpastartw" => ControlWord::LegacyDrawingStartArrowWidth(param_value),
            "dpaendsol" => {
                ControlWord::LegacyDrawingEndArrowFill(crate::LegacyDrawingArrowFill::Solid)
            },
            "dpaendhol" => {
                ControlWord::LegacyDrawingEndArrowFill(crate::LegacyDrawingArrowFill::Hollow)
            },
            "dpaendl" => ControlWord::LegacyDrawingEndArrowLength(param_value),
            "dpaendw" => ControlWord::LegacyDrawingEndArrowWidth(param_value),
            "dpshadow" => ControlWord::LegacyDrawingShadow,
            "dpshadx" => ControlWord::LegacyDrawingShadowX(param_value),
            "dpshady" => ControlWord::LegacyDrawingShadowY(param_value),
            "background" => ControlWord::BackgroundDestination(param),

            // Document info
            "title" => ControlWord::Title,
            "subject" => ControlWord::Subject,
            "author" => ControlWord::Author,
            "manager" => ControlWord::Manager,
            "company" => ControlWord::Company,
            "operator" => ControlWord::Operator,
            "category" => ControlWord::Category,
            "keywords" => ControlWord::Keywords,
            "comment" => ControlWord::Comment,
            "doccomm" => ControlWord::DocComment,
            "hlinkbase" => ControlWord::HyperlinkBase,
            "creatim" => ControlWord::CreationTime,
            "revtim" => ControlWord::RevisionTime,
            "printim" => ControlWord::PrintTime,
            "buptim" => ControlWord::BackupTime,
            "version" => ControlWord::InfoVersion(param_value),
            "vern" => ControlWord::InfoRevision(param_value),
            "edmins" => ControlWord::EditingTime(param_value),
            "nofpages" => ControlWord::NumberOfPages(param_value),
            "nofwords" => ControlWord::NumberOfWords(param_value),
            "nofchars" => ControlWord::NumberOfCharacters(param_value),
            "nofcharsws" => ControlWord::NumberOfCharactersWithSpaces(param_value),
            "id" => ControlWord::DocumentId(param_value),
            "yr" => ControlWord::Year(param_value),
            "mo" => ControlWord::Month(param_value),
            "dy" => ControlWord::Day(param_value),
            "hr" => ControlWord::Hour(param_value),
            "min" => ControlWord::Minute(param_value),
            "sec" => ControlWord::Second(param_value),

            // Unicode
            "u" => ControlWord::Unicode(param_value),
            "uc" => ControlWord::UnicodeSkip(param_value),

            // Special
            "tab" => ControlWord::Tab,
            "line" => ControlWord::Line,
            "emdash" => ControlWord::EmDash,
            "endash" => ControlWord::EnDash,
            "emspace" => ControlWord::EmSpace,
            "enspace" => ControlWord::EnSpace,
            "qmspace" => ControlWord::QuarterEmSpace,
            "bullet" => ControlWord::Bullet,
            "ltrmark" => ControlWord::LeftToRightMark,
            "rtlmark" => ControlWord::RightToLeftMark,
            "zwj" => ControlWord::ZeroWidthJoiner,
            "zwnj" => ControlWord::ZeroWidthNonJoiner,
            "chdate" => ControlWord::CurrentDate,
            "chdpl" => ControlWord::CurrentDateLong,
            "chdpa" => ControlWord::CurrentDateAbbreviated,
            "chtime" => ControlWord::CurrentTime,

            // Binary data
            "bin" => ControlWord::Binary(param_value),

            // Unknown
            _ => ControlWord::Unknown(word, param),
        };

        Ok(control)
    }

    /// Parse hexadecimal character escape (\').
    fn parse_hex_char(&mut self) -> RtfResult<Token<'a>> {
        let mut bytes = String::new();
        loop {
            self.advance(); // Skip '\''
            if self.pos + 1 >= self.input.len() {
                return Err(RtfError::InvalidUnicode(
                    "Incomplete hex escape".to_string(),
                ));
            }
            let hex = &self.input[self.pos..self.pos + 2];
            self.pos += 2;
            let byte = u8::from_str_radix(hex, 16)
                .map_err(|_| RtfError::InvalidUnicode(format!("Invalid hex escape: {hex}")))?;
            bytes.push(char::from(byte));
            if !self.input[self.pos..].starts_with("\\'") {
                break;
            }
            self.advance(); // Skip the next backslash; the loop skips its quote.
        }

        // Preserve source bytes. The parser applies the active code page after
        // interpreting document and group-level encoding controls. Consecutive
        // escapes stay together so multibyte encodings decode atomically.
        let text = self.arena.alloc_str(&bytes);
        Ok(Token::Text(Cow::Borrowed(text)))
    }

    /// Parse plain text until special character.
    fn parse_text(&mut self) -> RtfResult<Token<'a>> {
        let mut text = String::new();

        while self.pos < self.input.len() {
            let ch = self.current_char();
            match ch {
                '\\' | '{' | '}' => break,
                '\r' | '\n' => {
                    // Skip line breaks in plain text, but track them
                    self.advance();
                    // If we have accumulated text, break here
                    if !text.is_empty() {
                        break;
                    }
                },
                _ => {
                    text.push(ch);
                    self.advance();
                },
            }
        }

        if text.is_empty() {
            // If we hit only whitespace/newlines, try to consume at least one whitespace
            // and return a space token, or skip to next token
            if self.pos >= self.input.len() {
                // Trailing physical line breaks are insignificant in RTF. Reaching EOF
                // after consuming only those line breaks is a successful final token,
                // not a truncated control word or escape.
                let allocated = self.arena.alloc_str("");
                return Ok(Token::Text(Cow::Borrowed(allocated)));
            }
            // Return empty text for now - parser will handle it
            let allocated = self.arena.alloc_str("");
            return Ok(Token::Text(Cow::Borrowed(allocated)));
        }

        let allocated = self.arena.alloc_str(&text);
        Ok(Token::Text(Cow::Borrowed(allocated)))
    }

    /// Get current character without advancing.
    #[inline]
    fn current_char(&self) -> char {
        self.input[self.pos..].chars().next().unwrap_or('\0')
    }

    /// Advance position by one character.
    #[inline]
    fn advance(&mut self) {
        if self.pos < self.input.len() {
            let ch = self.current_char();
            self.pos += ch.len_utf8();
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_tokenization() {
        let arena = Bump::new();
        let input = r"{\rtf1\ansi Hello}";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();

        assert_eq!(tokens.len(), 5);
        assert!(matches!(tokens[0], Token::OpenBrace));
        assert!(matches!(tokens[1], Token::Control(ControlWord::Rtf(1))));
        assert!(matches!(tokens[2], Token::Control(ControlWord::Ansi)));
    }

    #[test]
    fn test_lexer_new() {
        let arena = Bump::new();
        let lexer = Lexer::new(r"{\rtf1}", &arena);
        assert_eq!(lexer.pos, 0);
    }

    #[test]
    fn test_character_set_variants() {
        assert_eq!(CharacterSet::default(), CharacterSet::Ansi);
        assert_ne!(CharacterSet::Ansi, CharacterSet::Mac);
    }

    #[test]
    fn test_token_variants() {
        let token = Token::OpenBrace;
        assert!(matches!(token, Token::OpenBrace));
        let token = Token::CloseBrace;
        assert!(matches!(token, Token::CloseBrace));
    }

    #[test]
    fn test_control_word_variants() {
        let word = ControlWord::Rtf(1);
        assert!(matches!(word, ControlWord::Rtf(1)));
        let word = ControlWord::Bold(true);
        assert!(matches!(word, ControlWord::Bold(true)));
        let word = ControlWord::FontNumber(0);
        assert!(matches!(word, ControlWord::FontNumber(0)));
    }

    #[test]
    fn test_tokenize_empty_braces() {
        let arena = Bump::new();
        let input = "{}";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 2);
        assert!(matches!(tokens[0], Token::OpenBrace));
        assert!(matches!(tokens[1], Token::CloseBrace));
    }

    #[test]
    fn test_tokenize_control_word_with_param() {
        let arena = Bump::new();
        let input = r"{\rtf1}";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(tokens[1], Token::Control(ControlWord::Rtf(1))));
    }

    #[test]
    fn test_tokenize_control_word_without_param() {
        let arena = Bump::new();
        let input = r"{\b}";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert!(matches!(tokens[1], Token::Control(ControlWord::Bold(true))));
    }

    #[test]
    fn test_tokenize_text() {
        let arena = Bump::new();
        let input = r"{Hello}";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(tokens[1], Token::Text(_)));
    }

    #[test]
    fn test_tokenize_multiple_control_words() {
        let arena = Bump::new();
        let input = r"{\rtf1\ansi\deff0}";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 5);
        assert!(matches!(tokens[1], Token::Control(ControlWord::Rtf(1))));
        assert!(matches!(tokens[2], Token::Control(ControlWord::Ansi)));
    }

    #[test]
    fn test_tokenize_escaped_braces() {
        let arena = Bump::new();
        let input = r"\{\}";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 2);
        // Should be text tokens containing { and }
        assert!(matches!(tokens[0], Token::Text(_)));
        assert!(matches!(tokens[1], Token::Text(_)));
    }

    #[test]
    fn test_tokenize_backslash_escape() {
        let arena = Bump::new();
        let input = r"\\";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert!(matches!(tokens[0], Token::Text(_)));
    }

    #[test]
    fn test_tokenize_hex_escape() {
        let arena = Bump::new();
        let input = r"\'41"; // 'A' in hex
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert!(matches!(tokens[0], Token::Text(_)));
    }

    #[test]
    fn test_hex_escape_preserves_following_literal_space() {
        let arena = Bump::new();
        let mut lexer = Lexer::new(r"\'80 value", &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 2);
        assert_eq!(tokens[0], Token::Text(Cow::Borrowed("\u{0080}")));
        assert_eq!(tokens[1], Token::Text(Cow::Borrowed(" value")));
    }

    #[test]
    fn test_tokenize_asterisk_destination() {
        let arena = Bump::new();
        let input = r"\*";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert!(matches!(
            tokens[0],
            Token::Control(ControlWord::IgnorableDestination)
        ));
    }

    #[test]
    fn test_tokenize_par_control_word() {
        let arena = Bump::new();
        let input = r"\par";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert!(matches!(tokens[0], Token::Control(ControlWord::Par)));
    }

    #[test]
    fn test_tokenize_non_breaking_space() {
        let arena = Bump::new();
        let input = r"\~";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert_eq!(tokens[0], Token::Control(ControlWord::NonBreakingSpace));
    }

    #[test]
    fn test_tokenize_optional_hyphen() {
        let arena = Bump::new();
        let input = r"\-";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert_eq!(tokens[0], Token::Control(ControlWord::OptionalHyphen));
    }

    #[test]
    fn test_tokenize_non_breaking_hyphen() {
        let arena = Bump::new();
        let input = r"\_";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert_eq!(tokens[0], Token::Control(ControlWord::NonBreakingHyphen));
    }

    #[test]
    fn test_tokenize_special_character_control_words() {
        let cases: &[(&str, ControlWord<'_>)] = &[
            (r"\emdash", ControlWord::EmDash),
            (r"\endash", ControlWord::EnDash),
            (r"\emspace", ControlWord::EmSpace),
            (r"\enspace", ControlWord::EnSpace),
            (r"\qmspace", ControlWord::QuarterEmSpace),
            (r"\bullet", ControlWord::Bullet),
            (r"\ltrmark", ControlWord::LeftToRightMark),
            (r"\rtlmark", ControlWord::RightToLeftMark),
            (r"\zwj", ControlWord::ZeroWidthJoiner),
            (r"\zwnj", ControlWord::ZeroWidthNonJoiner),
            (r"\chdate", ControlWord::CurrentDate),
            (r"\chdpl", ControlWord::CurrentDateLong),
            (r"\chdpa", ControlWord::CurrentDateAbbreviated),
            (r"\chtime", ControlWord::CurrentTime),
        ];
        for (input, expected) in cases {
            let arena = Bump::new();
            let mut lexer = Lexer::new(input, &arena);
            let tokens = lexer.tokenize().unwrap();
            assert_eq!(tokens.len(), 1, "tokenized {input}");
            assert_eq!(tokens[0], Token::Control(*expected), "lexed {input}");
        }
    }

    #[test]
    fn test_tokenize_column_break() {
        let arena = Bump::new();
        let input = r"\column";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert_eq!(tokens[0], Token::Control(ControlWord::Column(None)));
    }

    #[test]
    fn test_tokenize_tab() {
        let arena = Bump::new();
        let input = r"\tab";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert!(matches!(tokens[0], Token::Control(ControlWord::Tab)));
    }

    #[test]
    fn test_tokenize_par() {
        let arena = Bump::new();
        let input = r"\par";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert!(matches!(tokens[0], Token::Control(ControlWord::Par)));
    }

    #[test]
    fn test_tokenize_line() {
        let arena = Bump::new();
        let input = r"\line";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert!(matches!(tokens[0], Token::Control(ControlWord::Line)));
    }

    #[test]
    fn test_tokenize_page() {
        let arena = Bump::new();
        let input = r"\page";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert!(matches!(
            tokens[0],
            Token::Control(ControlWord::Page(None))
        ));
    }

    #[test]
    fn test_tokenize_font_size() {
        let arena = Bump::new();
        let input = r"\fs24";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert!(matches!(
            tokens[0],
            Token::Control(ControlWord::FontSize(24))
        ));
    }

    #[test]
    fn test_tokenize_font_number() {
        let arena = Bump::new();
        let input = r"\f0";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert!(matches!(
            tokens[0],
            Token::Control(ControlWord::FontNumber(0))
        ));
    }

    #[test]
    fn test_tokenize_bold_toggle() {
        let arena = Bump::new();
        let input = r"\b\b0";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 2);
        assert!(matches!(tokens[0], Token::Control(ControlWord::Bold(true))));
        assert!(matches!(
            tokens[1],
            Token::Control(ControlWord::Bold(false))
        ));
    }

    #[test]
    fn test_tokenize_italic_toggle() {
        let arena = Bump::new();
        let input = r"\i\i0";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 2);
        assert!(matches!(
            tokens[0],
            Token::Control(ControlWord::Italic(true))
        ));
        assert!(matches!(
            tokens[1],
            Token::Control(ControlWord::Italic(false))
        ));
    }

    #[test]
    fn test_tokenize_underline_variants() {
        let arena = Bump::new();
        let input = r"\ul\ulnone\uldb";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(
            tokens[0],
            Token::Control(ControlWord::Underline(true))
        ));
        assert!(matches!(
            tokens[1],
            Token::Control(ControlWord::UnderlineNone)
        ));
        assert!(matches!(
            tokens[2],
            Token::Control(ControlWord::UnderlineDouble)
        ));
    }

    #[test]
    fn test_tokenize_alignment() {
        let arena = Bump::new();
        let input = r"\ql\qr\qc\qj";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 4);
        assert!(matches!(tokens[0], Token::Control(ControlWord::LeftAlign)));
        assert!(matches!(tokens[1], Token::Control(ControlWord::RightAlign)));
        assert!(matches!(tokens[2], Token::Control(ControlWord::Center)));
        assert!(matches!(tokens[3], Token::Control(ControlWord::Justify)));
    }

    #[test]
    fn test_tokenize_color_table() {
        let arena = Bump::new();
        let input = r"{\colortbl;\red255\green0\blue0;}";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert!(matches!(tokens[1], Token::Control(ControlWord::ColorTable)));
        assert!(matches!(tokens[3], Token::Control(ControlWord::Red(255))));
        assert!(matches!(tokens[4], Token::Control(ControlWord::Green(0))));
        assert!(matches!(tokens[5], Token::Control(ControlWord::Blue(0))));
    }

    #[test]
    fn test_tokenize_font_table() {
        let arena = Bump::new();
        let input = r"{\fonttbl{\f0\fnil Arial;}}";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert!(matches!(tokens[1], Token::Control(ControlWord::FontTable)));
    }

    #[test]
    fn test_tokenize_unknown_control_word() {
        let arena = Bump::new();
        let input = r"\xyz123";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert!(matches!(
            tokens[0],
            Token::Control(ControlWord::Unknown(_, _))
        ));
    }

    #[test]
    fn test_tokenize_section_break() {
        let arena = Bump::new();
        let input = r"\sect";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1);
        assert!(matches!(tokens[0], Token::Control(ControlWord::Section)));
    }

    #[test]
    fn test_tokenize_page_dimensions() {
        let arena = Bump::new();
        let input = r"\paperw12240\paperh15840";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert!(matches!(
            tokens[0],
            Token::Control(ControlWord::PageWidth(12240))
        ));
        assert!(matches!(
            tokens[1],
            Token::Control(ControlWord::PageHeight(15840))
        ));
    }

    #[test]
    fn test_tokenize_margins() {
        let arena = Bump::new();
        let input = r"\margl1440\margr1440\margt1440\margb1440";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert!(matches!(
            tokens[0],
            Token::Control(ControlWord::MarginLeft(1440))
        ));
        assert!(matches!(
            tokens[1],
            Token::Control(ControlWord::MarginRight(1440))
        ));
        assert!(matches!(
            tokens[2],
            Token::Control(ControlWord::MarginTop(1440))
        ));
        assert!(matches!(
            tokens[3],
            Token::Control(ControlWord::MarginBottom(1440))
        ));
    }

    #[test]
    fn test_tokenize_lists() {
        let arena = Bump::new();
        let input = r"\listtable\listid1\listlevel1";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 3);
        assert!(matches!(tokens[0], Token::Control(ControlWord::ListTable)));
        assert!(matches!(tokens[1], Token::Control(ControlWord::ListId(1))));
        assert!(matches!(tokens[2], Token::Control(ControlWord::ListLevel)));
    }

    #[test]
    fn test_tokenize_table() {
        let arena = Bump::new();
        let input = r"\trowd\cellx4320";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 2);
        assert!(matches!(
            tokens[0],
            Token::Control(ControlWord::TableRowDefaults)
        ));
        assert!(matches!(
            tokens[1],
            Token::Control(ControlWord::CellX(4320))
        ));
    }

    #[test]
    fn test_tokenize_field() {
        let arena = Bump::new();
        let input = r"\field\fldinst HYPERLINK";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert!(matches!(tokens[0], Token::Control(ControlWord::Field)));
        assert!(matches!(
            tokens[1],
            Token::Control(ControlWord::FieldInstruction)
        ));
        // HYPERLINK is parsed as text since it's after a space
        assert!(matches!(tokens[2], Token::Text(_)));
    }

    #[test]
    fn test_tokenize_picture() {
        let arena = Bump::new();
        let input = r"{\pict\pngblip\picw100\pich100}";
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert!(matches!(tokens[1], Token::Control(ControlWord::Picture)));
        assert!(matches!(tokens[2], Token::Control(ControlWord::Pngblip)));
        assert!(matches!(
            tokens[3],
            Token::Control(ControlWord::PictureWidth(100))
        ));
        assert!(matches!(
            tokens[4],
            Token::Control(ControlWord::PictureHeight(100))
        ));
    }

    #[test]
    fn test_tokenize_binary() {
        let arena = Bump::new();
        let input = r"\bin4 ABCD"; // 4 bytes of binary data
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert!(matches!(&tokens[0], Token::Binary(data) if data.as_ref() == b"ABCD"));
    }

    #[test]
    fn test_rejects_truncated_binary_data() {
        let arena = Bump::new();
        let mut lexer = Lexer::new(r"\bin4 AB", &arena);
        assert!(matches!(lexer.tokenize(), Err(RtfError::UnexpectedEof)));
    }

    #[test]
    fn test_tokenize_document_with_trailing_line_breaks() {
        let arena = Bump::new();
        let mut lexer = Lexer::new("{\\rtf1 body}\r\n", &arena);
        let tokens = lexer.tokenize().unwrap();

        assert!(matches!(tokens.last(), Some(Token::Text(text)) if text.is_empty()));
        assert!(matches!(
            tokens.get(tokens.len().saturating_sub(2)),
            Some(Token::CloseBrace)
        ));
    }
}
