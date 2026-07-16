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

    // Stylesheet entries and metadata
    ParagraphStyle(i32),
    CharacterStyle(i32),
    SectionStyle(i32),
    TableStyle(i32),
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
    Object,
    Result,
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
    ObjectWidth(i32),
    ObjectHeight(i32),
    ObjectLocked(bool),
    ObjectUpdate(bool),
    ObjectSetSize(bool),

    // Picture properties
    PictureWidth(i32),
    PictureHeight(i32),
    PictureGoalWidth(i32),
    PictureGoalHeight(i32),
    PictureScaleX(i32),
    PictureScaleY(i32),
    Emfblip,
    Pngblip,
    Jpegblip,
    Macpict,
    Pmmetafile(i32),
    Wmetafile(i32),
    Dibitmap(i32),
    Wbitmap(i32),

    // Field support
    Field,
    FieldInstruction,
    FieldResult,
    FieldLock,
    FieldDirty,
    FieldEdit,
    FieldPrivate,
    FormField,
    DataField,
    FormFieldType(i32),
    FormFieldName,
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
    XmlNamespaceTable,
    XmlNamespace(i32),
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
    LeftToRightRow,
    RightToLeftRow,
    TableRightToLeft(bool),
    RightGutter(bool),

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

    // Colors
    Red(i32),
    Green(i32),
    Blue(i32),
    ColorForeground(i32),
    ColorBackground(i32),

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
    Strike(bool),
    DoubleStrike(bool),
    Superscript(bool),
    Subscript(bool),
    SmallCaps(bool),
    AllCaps(bool),
    Hidden(bool),
    Outline(bool),
    Shadow(bool),
    Emboss(bool),
    Imprint(bool),
    CharSpacing(i32),
    CharScale(i32),
    Kerning(i32),
    Highlight(i32),
    Plain,

    // Paragraph formatting
    Par,
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
    LeftIndent(i32),
    RightIndent(i32),
    FirstLineIndent(i32),

    // Paragraph additional properties
    KeepTogether,
    KeepNext,
    PageBreakBefore,
    WidowControl,

    // Tables
    TableRowDefaults,
    TableRow,
    TableCell,
    CellX(i32),
    InTable,

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
    BorderWidth(i32),
    BorderColor(i32),
    BorderSpace(i32),

    // Shading
    Shading(i32),
    ForegroundPattern(i32),
    BackgroundPattern(i32),

    // Tab stops
    TabLeft,
    TabRight,
    TabCenter,
    TabDecimal,
    TabBar,
    TabPosition(i32),
    TabLeaderDot,
    TabLeaderHyphen,
    TabLeaderUnderscore,
    TabLeaderThick,

    // Lists
    ListTable,
    List,
    ListTemplateId(i32),
    ListSimple(bool),
    ListHybrid(bool),
    ListName,
    ListId(i32),
    ListOverrideTable,
    ListOverride,
    ListOverrideCount(i32),
    ListOverrideIndex(i32),
    ListLevelIndex(i32),
    ListOverrideLevel,
    ListOverrideStartAt(bool),
    ListLevel,
    ListLevelType(i32),
    ListLevelJustification(i32),
    ListLevelFollow(i32),
    ListLevelSpace(i32),
    ListLevelIndent(i32),
    ListNumberText,
    ListLevelStartAt(i32),
    ListLevelNumbers,

    // Sections
    SectionBreak,
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
    HeaderDistance(i32),
    FooterDistance(i32),
    Landscape,
    Columns(i32),
    ColumnSpace(i32),
    PageNumberStart(i32),
    PageNumberDecimal,
    PageNumberUpperRoman,
    PageNumberLowerRoman,
    PageNumberUpperLetter,
    PageNumberLowerLetter,
    VerticalAlignTop,
    VerticalAlignCenter,
    VerticalAlignJustify,
    VerticalAlignBottom,
    LineNumbering(i32),
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
    EndnoteNumber(i32),

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
    Shape,
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
    ShapeGroup,
    ShapeProperty,
    ShapePropertyName,
    ShapePropertyValue,
    ShapeText,
    BackgroundDestination,

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
    Page,
    Section,
    SectionDefault,
    NonBreakingSpace,
    OptionalHyphen,
    NonBreakingHyphen,

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

            // Stylesheet entries and metadata
            "s" => ControlWord::ParagraphStyle(param_value),
            "cs" => ControlWord::CharacterStyle(param_value),
            "ds" => ControlWord::SectionStyle(param_value),
            "ts" => ControlWord::TableStyle(param_value),
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
            "object" => ControlWord::Object,
            "result" => ControlWord::Result,
            "objclass" => ControlWord::ObjectClass,
            "docvar" => ControlWord::DocumentVariable,
            "userprops" => ControlWord::UserProperties,
            "propname" => ControlWord::PropertyName,
            "proptype" => ControlWord::PropertyType(param),
            "staticval" => ControlWord::StaticValue,
            "linkval" => ControlWord::LinkValue,
            "objname" => ControlWord::ObjectName,
            "objdata" => ControlWord::ObjectData,
            "objemb" => ControlWord::ObjectEmbedded,
            "objlink" => ControlWord::ObjectLink,
            "objautlink" => ControlWord::ObjectAutoLink,
            "objhtml" => ControlWord::ObjectHtml,
            "objw" => ControlWord::ObjectWidth(param_value),
            "objh" => ControlWord::ObjectHeight(param_value),
            "objlock" => ControlWord::ObjectLocked(param_bool),
            "objupdate" => ControlWord::ObjectUpdate(param_bool),
            "objsetsize" => ControlWord::ObjectSetSize(param_bool),

            // Picture properties
            "picw" => ControlWord::PictureWidth(param_value),
            "pich" => ControlWord::PictureHeight(param_value),
            "picwgoal" => ControlWord::PictureGoalWidth(param_value),
            "pichgoal" => ControlWord::PictureGoalHeight(param_value),
            "picscalex" => ControlWord::PictureScaleX(param_value),
            "picscaley" => ControlWord::PictureScaleY(param_value),
            "emfblip" => ControlWord::Emfblip,
            "pngblip" => ControlWord::Pngblip,
            "jpegblip" => ControlWord::Jpegblip,
            "macpict" => ControlWord::Macpict,
            "pmmetafile" => ControlWord::Pmmetafile(param_value),
            "wmetafile" => ControlWord::Wmetafile(param_value),
            "dibitmap" => ControlWord::Dibitmap(param_value),
            "wbitmap" => ControlWord::Wbitmap(param_value),

            // Field support
            "field" => ControlWord::Field,
            "fldinst" => ControlWord::FieldInstruction,
            "fldrslt" => ControlWord::FieldResult,
            "fldlock" => ControlWord::FieldLock,
            "flddirty" => ControlWord::FieldDirty,
            "fldedit" => ControlWord::FieldEdit,
            "fldpriv" => ControlWord::FieldPrivate,
            "formfield" => ControlWord::FormField,
            "datafield" => ControlWord::DataField,
            "fftype" => ControlWord::FormFieldType(param_value),
            "ffname" => ControlWord::FormFieldName,
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
            "xmlnstbl" => ControlWord::XmlNamespaceTable,
            "xmlns" => ControlWord::XmlNamespace(param_value),
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
            "ltrrow" => ControlWord::LeftToRightRow,
            "rtlrow" => ControlWord::RightToLeftRow,
            "taprtl" => ControlWord::TableRightToLeft(param_bool),
            "rtlgutter" => ControlWord::RightGutter(param_bool),

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
            "fcharset" => ControlWord::FontCharset(param_value),
            "fnil" => ControlWord::FontFamily("nil"),
            "froman" => ControlWord::FontFamily("roman"),
            "fswiss" => ControlWord::FontFamily("swiss"),
            "fmodern" => ControlWord::FontFamily("modern"),
            "fscript" => ControlWord::FontFamily("script"),
            "fdecor" => ControlWord::FontFamily("decor"),
            "ftech" => ControlWord::FontFamily("tech"),

            // Colors
            "red" => ControlWord::Red(param_value),
            "green" => ControlWord::Green(param_value),
            "blue" => ControlWord::Blue(param_value),
            "cf" => ControlWord::ColorForeground(param_value),
            "cb" => ControlWord::ColorBackground(param_value),

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
            "strike" => ControlWord::Strike(param_bool),
            "striked" => ControlWord::DoubleStrike(param_bool),
            "super" => ControlWord::Superscript(param_bool),
            "sub" => ControlWord::Subscript(param_bool),
            "scaps" => ControlWord::SmallCaps(param_bool),
            "caps" => ControlWord::AllCaps(param_bool),
            "v" => ControlWord::Hidden(param_bool),
            "outl" => ControlWord::Outline(param_bool),
            "shad" => ControlWord::Shadow(param_bool),
            "embo" => ControlWord::Emboss(param_bool),
            "impr" => ControlWord::Imprint(param_bool),
            "expnd" => ControlWord::CharSpacing(param_value),
            "charscalex" => ControlWord::CharScale(param_value),
            "kerning" => ControlWord::Kerning(param_value),
            "highlight" => ControlWord::Highlight(param_value),
            "plain" => ControlWord::Plain,

            // Paragraph
            "par" => ControlWord::Par,
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
            "li" => ControlWord::LeftIndent(param_value),
            "ri" => ControlWord::RightIndent(param_value),
            "fi" => ControlWord::FirstLineIndent(param_value),

            // Paragraph additional properties
            "keep" => ControlWord::KeepTogether,
            "keepn" => ControlWord::KeepNext,
            "pagebb" => ControlWord::PageBreakBefore,
            "widctlpar" => ControlWord::WidowControl,

            // Tables
            "trowd" => ControlWord::TableRowDefaults,
            "row" => ControlWord::TableRow,
            "cell" => ControlWord::TableCell,
            "cellx" => ControlWord::CellX(param_value),
            "intbl" => ControlWord::InTable,

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
            "brdrw" => ControlWord::BorderWidth(param_value),
            "brdrcf" => ControlWord::BorderColor(param_value),
            "brsp" => ControlWord::BorderSpace(param_value),

            // Shading
            "shading" => ControlWord::Shading(param_value),
            "cfpat" => ControlWord::ForegroundPattern(param_value),
            "cbpat" => ControlWord::BackgroundPattern(param_value),

            // Tab stops
            "tql" => ControlWord::TabLeft,
            "tqr" => ControlWord::TabRight,
            "tqc" => ControlWord::TabCenter,
            "tqdec" => ControlWord::TabDecimal,
            "tb" => ControlWord::TabBar,
            "tx" => ControlWord::TabPosition(param_value),
            "tldot" => ControlWord::TabLeaderDot,
            "tlhyph" => ControlWord::TabLeaderHyphen,
            "tlul" => ControlWord::TabLeaderUnderscore,
            "tlth" => ControlWord::TabLeaderThick,

            // Lists
            "listtable" => ControlWord::ListTable,
            "list" => ControlWord::List,
            "listtemplateid" => ControlWord::ListTemplateId(param_value),
            "listsimple" => ControlWord::ListSimple(param_bool),
            "listhybrid" => ControlWord::ListHybrid(param_bool),
            "listname" => ControlWord::ListName,
            "listid" => ControlWord::ListId(param_value),
            "listoverridetable" => ControlWord::ListOverrideTable,
            "listoverride" => ControlWord::ListOverride,
            "listoverridecount" => ControlWord::ListOverrideCount(param_value),
            "ls" => ControlWord::ListOverrideIndex(param_value),
            "ilvl" => ControlWord::ListLevelIndex(param_value),
            "lfolevel" => ControlWord::ListOverrideLevel,
            "listoverridestartat" => ControlWord::ListOverrideStartAt(param_bool),
            "listlevel" => ControlWord::ListLevel,
            "levelnfc" => ControlWord::ListLevelType(param_value),
            "leveljc" => ControlWord::ListLevelJustification(param_value),
            "levelfollow" => ControlWord::ListLevelFollow(param_value),
            "levelspace" => ControlWord::ListLevelSpace(param_value),
            "levelindent" => ControlWord::ListLevelIndent(param_value),
            "leveltext" => ControlWord::ListNumberText,
            "levelstartat" => ControlWord::ListLevelStartAt(param_value),
            "levelnumbers" => ControlWord::ListLevelNumbers,

            // Sections
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
            "headery" => ControlWord::HeaderDistance(param_value),
            "footery" => ControlWord::FooterDistance(param_value),
            "landscape" | "lndscpsxn" => ControlWord::Landscape,
            "cols" => ControlWord::Columns(param_value),
            "colsx" => ControlWord::ColumnSpace(param_value),
            "pgnstarts" => ControlWord::PageNumberStart(param_value),
            "pgndec" => ControlWord::PageNumberDecimal,
            "pgnucrm" => ControlWord::PageNumberUpperRoman,
            "pgnlcrm" => ControlWord::PageNumberLowerRoman,
            "pgnucltr" => ControlWord::PageNumberUpperLetter,
            "pgnlcltr" => ControlWord::PageNumberLowerLetter,
            "vertalt" => ControlWord::VerticalAlignTop,
            "vertalc" => ControlWord::VerticalAlignCenter,
            "vertalj" => ControlWord::VerticalAlignJustify,
            "vertalb" => ControlWord::VerticalAlignBottom,
            "linemod" => ControlWord::LineNumbering(param_value),
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
            "chftnsepc" => ControlWord::EndnoteNumber(param_value),

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
            "shp" => ControlWord::Shape,
            "shpgrp" => ControlWord::ShapeGroup,
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
            "sp" => ControlWord::ShapeProperty,
            "sn" => ControlWord::ShapePropertyName,
            "sv" => ControlWord::ShapePropertyValue,
            "shptxt" => ControlWord::ShapeText,
            "background" => ControlWord::BackgroundDestination,

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
            "page" => ControlWord::Page,

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
        assert!(matches!(tokens[0], Token::Control(ControlWord::Page)));
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
