use crate::docx::enums::{WdHeaderFooter, WdSectionStart};
use crate::error::{OoxmlError, Result};
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};
use std::fmt::Write;

/// Page number format used by page, footnote, and endnote numbering.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageNumberFormat {
    Decimal,
    UpperRoman,
    LowerRoman,
    UpperLetter,
    LowerLetter,
}

impl PageNumberFormat {
    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::Decimal => "decimal",
            Self::UpperRoman => "upperRoman",
            Self::LowerRoman => "lowerRoman",
            Self::UpperLetter => "upperLetter",
            Self::LowerLetter => "lowerLetter",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "decimal" => Ok(Self::Decimal),
            "upperRoman" => Ok(Self::UpperRoman),
            "lowerRoman" => Ok(Self::LowerRoman),
            "upperLetter" => Ok(Self::UpperLetter),
            "lowerLetter" => Ok(Self::LowerLetter),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "unsupported section numbering format '{value}'"
            ))),
        }
    }
}

/// Page orientation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageOrientation {
    Portrait,
    Landscape,
}

impl PageOrientation {
    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::Portrait => "portrait",
            Self::Landscape => "landscape",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "portrait" => Ok(Self::Portrait),
            "landscape" => Ok(Self::Landscape),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid section page orientation '{value}'"
            ))),
        }
    }
}

/// One header or footer relationship used by a section.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionHeaderFooterReference {
    pub kind: WdHeaderFooter,
    /// Existing relationship ID. `None` binds a default reference to content
    /// created through [`crate::docx::writer::MutableDocument::header`] or `footer`.
    pub relationship_id: Option<String>,
    /// New or replacement part owned by this section. References with the same
    /// non-empty key are deliberately shared and must have identical XML.
    pub part: Option<SectionHeaderFooterPart>,
}

/// Header/footer XML staged as a package part during save.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionHeaderFooterPart {
    pub key: String,
    pub xml: String,
}

impl SectionHeaderFooterReference {
    pub fn existing(kind: WdHeaderFooter, relationship_id: impl Into<String>) -> Self {
        Self {
            kind,
            relationship_id: Some(relationship_id.into()),
            part: None,
        }
    }

    pub fn owned(
        kind: WdHeaderFooter,
        key: impl Into<String>,
        xml: impl Into<String>,
    ) -> Self {
        Self {
            kind,
            relationship_id: None,
            part: Some(SectionHeaderFooterPart {
                key: key.into(),
                xml: xml.into(),
            }),
        }
    }

    pub fn managed_default(kind: WdHeaderFooter) -> Self {
        Self {
            kind,
            relationship_id: None,
            part: None,
        }
    }
}

/// Page-numbering settings for a section.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionPageNumbering {
    pub format: PageNumberFormat,
    pub start: Option<u32>,
    pub chapter_style: Option<u8>,
    pub chapter_separator: Option<String>,
}

/// One explicitly sized newspaper-style column.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionColumn {
    pub width: u32,
    pub space: Option<u32>,
}

/// Section column layout.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionColumns {
    pub equal_width: bool,
    pub count: u16,
    pub space: Option<u32>,
    pub separator: bool,
    pub columns: Vec<SectionColumn>,
}

impl Default for SectionColumns {
    fn default() -> Self {
        Self {
            equal_width: true,
            count: 1,
            space: None,
            separator: false,
            columns: Vec::new(),
        }
    }
}

/// Restart policy for note numbering.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NoteNumberRestart {
    Continuous,
    EachSection,
    EachPage,
}

impl NoteNumberRestart {
    fn as_str(self) -> &'static str {
        match self {
            Self::Continuous => "continuous",
            Self::EachSection => "eachSect",
            Self::EachPage => "eachPage",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "continuous" => Ok(Self::Continuous),
            "eachSect" => Ok(Self::EachSection),
            "eachPage" => Ok(Self::EachPage),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid note-number restart '{value}'"
            ))),
        }
    }
}

/// Footnote/endnote numbering options within a section.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionNoteProperties {
    pub format: PageNumberFormat,
    pub start: Option<u32>,
    pub restart: Option<NoteNumberRestart>,
    /// Schema token such as `pageBottom`, `beneathText`, `sectEnd`, or `docEnd`.
    pub position: Option<String>,
}

/// Text flow for a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionTextDirection {
    LeftToRightTopToBottom,
    TopToBottomRightToLeft,
    BottomToTopLeftToRight,
    LeftToRightTopToBottomRotated,
    TopToBottomRightToLeftRotated,
    TopToBottomLeftToRightRotated,
}

impl SectionTextDirection {
    fn as_str(self) -> &'static str {
        match self {
            Self::LeftToRightTopToBottom => "lrTb",
            Self::TopToBottomRightToLeft => "tbRl",
            Self::BottomToTopLeftToRight => "btLr",
            Self::LeftToRightTopToBottomRotated => "lrTbV",
            Self::TopToBottomRightToLeftRotated => "tbRlV",
            Self::TopToBottomLeftToRightRotated => "tbLrV",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "lrTb" => Ok(Self::LeftToRightTopToBottom),
            "tbRl" => Ok(Self::TopToBottomRightToLeft),
            "btLr" => Ok(Self::BottomToTopLeftToRight),
            "lrTbV" => Ok(Self::LeftToRightTopToBottomRotated),
            "tbRlV" => Ok(Self::TopToBottomRightToLeftRotated),
            "tbLrV" => Ok(Self::TopToBottomLeftToRightRotated),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid section text direction '{value}'"
            ))),
        }
    }
}

/// Document-grid mode for East Asian layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentGridType {
    Default,
    Lines,
    LinesAndChars,
    SnapToChars,
}

impl DocumentGridType {
    fn as_str(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::Lines => "lines",
            Self::LinesAndChars => "linesAndChars",
            Self::SnapToChars => "snapToChars",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "default" => Ok(Self::Default),
            "lines" => Ok(Self::Lines),
            "linesAndChars" => Ok(Self::LinesAndChars),
            "snapToChars" => Ok(Self::SnapToChars),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid document grid type '{value}'"
            ))),
        }
    }
}

/// Section document-grid settings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionDocumentGrid {
    pub grid_type: DocumentGridType,
    pub line_pitch: Option<u32>,
    pub char_space: Option<i32>,
}

/// Maximum page-border size in eighths of a point for line-style borders
/// (`ST_EighthPointMeasure`, ECMA-376 §17.18.11).
const MAX_PAGE_BORDER_LINE_SIZE: u32 = 96;
/// Maximum page-border size in points for art borders
/// (`ST_PointMeasure`, ECMA-376 §17.18.68).
const MAX_PAGE_BORDER_ART_SIZE: u32 = 1638;
/// Maximum page-border spacing from text or page edge in points
/// (`w:space` on `CT_Border` is limited to 31, ECMA-376 §17.6.16).
const MAX_PAGE_BORDER_SPACE: u32 = 31;
/// Maximum length of a page-border art style name.
const MAX_PAGE_BORDER_ART_NAME_LEN: usize = 64;

/// Page-border style (`ST_Border`, ECMA-376 §17.18.2).
///
/// Line styles map to the fixed `ST_Border` tokens; border artwork uses
/// [`PageBorderStyle::Art`] with the art token preserved verbatim.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PageBorderStyle {
    Nil,
    None,
    Single,
    Thick,
    Double,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Triple,
    ThinThickSmallGap,
    ThinThickMediumGap,
    ThinThickLargeGap,
    ThickThinSmallGap,
    ThickThinMediumGap,
    ThickThinLargeGap,
    ThinThickThinSmallGap,
    ThinThickThinMediumGap,
    ThinThickThinLargeGap,
    Wave,
    DoubleWave,
    DashSmallGap,
    DashDotStroked,
    ThreeDEmboss,
    ThreeDEngrave,
    Outset,
    Inset,
    /// Border artwork token such as `apples` or `starsTop`.
    Art(String),
}

impl PageBorderStyle {
    fn as_str(&self) -> &str {
        match self {
            Self::Nil => "nil",
            Self::None => "none",
            Self::Single => "single",
            Self::Thick => "thick",
            Self::Double => "double",
            Self::Dotted => "dotted",
            Self::Dashed => "dashed",
            Self::DotDash => "dotDash",
            Self::DotDotDash => "dotDotDash",
            Self::Triple => "triple",
            Self::ThinThickSmallGap => "thinThickSmallGap",
            Self::ThinThickMediumGap => "thinThickMediumGap",
            Self::ThinThickLargeGap => "thinThickLargeGap",
            Self::ThickThinSmallGap => "thickThinSmallGap",
            Self::ThickThinMediumGap => "thickThinMediumGap",
            Self::ThickThinLargeGap => "thickThinLargeGap",
            Self::ThinThickThinSmallGap => "thinThickThinSmallGap",
            Self::ThinThickThinMediumGap => "thinThickThinMediumGap",
            Self::ThinThickThinLargeGap => "thinThickThinLargeGap",
            Self::Wave => "wave",
            Self::DoubleWave => "doubleWave",
            Self::DashSmallGap => "dashSmallGap",
            Self::DashDotStroked => "dashDotStroked",
            Self::ThreeDEmboss => "threeDEmboss",
            Self::ThreeDEngrave => "threeDEngrave",
            Self::Outset => "outset",
            Self::Inset => "inset",
            Self::Art(name) => name.as_str(),
        }
    }

    fn parse(value: &str) -> Result<Self> {
        Ok(match value {
            "nil" => Self::Nil,
            "none" => Self::None,
            "single" => Self::Single,
            "thick" => Self::Thick,
            "double" => Self::Double,
            "dotted" => Self::Dotted,
            "dashed" => Self::Dashed,
            "dotDash" => Self::DotDash,
            "dotDotDash" => Self::DotDotDash,
            "triple" => Self::Triple,
            "thinThickSmallGap" => Self::ThinThickSmallGap,
            "thinThickMediumGap" => Self::ThinThickMediumGap,
            "thinThickLargeGap" => Self::ThinThickLargeGap,
            "thickThinSmallGap" => Self::ThickThinSmallGap,
            "thickThinMediumGap" => Self::ThickThinMediumGap,
            "thickThinLargeGap" => Self::ThickThinLargeGap,
            "thinThickThinSmallGap" => Self::ThinThickThinSmallGap,
            "thinThickThinMediumGap" => Self::ThinThickThinMediumGap,
            "thinThickThinLargeGap" => Self::ThinThickThinLargeGap,
            "wave" => Self::Wave,
            "doubleWave" => Self::DoubleWave,
            "dashSmallGap" => Self::DashSmallGap,
            "dashDotStroked" => Self::DashDotStroked,
            "threeDEmboss" => Self::ThreeDEmboss,
            "threeDEngrave" => Self::ThreeDEngrave,
            "outset" => Self::Outset,
            "inset" => Self::Inset,
            art
                if !art.is_empty()
                    && art.len() <= MAX_PAGE_BORDER_ART_NAME_LEN
                    && art.bytes().all(|byte| byte.is_ascii_alphanumeric()) =>
            {
                Self::Art(art.to_string())
            },
            _ => {
                return Err(OoxmlError::InvalidFormat(format!(
                    "invalid page border style '{value}'"
                )));
            },
        })
    }
}

/// Which pages display the section page border (`w:display`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageBorderDisplay {
    AllPages,
    FirstPage,
    NotFirstPage,
}

impl PageBorderDisplay {
    fn as_str(self) -> &'static str {
        match self {
            Self::AllPages => "allPages",
            Self::FirstPage => "firstPage",
            Self::NotFirstPage => "notFirstPage",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "allPages" => Ok(Self::AllPages),
            "firstPage" => Ok(Self::FirstPage),
            "notFirstPage" => Ok(Self::NotFirstPage),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid page border display '{value}'"
            ))),
        }
    }
}

/// Page-border measurement origin (`w:offsetFrom`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageBorderOffsetFrom {
    Page,
    Text,
}

impl PageBorderOffsetFrom {
    fn as_str(self) -> &'static str {
        match self {
            Self::Page => "page",
            Self::Text => "text",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "page" => Ok(Self::Page),
            "text" => Ok(Self::Text),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid page border offset '{value}'"
            ))),
        }
    }
}

/// Whether the page border renders in front of or behind text (`w:zOrder`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageBorderZOrder {
    Front,
    Back,
}

impl PageBorderZOrder {
    fn as_str(self) -> &'static str {
        match self {
            Self::Front => "front",
            Self::Back => "back",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "front" => Ok(Self::Front),
            "back" => Ok(Self::Back),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid page border z-order '{value}'"
            ))),
        }
    }
}

/// One page-border edge (`CT_Border`, ECMA-376 §17.6.16).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionPageBorder {
    pub style: PageBorderStyle,
    /// Border size in eighths of a point for line styles, points for art.
    pub size: Option<u32>,
    /// Border offset space in points (`0..=31`).
    pub space: Option<u32>,
    /// Border color as `RRGGBB` hex or `auto`.
    pub color: Option<String>,
    pub shadow: bool,
    pub frame: bool,
}

/// Section page-border settings (`CT_PageBorders`, ECMA-376 §17.6.16).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionPageBorders {
    pub offset_from: PageBorderOffsetFrom,
    pub z_order: PageBorderZOrder,
    pub display: PageBorderDisplay,
    pub top: Option<SectionPageBorder>,
    pub left: Option<SectionPageBorder>,
    pub bottom: Option<SectionPageBorder>,
    pub right: Option<SectionPageBorder>,
}

impl Default for SectionPageBorders {
    fn default() -> Self {
        Self {
            offset_from: PageBorderOffsetFrom::Page,
            z_order: PageBorderZOrder::Back,
            display: PageBorderDisplay::AllPages,
            top: None,
            left: None,
            bottom: None,
            right: None,
        }
    }
}

/// Paper-source tray codes (`CT_PaperSource`, ECMA-376 §17.6.19).
///
/// Values are printer-driver tray codes; `first` applies to the first page,
/// `other` to the remaining pages of the section.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct SectionPaperSource {
    pub first: Option<u32>,
    pub other: Option<u32>,
}

/// Line-number restart policy (`ST_LineNumberRestart`, ECMA-376 §17.18.49).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineNumberRestart {
    NewPage,
    NewSection,
    Continuous,
}

impl LineNumberRestart {
    fn as_str(self) -> &'static str {
        match self {
            Self::NewPage => "newPage",
            Self::NewSection => "newSection",
            Self::Continuous => "continuous",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "newPage" => Ok(Self::NewPage),
            "newSection" => Ok(Self::NewSection),
            "continuous" => Ok(Self::Continuous),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid line-number restart '{value}'"
            ))),
        }
    }
}

/// Section line-numbering settings (`CT_LineNumber`, ECMA-376 §17.6.12).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct SectionLineNumbering {
    /// Line-number increment; lines are numbered when divisible by this value.
    pub count_by: Option<u32>,
    /// First line number used by the numbering scheme.
    pub start: Option<u32>,
    /// Distance between line numbers and text in twips.
    pub distance: Option<u32>,
    pub restart: Option<LineNumberRestart>,
}

/// Vertical content alignment (`ST_VerticalJc`, ECMA-376 §17.18.100).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionVerticalAlignment {
    Top,
    Center,
    Justified,
    Bottom,
}

impl SectionVerticalAlignment {
    fn as_str(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Center => "center",
            Self::Justified => "both",
            Self::Bottom => "bottom",
        }
    }

    fn parse(value: &str) -> Result<Self> {
        match value {
            "top" => Ok(Self::Top),
            "center" => Ok(Self::Center),
            "both" => Ok(Self::Justified),
            "bottom" => Ok(Self::Bottom),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid section vertical alignment '{value}'"
            ))),
        }
    }
}

/// Typed WordprocessingML `sectPr` properties.
#[derive(Debug, Clone)]
pub struct SectionProperties {
    pub page_width: u32,
    pub page_height: u32,
    pub orientation: PageOrientation,
    pub margin_top: u32,
    pub margin_bottom: u32,
    pub margin_left: u32,
    pub margin_right: u32,
    pub header_distance: u32,
    pub footer_distance: u32,
    pub gutter: u32,
    pub start_type: Option<WdSectionStart>,
    pub headers: Vec<SectionHeaderFooterReference>,
    pub footers: Vec<SectionHeaderFooterReference>,
    pub page_numbering: Option<SectionPageNumbering>,
    pub columns: Option<SectionColumns>,
    pub footnotes: Option<SectionNoteProperties>,
    pub endnotes: Option<SectionNoteProperties>,
    pub text_direction: Option<SectionTextDirection>,
    pub document_grid: Option<SectionDocumentGrid>,
    pub form_protection: bool,
    pub paper_source: Option<SectionPaperSource>,
    pub page_borders: Option<SectionPageBorders>,
    pub line_numbering: Option<SectionLineNumbering>,
    pub vertical_alignment: Option<SectionVerticalAlignment>,
    /// Different first-page header/footer (`w:titlePg`).
    pub title_page: bool,
    /// Right-to-left section layout (`w:bidi`).
    pub bidirectional: bool,
    /// Gutter positioned on the right side (`w:rtlGutter`).
    pub rtl_gutter: bool,
    /// Relationship ID of the printer-settings part (`w:printerSettings`).
    pub printer_settings_relationship_id: Option<String>,
    preserved_unknown_children: Vec<String>,
}

impl Default for SectionProperties {
    fn default() -> Self {
        Self {
            page_width: 12240,
            page_height: 15840,
            orientation: PageOrientation::Portrait,
            margin_top: 1440,
            margin_bottom: 1440,
            margin_left: 1440,
            margin_right: 1440,
            header_distance: 720,
            footer_distance: 720,
            gutter: 0,
            start_type: None,
            headers: Vec::new(),
            footers: Vec::new(),
            page_numbering: None,
            columns: None,
            footnotes: None,
            endnotes: None,
            text_direction: None,
            document_grid: None,
            form_protection: false,
            paper_source: None,
            page_borders: None,
            line_numbering: None,
            vertical_alignment: None,
            title_page: false,
            bidirectional: false,
            rtl_gutter: false,
            printer_settings_relationship_id: None,
            preserved_unknown_children: Vec::new(),
        }
    }
}

impl SectionProperties {
    pub fn a4() -> Self {
        Self {
            page_width: 11906,
            page_height: 16838,
            ..Self::default()
        }
    }

    pub fn letter() -> Self {
        Self::default()
    }

    pub fn legal() -> Self {
        Self {
            page_height: 20160,
            ..Self::default()
        }
    }

    pub fn landscape(mut self) -> Self {
        self.orientation = PageOrientation::Landscape;
        std::mem::swap(&mut self.page_width, &mut self.page_height);
        self
    }

    pub fn margins(mut self, top: f64, bottom: f64, left: f64, right: f64) -> Self {
        self.margin_top = inches_to_twips(top);
        self.margin_bottom = inches_to_twips(bottom);
        self.margin_left = inches_to_twips(left);
        self.margin_right = inches_to_twips(right);
        self
    }

    pub fn with_start_type(mut self, start_type: WdSectionStart) -> Self {
        self.start_type = Some(start_type);
        self
    }

    /// Create or replace a section-owned header part of the selected kind.
    pub fn set_header_part(
        &mut self,
        kind: WdHeaderFooter,
        key: impl Into<String>,
        xml: impl Into<String>,
    ) -> Result<()> {
        self.set_header_footer_part(true, kind, key.into(), xml.into())
    }

    /// Create or replace a section-owned footer part of the selected kind.
    pub fn set_footer_part(
        &mut self,
        kind: WdHeaderFooter,
        key: impl Into<String>,
        xml: impl Into<String>,
    ) -> Result<()> {
        self.set_header_footer_part(false, kind, key.into(), xml.into())
    }

    /// Remove a header reference and any section-owned replacement of that kind.
    pub fn remove_header(&mut self, kind: WdHeaderFooter) -> bool {
        remove_reference(&mut self.headers, kind)
    }

    /// Remove a footer reference and any section-owned replacement of that kind.
    pub fn remove_footer(&mut self, kind: WdHeaderFooter) -> bool {
        remove_reference(&mut self.footers, kind)
    }

    fn set_header_footer_part(
        &mut self,
        header: bool,
        kind: WdHeaderFooter,
        key: String,
        xml: String,
    ) -> Result<()> {
        validate_header_footer_xml(&xml, header)?;
        if key.is_empty() {
            return Err(OoxmlError::InvalidFormat(
                "section header/footer part key is empty".to_string(),
            ));
        }
        let references = if header {
            &mut self.headers
        } else {
            &mut self.footers
        };
        remove_reference(references, kind);
        references.push(SectionHeaderFooterReference::owned(kind, key, xml));
        Ok(())
    }

    pub(crate) fn from_xml(xml: &str) -> Result<Self> {
        let children = direct_children(xml)?;
        let mut properties = Self::default();
        let mut seen = std::collections::HashSet::new();
        let mut last_rank = 0u8;
        for (name, raw) in children {
            if let Some(rank) = section_child_rank(&name) {
                if rank < last_rank {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "section property '{name}' is out of schema order"
                    )));
                }
                last_rank = rank;
            }
            if !seen.insert(name.clone())
                && !matches!(name.as_str(), "headerReference" | "footerReference")
            {
                return Err(OoxmlError::InvalidFormat(format!(
                    "section properties contain duplicate '{name}'"
                )));
            }
            match name.as_str() {
                "headerReference" => properties.headers.push(parse_header_footer(&raw)?),
                "footerReference" => properties.footers.push(parse_header_footer(&raw)?),
                "footnotePr" => properties.footnotes = Some(parse_note_properties(&raw)?),
                "endnotePr" => properties.endnotes = Some(parse_note_properties(&raw)?),
                "type" => {
                    let value = required_attr(&raw, b"val")?;
                    properties.start_type = Some(WdSectionStart::from_xml(&value).ok_or_else(|| {
                        OoxmlError::InvalidFormat(format!("invalid section type '{value}'"))
                    })?);
                },
                "pgSz" => {
                    let attrs = attributes(&raw)?;
                    if let Some(value) = attr(&attrs, "w") {
                        properties.page_width = parse_u32(value, "page width")?;
                    }
                    if let Some(value) = attr(&attrs, "h") {
                        properties.page_height = parse_u32(value, "page height")?;
                    }
                    if let Some(value) = attr(&attrs, "orient") {
                        properties.orientation = PageOrientation::parse(value)?;
                    }
                },
                "pgMar" => {
                    let attrs = attributes(&raw)?;
                    assign_u32(&attrs, "top", &mut properties.margin_top)?;
                    assign_u32(&attrs, "bottom", &mut properties.margin_bottom)?;
                    assign_u32(&attrs, "left", &mut properties.margin_left)?;
                    assign_u32(&attrs, "right", &mut properties.margin_right)?;
                    assign_u32(&attrs, "header", &mut properties.header_distance)?;
                    assign_u32(&attrs, "footer", &mut properties.footer_distance)?;
                    assign_u32(&attrs, "gutter", &mut properties.gutter)?;
                },
                "pgNumType" => properties.page_numbering = Some(parse_page_numbering(&raw)?),
                "paperSrc" => {
                    let attrs = attributes(&raw)?;
                    properties.paper_source = Some(SectionPaperSource {
                        first: attr(&attrs, "first")
                            .map(|value| parse_u32(value, "first paper source"))
                            .transpose()?,
                        other: attr(&attrs, "other")
                            .map(|value| parse_u32(value, "other paper source"))
                            .transpose()?,
                    });
                },
                "pgBorders" => properties.page_borders = Some(parse_page_borders(&raw)?),
                "lnNumType" => {
                    properties.line_numbering = Some(parse_line_numbering(&raw)?);
                },
                "cols" => properties.columns = Some(parse_columns(&raw)?),
                "formProt" => properties.form_protection = parse_on_off(&raw)?,
                "vAlign" => {
                    properties.vertical_alignment = Some(SectionVerticalAlignment::parse(
                        &required_attr(&raw, b"val")?,
                    )?);
                },
                "titlePg" => properties.title_page = parse_on_off(&raw)?,
                "textDirection" => {
                    properties.text_direction = Some(SectionTextDirection::parse(&required_attr(
                        &raw, b"val",
                    )?)?);
                },
                "bidi" => properties.bidirectional = parse_on_off(&raw)?,
                "rtlGutter" => properties.rtl_gutter = parse_on_off(&raw)?,
                "docGrid" => properties.document_grid = Some(parse_grid(&raw)?),
                "printerSettings" => {
                    properties.printer_settings_relationship_id =
                        Some(required_attr(&raw, b"id")?);
                },
                _ => properties.preserved_unknown_children.push(raw),
            }
        }
        properties.validate()?;
        Ok(properties)
    }

    pub(crate) fn validate(&self) -> Result<()> {
        if self.page_width == 0 || self.page_height == 0 {
            return Err(OoxmlError::InvalidFormat(
                "section page dimensions must be nonzero".to_string(),
            ));
        }
        for references in [&self.headers, &self.footers] {
            let mut kinds = std::collections::HashSet::new();
            for reference in references {
                if !kinds.insert(reference.kind) {
                    return Err(OoxmlError::InvalidFormat(
                        "section has duplicate header/footer reference type".to_string(),
                    ));
                }
                if reference.relationship_id.as_deref() == Some("") {
                    return Err(OoxmlError::InvalidFormat(
                        "section header/footer relationship ID is empty".to_string(),
                    ));
                }
                if reference.relationship_id.is_some() && reference.part.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "section header/footer cannot be both existing and owned".to_string(),
                    ));
                }
                if let Some(part) = &reference.part
                    && (part.key.is_empty() || part.xml.is_empty())
                {
                    return Err(OoxmlError::InvalidFormat(
                        "section header/footer part key and XML must be non-empty".to_string(),
                    ));
                }
                if let Some(part) = &reference.part {
                    validate_header_footer_xml(
                        &part.xml,
                        std::ptr::eq(references, &self.headers),
                    )?;
                }
            }
        }
        if let Some(columns) = &self.columns {
            if columns.count == 0 || columns.count > 45 {
                return Err(OoxmlError::InvalidFormat(
                    "section column count must be in 1..=45".to_string(),
                ));
            }
            if !columns.equal_width && usize::from(columns.count) != columns.columns.len() {
                return Err(OoxmlError::InvalidFormat(
                    "unequal section columns require one width per column".to_string(),
                ));
            }
        }
        if let Some(borders) = &self.page_borders {
            for border in [&borders.top, &borders.left, &borders.bottom, &borders.right]
                .into_iter()
                .flatten()
            {
                if let Some(size) = border.size {
                    let max = match border.style {
                        PageBorderStyle::Art(_) => MAX_PAGE_BORDER_ART_SIZE,
                        _ => MAX_PAGE_BORDER_LINE_SIZE,
                    };
                    if size > max {
                        return Err(OoxmlError::InvalidFormat(format!(
                            "page border size {size} exceeds the {max} limit"
                        )));
                    }
                }
                if let Some(space) = border.space
                    && space > MAX_PAGE_BORDER_SPACE
                {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "page border space {space} exceeds the {MAX_PAGE_BORDER_SPACE} limit"
                    )));
                }
                if let Some(color) = &border.color
                    && !(color == "auto"
                        || (color.len() == 6
                            && color.bytes().all(|byte| byte.is_ascii_hexdigit())))
                {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "invalid page border color '{color}'"
                    )));
                }
            }
        }
        if self.printer_settings_relationship_id.as_deref() == Some("") {
            return Err(OoxmlError::InvalidFormat(
                "section printer-settings relationship ID is empty".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn write_xml(
        &self,
        xml: &mut String,
        rels: Option<&super::relmap::RelationshipMapper>,
    ) -> Result<()> {
        self.validate()?;
        xml.push_str("<w:sectPr>");
        write_references(xml, "headerReference", &self.headers, rels, true)?;
        write_references(xml, "footerReference", &self.footers, rels, false)?;
        if let Some(note) = &self.footnotes {
            write_note_properties(xml, "footnotePr", note)?;
        } else if rels.is_some_and(|rels| rels.get_footnotes_id().is_some()) {
            xml.push_str("<w:footnotePr><w:numFmt w:val=\"decimal\"/></w:footnotePr>");
        }
        if let Some(note) = &self.endnotes {
            write_note_properties(xml, "endnotePr", note)?;
        } else if rels.is_some_and(|rels| rels.get_endnotes_id().is_some()) {
            xml.push_str("<w:endnotePr><w:numFmt w:val=\"decimal\"/></w:endnotePr>");
        }
        if let Some(start_type) = self.start_type {
            write!(xml, "<w:type w:val=\"{}\"/>", start_type.to_xml())
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        }
        write!(
            xml,
            "<w:pgSz w:w=\"{}\" w:h=\"{}\" w:orient=\"{}\"/>",
            self.page_width,
            self.page_height,
            self.orientation.as_str()
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        write!(
            xml,
            "<w:pgMar w:top=\"{}\" w:right=\"{}\" w:bottom=\"{}\" w:left=\"{}\" w:header=\"{}\" w:footer=\"{}\" w:gutter=\"{}\"/>",
            self.margin_top,
            self.margin_right,
            self.margin_bottom,
            self.margin_left,
            self.header_distance,
            self.footer_distance,
            self.gutter
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if let Some(paper_source) = &self.paper_source {
            xml.push_str("<w:paperSrc");
            if let Some(first) = paper_source.first {
                write!(xml, " w:first=\"{first}\"")
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            }
            if let Some(other) = paper_source.other {
                write!(xml, " w:other=\"{other}\"")
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            }
            xml.push_str("/>");
        }
        if let Some(borders) = &self.page_borders {
            write_page_borders(xml, borders)?;
        }
        if let Some(numbering) = &self.line_numbering {
            write_line_numbering(xml, numbering)?;
        }
        if let Some(numbering) = &self.page_numbering {
            write_page_numbering(xml, numbering)?;
        }
        if let Some(columns) = &self.columns {
            write_columns(xml, columns)?;
        }
        if self.form_protection {
            xml.push_str("<w:formProt/>");
        }
        if let Some(alignment) = self.vertical_alignment {
            write!(xml, "<w:vAlign w:val=\"{}\"/>", alignment.as_str())
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        }
        if self.title_page {
            xml.push_str("<w:titlePg/>");
        }
        if let Some(direction) = self.text_direction {
            write!(xml, "<w:textDirection w:val=\"{}\"/>", direction.as_str())
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        }
        if self.bidirectional {
            xml.push_str("<w:bidi/>");
        }
        if self.rtl_gutter {
            xml.push_str("<w:rtlGutter/>");
        }
        if let Some(grid) = &self.document_grid {
            write_grid(xml, grid)?;
        }
        if let Some(id) = &self.printer_settings_relationship_id {
            write!(xml, "<w:printerSettings r:id=\"{}\"/>", escape(id))
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        }
        for child in &self.preserved_unknown_children {
            xml.push_str(child);
        }
        xml.push_str("</w:sectPr>");
        Ok(())
    }
}

fn section_child_rank(name: &str) -> Option<u8> {
    match name {
        "headerReference" => Some(0),
        "footerReference" => Some(1),
        "footnotePr" => Some(2),
        "endnotePr" => Some(3),
        "type" => Some(4),
        "pgSz" => Some(5),
        "pgMar" => Some(6),
        "paperSrc" => Some(7),
        "pgBorders" => Some(8),
        "lnNumType" => Some(9),
        "pgNumType" => Some(10),
        "cols" => Some(11),
        "formProt" => Some(12),
        "vAlign" => Some(13),
        "titlePg" => Some(14),
        "textDirection" => Some(15),
        "bidi" => Some(16),
        "rtlGutter" => Some(17),
        "docGrid" => Some(18),
        "printerSettings" => Some(19),
        _ => None,
    }
}

fn remove_reference(
    references: &mut Vec<SectionHeaderFooterReference>,
    kind: WdHeaderFooter,
) -> bool {
    let length = references.len();
    references.retain(|reference| reference.kind != kind);
    references.len() != length
}

fn validate_header_footer_xml(xml: &str, header: bool) -> Result<()> {
    use quick_xml::reader::NsReader;
    let mut reader = NsReader::from_str(xml);
    let mut depth = 0usize;
    let mut root = false;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    let expected = if header { b"hdr".as_slice() } else { b"ftr".as_slice() };
                    if root
                        || !crate::docx::namespace::is_wordprocessing_namespace(&namespace)
                        || element.local_name().as_ref() != expected
                    {
                        return Err(OoxmlError::InvalidFormat(
                            "section header/footer XML has an invalid root".to_string(),
                        ));
                    }
                    root = true;
                }
                depth += 1;
            },
            Event::Empty(element) if depth == 0 => {
                let expected = if header { b"hdr".as_slice() } else { b"ftr".as_slice() };
                if root
                    || !crate::docx::namespace::is_wordprocessing_namespace(&namespace)
                    || element.local_name().as_ref() != expected
                {
                    return Err(OoxmlError::InvalidFormat(
                        "section header/footer XML has an invalid root".to_string(),
                    ));
                }
                root = true;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid header/footer XML nesting".to_string())
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root || depth != 0 {
        return Err(OoxmlError::InvalidFormat(
            "unterminated section header/footer XML".to_string(),
        ));
    }
    Ok(())
}

fn inches_to_twips(inches: f64) -> u32 {
    (inches * 1440.0).round().max(0.0) as u32
}

fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
}

fn direct_children(xml: &str) -> Result<Vec<(String, String)>> {
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut child: Option<(String, usize, usize)> = None;
    let mut children = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| OoxmlError::InvalidFormat("section XML offset overflow".into()))?;
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| OoxmlError::InvalidFormat("section XML offset overflow".into()))?;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if root_seen || element.local_name().as_ref() != b"sectPr" {
                        return Err(OoxmlError::InvalidFormat(
                            "section properties have an invalid root".into(),
                        ));
                    }
                    root_seen = true;
                } else if depth == 1 {
                    child = Some((
                        String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
                        start,
                        1,
                    ));
                } else if let Some((_, _, child_depth)) = child.as_mut() {
                    *child_depth += 1;
                }
                depth += 1;
            },
            Event::Empty(element) if depth == 1 => {
                let name = String::from_utf8_lossy(element.local_name().as_ref()).into_owned();
                children.push((name, xml[start..end].to_string()));
            },
            Event::End(_) => {
                if let Some((_, _, child_depth)) = child.as_mut() {
                    *child_depth -= 1;
                    if *child_depth == 0 {
                        let (name, child_start, _) = child.take().expect("present");
                        children.push((name, xml[child_start..end].to_string()));
                    }
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid section XML nesting".into())
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || depth != 0 {
        return Err(OoxmlError::InvalidFormat(
            "unterminated section properties".into(),
        ));
    }
    Ok(children)
}

fn attributes(xml: &str) -> Result<Vec<(String, String)>> {
    let mut reader = Reader::from_str(xml);
    let element = loop {
        match reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
        {
            Event::Start(element) | Event::Empty(element) => break element,
            Event::Eof => {
                return Err(OoxmlError::InvalidFormat(
                    "section property has no element".into(),
                ));
            },
            _ => {},
        }
    };
    let mut result = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let name = String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
        if result.iter().any(|(candidate, _)| candidate == &name) {
            return Err(OoxmlError::InvalidFormat(format!(
                "duplicate section property attribute '{name}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        result.push((name, value));
    }
    Ok(result)
}

fn attr<'a>(attrs: &'a [(String, String)], name: &str) -> Option<&'a str> {
    attrs
        .iter()
        .find_map(|(candidate, value)| (candidate == name).then_some(value.as_str()))
}

fn required_attr(xml: &str, name: &[u8]) -> Result<String> {
    let name = String::from_utf8_lossy(name);
    let attrs = attributes(xml)?;
    attr(&attrs, &name)
        .map(ToOwned::to_owned)
        .ok_or_else(|| OoxmlError::InvalidFormat(format!("missing section attribute '{name}'")))
}

fn parse_u32(value: &str, description: &str) -> Result<u32> {
    value.parse().map_err(|_| {
        OoxmlError::InvalidFormat(format!("invalid {description} value '{value}'"))
    })
}

fn assign_u32(attrs: &[(String, String)], name: &str, slot: &mut u32) -> Result<()> {
    if let Some(value) = attr(attrs, name) {
        *slot = parse_u32(value, name)?;
    }
    Ok(())
}

fn parse_header_footer(xml: &str) -> Result<SectionHeaderFooterReference> {
    let kind = WdHeaderFooter::from_xml(&required_attr(xml, b"type")?).ok_or_else(|| {
        OoxmlError::InvalidFormat("invalid section header/footer type".to_string())
    })?;
    Ok(SectionHeaderFooterReference {
        kind,
        relationship_id: Some(required_attr(xml, b"id")?),
        part: None,
    })
}

fn parse_page_numbering(xml: &str) -> Result<SectionPageNumbering> {
    let attrs = attributes(xml)?;
    Ok(SectionPageNumbering {
        format: attr(&attrs, "fmt")
            .map(PageNumberFormat::parse)
            .transpose()?
            .unwrap_or(PageNumberFormat::Decimal),
        start: attr(&attrs, "start")
            .map(|value| parse_u32(value, "page number start"))
            .transpose()?,
        chapter_style: attr(&attrs, "chapStyle")
            .map(|value| value.parse::<u8>().map_err(|_| OoxmlError::InvalidFormat("invalid chapter style".into())))
            .transpose()?,
        chapter_separator: attr(&attrs, "chapSep").map(ToOwned::to_owned),
    })
}

fn parse_columns(xml: &str) -> Result<SectionColumns> {
    let attrs = attributes(xml)?;
    let mut columns = SectionColumns {
        equal_width: attr(&attrs, "equalWidth").is_none_or(|value| value != "0" && value != "false"),
        count: attr(&attrs, "num")
            .map(|value| value.parse::<u16>().map_err(|_| OoxmlError::InvalidFormat("invalid section column count".into())))
            .transpose()?
            .unwrap_or(1),
        space: attr(&attrs, "space")
            .map(|value| parse_u32(value, "column space"))
            .transpose()?,
        separator: attr(&attrs, "sep").is_some_and(|value| value == "1" || value == "true"),
        columns: Vec::new(),
    };
    for (name, raw) in direct_nested_children(xml)? {
        if name != "col" {
            return Err(OoxmlError::InvalidFormat(format!("invalid child '{name}' in section columns")));
        }
        let attrs = attributes(&raw)?;
        columns.columns.push(SectionColumn {
            width: parse_u32(attr(&attrs, "w").ok_or_else(|| OoxmlError::InvalidFormat("section column omits width".into()))?, "column width")?,
            space: attr(&attrs, "space").map(|value| parse_u32(value, "column space")).transpose()?,
        });
    }
    Ok(columns)
}

fn direct_nested_children(xml: &str) -> Result<Vec<(String, String)>> {
    let open_end = xml.find('>').ok_or_else(|| OoxmlError::InvalidFormat("invalid section property".into()))?;
    let close = xml.rfind("</").unwrap_or(xml.len());
    if close <= open_end + 1 {
        return Ok(Vec::new());
    }
    direct_children(&format!("<w:sectPr>{}</w:sectPr>", &xml[open_end + 1..close]))
}

fn parse_note_properties(xml: &str) -> Result<SectionNoteProperties> {
    let mut result = SectionNoteProperties {
        format: PageNumberFormat::Decimal,
        start: None,
        restart: None,
        position: None,
    };
    for (name, raw) in direct_nested_children(xml)? {
        let value = required_attr(&raw, b"val")?;
        match name.as_str() {
            "numFmt" => result.format = PageNumberFormat::parse(&value)?,
            "numStart" => result.start = Some(parse_u32(&value, "note number start")?),
            "numRestart" => result.restart = Some(NoteNumberRestart::parse(&value)?),
            "pos" => result.position = Some(value),
            _ => return Err(OoxmlError::InvalidFormat(format!("invalid note property '{name}'"))),
        }
    }
    Ok(result)
}

fn parse_grid(xml: &str) -> Result<SectionDocumentGrid> {
    let attrs = attributes(xml)?;
    Ok(SectionDocumentGrid {
        grid_type: attr(&attrs, "type")
            .map(DocumentGridType::parse)
            .transpose()?
            .unwrap_or(DocumentGridType::Default),
        line_pitch: attr(&attrs, "linePitch").map(|value| parse_u32(value, "grid line pitch")).transpose()?,
        char_space: attr(&attrs, "charSpace")
            .map(|value| value.parse::<i32>().map_err(|_| OoxmlError::InvalidFormat("invalid grid character space".into())))
            .transpose()?,
    })
}

fn parse_page_borders(xml: &str) -> Result<SectionPageBorders> {
    let attrs = attributes(xml)?;
    let mut borders = SectionPageBorders {
        offset_from: attr(&attrs, "offsetFrom")
            .map(PageBorderOffsetFrom::parse)
            .transpose()?
            .unwrap_or(PageBorderOffsetFrom::Page),
        z_order: attr(&attrs, "zOrder")
            .map(PageBorderZOrder::parse)
            .transpose()?
            .unwrap_or(PageBorderZOrder::Back),
        display: attr(&attrs, "display")
            .map(PageBorderDisplay::parse)
            .transpose()?
            .unwrap_or(PageBorderDisplay::AllPages),
        ..SectionPageBorders::default()
    };
    for (name, raw) in direct_nested_children(xml)? {
        let edge = match name.as_str() {
            "top" => &mut borders.top,
            "left" => &mut borders.left,
            "bottom" => &mut borders.bottom,
            "right" => &mut borders.right,
            _ => {
                return Err(OoxmlError::InvalidFormat(format!(
                    "invalid child '{name}' in section page borders"
                )));
            },
        };
        if edge.is_some() {
            return Err(OoxmlError::InvalidFormat(format!(
                "duplicate '{name}' page border edge"
            )));
        }
        *edge = Some(parse_page_border(&raw)?);
    }
    Ok(borders)
}

fn parse_page_border(xml: &str) -> Result<SectionPageBorder> {
    let attrs = attributes(xml)?;
    let on_off = |name: &str| {
        attr(&attrs, name).is_some_and(|value| matches!(value, "1" | "true" | "on"))
    };
    Ok(SectionPageBorder {
        style: PageBorderStyle::parse(
            attr(&attrs, "val")
                .ok_or_else(|| OoxmlError::InvalidFormat("page border omits style".into()))?,
        )?,
        size: attr(&attrs, "sz")
            .map(|value| parse_u32(value, "page border size"))
            .transpose()?,
        space: attr(&attrs, "space")
            .map(|value| parse_u32(value, "page border space"))
            .transpose()?,
        color: attr(&attrs, "color").map(ToOwned::to_owned),
        shadow: on_off("shadow"),
        frame: on_off("frame"),
    })
}

fn parse_line_numbering(xml: &str) -> Result<SectionLineNumbering> {
    let attrs = attributes(xml)?;
    Ok(SectionLineNumbering {
        count_by: attr(&attrs, "countBy")
            .map(|value| parse_u32(value, "line-number increment"))
            .transpose()?,
        start: attr(&attrs, "start")
            .map(|value| parse_u32(value, "line-number start"))
            .transpose()?,
        distance: attr(&attrs, "distance")
            .map(|value| parse_u32(value, "line-number distance"))
            .transpose()?,
        restart: attr(&attrs, "restart")
            .map(LineNumberRestart::parse)
            .transpose()?,
    })
}

fn parse_on_off(xml: &str) -> Result<bool> {
    let attrs = attributes(xml)?;
    Ok(attr(&attrs, "val").is_none_or(|value| matches!(value, "1" | "true" | "on")))
}

fn write_references(
    xml: &mut String,
    element: &str,
    references: &[SectionHeaderFooterReference],
    rels: Option<&super::relmap::RelationshipMapper>,
    header: bool,
) -> Result<()> {
    if references.is_empty() {
        let managed = rels.and_then(|rels| if header { rels.get_header_id() } else { rels.get_footer_id() });
        if let Some(id) = managed {
            write!(xml, "<w:{element} w:type=\"default\" r:id=\"{}\"/>", escape(id))
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        }
        return Ok(());
    }
    for reference in references {
        let managed = rels.and_then(|rels| if header { rels.get_header_id() } else { rels.get_footer_id() });
        let owned = reference.part.as_ref().and_then(|part| rels.and_then(|rels| rels.get_section_header_footer_id(&part.key)));
        let id = reference.relationship_id.as_deref().or(owned).or(managed).ok_or_else(|| {
            OoxmlError::InvalidFormat(format!("section {element} has no relationship ID"))
        })?;
        write!(xml, "<w:{element} w:type=\"{}\" r:id=\"{}\"/>", reference.kind.to_xml(), escape(id))
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    Ok(())
}

fn write_note_properties(xml: &mut String, element: &str, note: &SectionNoteProperties) -> Result<()> {
    write!(xml, "<w:{element}><w:numFmt w:val=\"{}\"/>", note.format.as_str())
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    if let Some(start) = note.start { write!(xml, "<w:numStart w:val=\"{start}\"/>").map_err(|error| OoxmlError::Xml(error.to_string()))?; }
    if let Some(restart) = note.restart { write!(xml, "<w:numRestart w:val=\"{}\"/>", restart.as_str()).map_err(|error| OoxmlError::Xml(error.to_string()))?; }
    if let Some(position) = &note.position { write!(xml, "<w:pos w:val=\"{}\"/>", escape(position)).map_err(|error| OoxmlError::Xml(error.to_string()))?; }
    write!(xml, "</w:{element}>").map_err(|error| OoxmlError::Xml(error.to_string()))
}

fn write_page_numbering(xml: &mut String, numbering: &SectionPageNumbering) -> Result<()> {
    write!(xml, "<w:pgNumType w:fmt=\"{}\"", numbering.format.as_str()).map_err(|error| OoxmlError::Xml(error.to_string()))?;
    if let Some(start) = numbering.start { write!(xml, " w:start=\"{start}\"").map_err(|error| OoxmlError::Xml(error.to_string()))?; }
    if let Some(style) = numbering.chapter_style { write!(xml, " w:chapStyle=\"{style}\"").map_err(|error| OoxmlError::Xml(error.to_string()))?; }
    if let Some(separator) = &numbering.chapter_separator { write!(xml, " w:chapSep=\"{}\"", escape(separator)).map_err(|error| OoxmlError::Xml(error.to_string()))?; }
    xml.push_str("/>");
    Ok(())
}

fn write_columns(xml: &mut String, columns: &SectionColumns) -> Result<()> {
    write!(xml, "<w:cols w:equalWidth=\"{}\" w:num=\"{}\"", if columns.equal_width { 1 } else { 0 }, columns.count).map_err(|error| OoxmlError::Xml(error.to_string()))?;
    if let Some(space) = columns.space { write!(xml, " w:space=\"{space}\"").map_err(|error| OoxmlError::Xml(error.to_string()))?; }
    if columns.separator { xml.push_str(" w:sep=\"1\""); }
    if columns.columns.is_empty() { xml.push_str("/>"); } else {
        xml.push('>');
        for column in &columns.columns {
            write!(xml, "<w:col w:w=\"{}\"", column.width).map_err(|error| OoxmlError::Xml(error.to_string()))?;
            if let Some(space) = column.space { write!(xml, " w:space=\"{space}\"").map_err(|error| OoxmlError::Xml(error.to_string()))?; }
            xml.push_str("/>");
        }
        xml.push_str("</w:cols>");
    }
    Ok(())
}

fn write_grid(xml: &mut String, grid: &SectionDocumentGrid) -> Result<()> {
    write!(xml, "<w:docGrid w:type=\"{}\"", grid.grid_type.as_str()).map_err(|error| OoxmlError::Xml(error.to_string()))?;
    if let Some(pitch) = grid.line_pitch { write!(xml, " w:linePitch=\"{pitch}\"").map_err(|error| OoxmlError::Xml(error.to_string()))?; }
    if let Some(space) = grid.char_space { write!(xml, " w:charSpace=\"{space}\"").map_err(|error| OoxmlError::Xml(error.to_string()))?; }
    xml.push_str("/>");
    Ok(())
}

fn write_page_borders(xml: &mut String, borders: &SectionPageBorders) -> Result<()> {
    write!(
        xml,
        "<w:pgBorders w:offsetFrom=\"{}\" w:zOrder=\"{}\" w:display=\"{}\"",
        borders.offset_from.as_str(),
        borders.z_order.as_str(),
        borders.display.as_str()
    )
    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    let edges = [
        ("top", &borders.top),
        ("left", &borders.left),
        ("bottom", &borders.bottom),
        ("right", &borders.right),
    ];
    if edges.iter().all(|(_, edge)| edge.is_none()) {
        xml.push_str("/>");
        return Ok(());
    }
    xml.push('>');
    for (name, edge) in edges {
        if let Some(border) = edge {
            write_page_border(xml, name, border)?;
        }
    }
    xml.push_str("</w:pgBorders>");
    Ok(())
}

fn write_page_border(xml: &mut String, name: &str, border: &SectionPageBorder) -> Result<()> {
    write!(xml, "<w:{name} w:val=\"{}\"", escape(border.style.as_str()))
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    if let Some(size) = border.size {
        write!(xml, " w:sz=\"{size}\"").map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if let Some(space) = border.space {
        write!(xml, " w:space=\"{space}\"")
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if let Some(color) = &border.color {
        write!(xml, " w:color=\"{}\"", escape(color))
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if border.shadow {
        xml.push_str(" w:shadow=\"1\"");
    }
    if border.frame {
        xml.push_str(" w:frame=\"1\"");
    }
    xml.push_str("/>");
    Ok(())
}

fn write_line_numbering(xml: &mut String, numbering: &SectionLineNumbering) -> Result<()> {
    xml.push_str("<w:lnNumType");
    if let Some(count_by) = numbering.count_by {
        write!(xml, " w:countBy=\"{count_by}\"")
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if let Some(start) = numbering.start {
        write!(xml, " w:start=\"{start}\"")
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if let Some(distance) = numbering.distance {
        write!(xml, " w:distance=\"{distance}\"")
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if let Some(restart) = numbering.restart {
        write!(xml, " w:restart=\"{}\"", restart.as_str())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    xml.push_str("/>");
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn typed_section_round_trips_and_preserves_unknown_children() {
        let xml = r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:test"><w:type w:val="continuous"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="100" w:right="200" w:bottom="300" w:left="400" w:header="500" w:footer="600" w:gutter="50"/><w:cols w:num="2" w:space="720"/><x:ext x:v="keep"/></w:sectPr>"#;
        let mut section = SectionProperties::from_xml(xml).unwrap();
        assert_eq!(section.start_type, Some(WdSectionStart::Continuous));
        section.margin_left = 900;
        let mut output = String::new();
        section.write_xml(&mut output, None).unwrap();
        assert!(output.contains("w:left=\"900\""));
        assert!(output.contains("<x:ext x:v=\"keep\"/>"));
    }

    #[test]
    fn rejects_duplicate_and_invalid_section_properties() {
        assert!(SectionProperties::from_xml("<w:sectPr><w:type w:val=\"nextPage\"/><w:type w:val=\"continuous\"/></w:sectPr>").is_err());
        let mut section = SectionProperties::default();
        section.columns = Some(SectionColumns { count: 0, ..SectionColumns::default() });
        assert!(section.validate().is_err());
    }

    #[test]
    fn page_layout_properties_round_trip() {
        let xml = r#"<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:paperSrc w:first="1" w:other="260"/><w:pgBorders w:offsetFrom="text" w:zOrder="front" w:display="firstPage"><w:top w:val="double" w:sz="8" w:space="24" w:color="FF0000" w:shadow="1"/><w:bottom w:val="starsTop" w:sz="120" w:space="4" w:color="auto" w:frame="true"/></w:pgBorders><w:lnNumType w:countBy="5" w:start="0" w:distance="240" w:restart="newSection"/><w:pgNumType w:fmt="lowerRoman" w:start="3"/><w:cols w:num="2"/><w:vAlign w:val="both"/><w:titlePg/><w:bidi/><w:rtlGutter/><w:printerSettings r:id="rId9"/></w:sectPr>"#;
        let section = SectionProperties::from_xml(xml).unwrap();
        assert_eq!(
            section.paper_source,
            Some(SectionPaperSource {
                first: Some(1),
                other: Some(260),
            })
        );
        let borders = section.page_borders.as_ref().unwrap();
        assert_eq!(borders.offset_from, PageBorderOffsetFrom::Text);
        assert_eq!(borders.z_order, PageBorderZOrder::Front);
        assert_eq!(borders.display, PageBorderDisplay::FirstPage);
        let top = borders.top.as_ref().unwrap();
        assert_eq!(top.style, PageBorderStyle::Double);
        assert_eq!(top.size, Some(8));
        assert_eq!(top.space, Some(24));
        assert_eq!(top.color.as_deref(), Some("FF0000"));
        assert!(top.shadow);
        assert!(!top.frame);
        let bottom = borders.bottom.as_ref().unwrap();
        assert_eq!(bottom.style, PageBorderStyle::Art("starsTop".to_string()));
        assert_eq!(bottom.size, Some(120));
        assert_eq!(bottom.color.as_deref(), Some("auto"));
        assert!(bottom.frame);
        assert!(borders.left.is_none() && borders.right.is_none());
        assert_eq!(
            section.line_numbering,
            Some(SectionLineNumbering {
                count_by: Some(5),
                start: Some(0),
                distance: Some(240),
                restart: Some(LineNumberRestart::NewSection),
            })
        );
        assert_eq!(
            section.vertical_alignment,
            Some(SectionVerticalAlignment::Justified)
        );
        assert!(section.title_page);
        assert!(section.bidirectional);
        assert!(section.rtl_gutter);
        assert_eq!(
            section.printer_settings_relationship_id.as_deref(),
            Some("rId9")
        );

        let mut output = String::new();
        section.write_xml(&mut output, None).unwrap();
        let reparsed = SectionProperties::from_xml(&output).unwrap();
        assert_eq!(reparsed.paper_source, section.paper_source);
        assert_eq!(reparsed.page_borders, section.page_borders);
        assert_eq!(reparsed.line_numbering, section.line_numbering);
        assert_eq!(reparsed.vertical_alignment, section.vertical_alignment);
        assert!(reparsed.title_page && reparsed.bidirectional && reparsed.rtl_gutter);
        assert_eq!(
            reparsed.printer_settings_relationship_id,
            section.printer_settings_relationship_id
        );
    }

    #[test]
    fn page_layout_defaults_and_empty_edges() {
        let xml = r#"<w:sectPr><w:pgBorders/><w:lnNumType w:countBy="2"/></w:sectPr>"#;
        let section = SectionProperties::from_xml(xml).unwrap();
        let borders = section.page_borders.as_ref().unwrap();
        assert_eq!(borders.offset_from, PageBorderOffsetFrom::Page);
        assert_eq!(borders.z_order, PageBorderZOrder::Back);
        assert_eq!(borders.display, PageBorderDisplay::AllPages);
        assert_eq!(
            section.line_numbering,
            Some(SectionLineNumbering {
                count_by: Some(2),
                ..SectionLineNumbering::default()
            })
        );
        let mut output = String::new();
        section.write_xml(&mut output, None).unwrap();
        assert!(output.contains("<w:pgBorders w:offsetFrom=\"page\" w:zOrder=\"back\" w:display=\"allPages\"/>"));
        assert!(output.contains("<w:lnNumType w:countBy=\"2\"/>"));
    }

    #[test]
    fn page_border_style_enum_round_trips() {
        let styles = [
            PageBorderStyle::Nil,
            PageBorderStyle::None,
            PageBorderStyle::Single,
            PageBorderStyle::Thick,
            PageBorderStyle::Double,
            PageBorderStyle::Dotted,
            PageBorderStyle::Dashed,
            PageBorderStyle::DotDash,
            PageBorderStyle::DotDotDash,
            PageBorderStyle::Triple,
            PageBorderStyle::ThinThickSmallGap,
            PageBorderStyle::ThinThickMediumGap,
            PageBorderStyle::ThinThickLargeGap,
            PageBorderStyle::ThickThinSmallGap,
            PageBorderStyle::ThickThinMediumGap,
            PageBorderStyle::ThickThinLargeGap,
            PageBorderStyle::ThinThickThinSmallGap,
            PageBorderStyle::ThinThickThinMediumGap,
            PageBorderStyle::ThinThickThinLargeGap,
            PageBorderStyle::Wave,
            PageBorderStyle::DoubleWave,
            PageBorderStyle::DashSmallGap,
            PageBorderStyle::DashDotStroked,
            PageBorderStyle::ThreeDEmboss,
            PageBorderStyle::ThreeDEngrave,
            PageBorderStyle::Outset,
            PageBorderStyle::Inset,
        ];
        for style in &styles {
            assert_eq!(&PageBorderStyle::parse(style.as_str()).unwrap(), style);
        }
        assert_eq!(
            PageBorderStyle::parse("apples").unwrap(),
            PageBorderStyle::Art("apples".to_string())
        );
        assert!(PageBorderStyle::parse("not a style!").is_err());
        assert!(PageBorderStyle::parse("").is_err());
    }

    #[test]
    fn rejects_malformed_page_layout_properties() {
        // Unknown enum tokens.
        assert!(SectionProperties::from_xml("<w:sectPr><w:vAlign w:val=\"diagonal\"/></w:sectPr>").is_err());
        assert!(SectionProperties::from_xml("<w:sectPr><w:lnNumType w:restart=\"weekly\"/></w:sectPr>").is_err());
        assert!(SectionProperties::from_xml("<w:sectPr><w:pgBorders w:offsetFrom=\"margin\"/></w:sectPr>").is_err());
        assert!(SectionProperties::from_xml("<w:sectPr><w:pgBorders><w:top w:val=\"single\"/><w:top w:val=\"thick\"/></w:pgBorders></w:sectPr>").is_err());
        assert!(SectionProperties::from_xml("<w:sectPr><w:pgBorders><w:diagonal w:val=\"single\"/></w:pgBorders></w:sectPr>").is_err());
        assert!(SectionProperties::from_xml("<w:sectPr><w:pgBorders><w:top/></w:pgBorders></w:sectPr>").is_err());
        // Out-of-bounds values rejected through validation.
        assert!(SectionProperties::from_xml("<w:sectPr><w:pgBorders><w:top w:val=\"single\" w:sz=\"97\"/></w:pgBorders></w:sectPr>").is_err());
        assert!(SectionProperties::from_xml("<w:sectPr><w:pgBorders><w:top w:val=\"single\" w:space=\"32\"/></w:pgBorders></w:sectPr>").is_err());
        assert!(SectionProperties::from_xml("<w:sectPr><w:pgBorders><w:top w:val=\"starsTop\" w:sz=\"1639\"/></w:pgBorders></w:sectPr>").is_err());
        assert!(SectionProperties::from_xml("<w:sectPr><w:pgBorders><w:top w:val=\"single\" w:color=\"FFF\"/></w:pgBorders></w:sectPr>").is_err());
        // Schema-order violations.
        assert!(SectionProperties::from_xml("<w:sectPr><w:pgNumType w:fmt=\"decimal\"/><w:lnNumType w:countBy=\"5\"/></w:sectPr>").is_err());
        assert!(SectionProperties::from_xml("<w:sectPr><w:printerSettings r:id=\"rId1\"/><w:docGrid/></w:sectPr>").is_err());
        // Empty relationship ID rejected through validation.
        let mut section = SectionProperties::default();
        section.printer_settings_relationship_id = Some(String::new());
        assert!(section.validate().is_err());
    }
}
