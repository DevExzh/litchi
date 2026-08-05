/// Section - document section with page setup and layout properties.
use crate::error::Result;
use litchi_core::unit::{EMUS_PER_CM, EMUS_PER_INCH, EMUS_PER_PT, EMUS_PER_TWIP, emu_to_twip_i64};
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};
use std::fmt;

/// Page layout orientation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Orientation {
    /// Portrait orientation.
    Portrait = 0,
    /// Landscape orientation.
    Landscape = 1,
}

impl Orientation {
    /// Convert the orientation to its XML attribute value.
    #[inline]
    pub const fn to_xml(self) -> &'static str {
        match self {
            Self::Portrait => "portrait",
            Self::Landscape => "landscape",
        }
    }

    /// Parse orientation from an XML attribute value.
    #[inline]
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "portrait" => Some(Self::Portrait),
            "landscape" => Some(Self::Landscape),
            _ => None,
        }
    }
}

impl Default for Orientation {
    #[inline]
    fn default() -> Self {
        Self::Portrait
    }
}

impl fmt::Display for Orientation {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Portrait => write!(formatter, "Portrait"),
            Self::Landscape => write!(formatter, "Landscape"),
        }
    }
}

/// Section break start type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Start {
    /// Continuous section break.
    Continuous = 0,
    /// New column section break.
    NewColumn = 1,
    /// New page section break.
    NewPage = 2,
    /// Even pages section break.
    EvenPage = 3,
    /// Section begins on the next odd page.
    OddPage = 4,
}

impl Start {
    /// Convert the section start type to its XML attribute value.
    #[inline]
    pub const fn to_xml(self) -> &'static str {
        match self {
            Self::Continuous => "continuous",
            Self::NewColumn => "nextColumn",
            Self::NewPage => "nextPage",
            Self::EvenPage => "evenPage",
            Self::OddPage => "oddPage",
        }
    }

    /// Parse section start type from an XML attribute value.
    #[inline]
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "continuous" => Some(Self::Continuous),
            "nextColumn" => Some(Self::NewColumn),
            "nextPage" => Some(Self::NewPage),
            "evenPage" => Some(Self::EvenPage),
            "oddPage" => Some(Self::OddPage),
            _ => None,
        }
    }
}

impl Default for Start {
    #[inline]
    fn default() -> Self {
        Self::NewPage
    }
}

impl fmt::Display for Start {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Continuous => write!(formatter, "Continuous"),
            Self::NewColumn => write!(formatter, "New Column"),
            Self::NewPage => write!(formatter, "New Page"),
            Self::EvenPage => write!(formatter, "Even Page"),
            Self::OddPage => write!(formatter, "Odd Page"),
        }
    }
}

/// Length in English Metric Units (EMUs).
///
/// 1 EMU = 1/914,400 of an inch = 1/360,000 of a centimeter.
/// This is the standard unit used in OOXML for measurements.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Emu(pub i64);

impl Emu {
    /// Create from inches.
    #[inline]
    pub const fn from_inches(inches: f64) -> Self {
        Self((inches * EMUS_PER_INCH as f64) as i64)
    }

    /// Create from centimeters.
    #[inline]
    pub const fn from_cm(cm: f64) -> Self {
        Self((cm * EMUS_PER_CM as f64) as i64)
    }

    /// Create from points (1/72 of an inch).
    #[inline]
    pub const fn from_pt(pt: f64) -> Self {
        Self((pt * EMUS_PER_PT as f64) as i64)
    }

    /// Create from twips (1/20 of a point, 1/1440 of an inch).
    #[inline]
    pub const fn from_twips(twips: i64) -> Self {
        Self(twips * EMUS_PER_TWIP)
    }

    /// Convert to inches.
    #[inline]
    pub fn to_inches(self) -> f64 {
        self.0 as f64 / EMUS_PER_INCH as f64
    }

    /// Convert to centimeters.
    #[inline]
    pub fn to_cm(self) -> f64 {
        self.0 as f64 / EMUS_PER_CM as f64
    }

    /// Convert to points.
    #[inline]
    pub fn to_pt(self) -> f64 {
        self.0 as f64 / EMUS_PER_PT as f64
    }

    /// Convert to twips.
    #[inline]
    pub fn to_twips(self) -> i64 {
        emu_to_twip_i64(self.0)
    }
}

impl From<i64> for Emu {
    #[inline]
    fn from(value: i64) -> Self {
        Self(value)
    }
}

/// Page margins for a section.
///
/// All measurements are in EMUs (English Metric Units).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Margins {
    /// Top margin
    pub top: Option<Emu>,
    /// Right margin
    pub right: Option<Emu>,
    /// Bottom margin
    pub bottom: Option<Emu>,
    /// Left margin
    pub left: Option<Emu>,
    /// Header distance from top edge
    pub header: Option<Emu>,
    /// Footer distance from bottom edge
    pub footer: Option<Emu>,
    /// Gutter margin (for binding)
    pub gutter: Option<Emu>,
}

/// Page size for a section.
///
/// Both dimensions are in EMUs (English Metric Units).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PageSize {
    /// Page width
    pub width: Option<Emu>,
    /// Page height
    pub height: Option<Emu>,
    /// Page orientation
    pub orientation: Orientation,
}

impl Default for PageSize {
    fn default() -> Self {
        Self {
            width: None,
            height: None,
            orientation: Orientation::Portrait,
        }
    }
}

/// A section in a Word document.
///
/// Represents a `<w:sectPr>` element in the document XML.
/// Each section can have different page setup properties such as
/// margins, page size, orientation, and headers/footers.
///
/// # Examples
///
/// ```rust,ignore
/// use litchi_docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// for section in doc.sections()? {
///     if let Some(width) = section.page_width() {
///         println!("Page width: {} inches", width.to_inches());
///     }
///     println!("Orientation: {}", section.orientation());
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct Section {
    /// XML bytes containing the sectPr element
    xml_bytes: Vec<u8>,
    /// Cached page size
    page_size: Option<PageSize>,
    /// Cached margins
    margins: Option<Margins>,
    /// Cached start type
    start_type: Option<Start>,
}

impl Section {
    /// Create a new Section from XML bytes containing a `<w:sectPr>` element.
    ///
    /// This is typically called internally when parsing sections from a document.
    pub fn from_xml_bytes(xml_bytes: Vec<u8>) -> Result<Self> {
        Ok(Self {
            xml_bytes,
            page_size: None,
            margins: None,
            start_type: None,
        })
    }

    /// Get the page width for this section.
    ///
    /// Returns `None` if no page width is specified in the section properties.
    pub fn page_width(&mut self) -> Option<Emu> {
        self.ensure_page_size_parsed();
        self.page_size.and_then(|ps| ps.width)
    }

    /// Get the page height for this section.
    ///
    /// Returns `None` if no page height is specified in the section properties.
    pub fn page_height(&mut self) -> Option<Emu> {
        self.ensure_page_size_parsed();
        self.page_size.and_then(|ps| ps.height)
    }

    /// Get the page orientation for this section.
    ///
    /// Defaults to `Portrait` if not specified.
    pub fn orientation(&mut self) -> Orientation {
        self.ensure_page_size_parsed();
        self.page_size.map(|ps| ps.orientation).unwrap_or_default()
    }

    /// Get the top margin for this section.
    pub fn top_margin(&mut self) -> Option<Emu> {
        self.ensure_margins_parsed();
        self.margins.and_then(|m| m.top)
    }

    /// Get the right margin for this section.
    pub fn right_margin(&mut self) -> Option<Emu> {
        self.ensure_margins_parsed();
        self.margins.and_then(|m| m.right)
    }

    /// Get the bottom margin for this section.
    pub fn bottom_margin(&mut self) -> Option<Emu> {
        self.ensure_margins_parsed();
        self.margins.and_then(|m| m.bottom)
    }

    /// Get the left margin for this section.
    pub fn left_margin(&mut self) -> Option<Emu> {
        self.ensure_margins_parsed();
        self.margins.and_then(|m| m.left)
    }

    /// Get the header distance for this section.
    ///
    /// This is the distance from the top edge of the page to the top edge of the header.
    pub fn header_distance(&mut self) -> Option<Emu> {
        self.ensure_margins_parsed();
        self.margins.and_then(|m| m.header)
    }

    /// Get the footer distance for this section.
    ///
    /// This is the distance from the bottom edge of the page to the bottom edge of the footer.
    pub fn footer_distance(&mut self) -> Option<Emu> {
        self.ensure_margins_parsed();
        self.margins.and_then(|m| m.footer)
    }

    /// Get the gutter margin for this section.
    ///
    /// The gutter is extra spacing added to the inner margin for binding.
    pub fn gutter(&mut self) -> Option<Emu> {
        self.ensure_margins_parsed();
        self.margins.and_then(|m| m.gutter)
    }

    /// Get the section start type.
    ///
    /// Determines how the section break is inserted (continuous, new page, etc.).
    pub fn start_type(&mut self) -> Start {
        self.ensure_start_type_parsed();
        self.start_type.unwrap_or_default()
    }

    /// Parse page size from the XML if not already cached.
    fn ensure_page_size_parsed(&mut self) {
        if self.page_size.is_some() {
            return;
        }

        let mut page_size = PageSize::default();
        let mut reader = Reader::from_reader(self.xml_bytes.as_slice());
        reader.config_mut().trim_text(true);

        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) if e.local_name().as_ref() == b"pgSz" => {
                    // Parse page size attributes
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"w" => {
                                if let Ok(value) = attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) && let Ok(twips) = value.parse::<i64>()
                                {
                                    page_size.width = Some(Emu::from_twips(twips));
                                }
                            },
                            b"h" => {
                                if let Ok(value) = attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) && let Ok(twips) = value.parse::<i64>()
                                {
                                    page_size.height = Some(Emu::from_twips(twips));
                                }
                            },
                            b"orient" => {
                                if let Ok(value) = attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) {
                                    page_size.orientation = Orientation::from_xml(&value)
                                        .unwrap_or(Orientation::Portrait);
                                }
                            },
                            _ => {},
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
        }

        self.page_size = Some(page_size);
    }

    /// Parse margins from the XML if not already cached.
    fn ensure_margins_parsed(&mut self) {
        if self.margins.is_some() {
            return;
        }

        let mut margins = Margins::default();
        let mut reader = Reader::from_reader(self.xml_bytes.as_slice());
        reader.config_mut().trim_text(true);

        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) if e.local_name().as_ref() == b"pgMar" => {
                    // Parse margin attributes
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"top" => {
                                if let Ok(value) = attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) && let Ok(twips) = value.parse::<i64>()
                                {
                                    margins.top = Some(Emu::from_twips(twips));
                                }
                            },
                            b"right" => {
                                if let Ok(value) = attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) && let Ok(twips) = value.parse::<i64>()
                                {
                                    margins.right = Some(Emu::from_twips(twips));
                                }
                            },
                            b"bottom" => {
                                if let Ok(value) = attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) && let Ok(twips) = value.parse::<i64>()
                                {
                                    margins.bottom = Some(Emu::from_twips(twips));
                                }
                            },
                            b"left" => {
                                if let Ok(value) = attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) && let Ok(twips) = value.parse::<i64>()
                                {
                                    margins.left = Some(Emu::from_twips(twips));
                                }
                            },
                            b"header" => {
                                if let Ok(value) = attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) && let Ok(twips) = value.parse::<i64>()
                                {
                                    margins.header = Some(Emu::from_twips(twips));
                                }
                            },
                            b"footer" => {
                                if let Ok(value) = attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) && let Ok(twips) = value.parse::<i64>()
                                {
                                    margins.footer = Some(Emu::from_twips(twips));
                                }
                            },
                            b"gutter" => {
                                if let Ok(value) = attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) && let Ok(twips) = value.parse::<i64>()
                                {
                                    margins.gutter = Some(Emu::from_twips(twips));
                                }
                            },
                            _ => {},
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
        }

        self.margins = Some(margins);
    }

    /// Parse section start type from the XML if not already cached.
    fn ensure_start_type_parsed(&mut self) {
        if self.start_type.is_some() {
            return;
        }

        let mut start_type = Start::default();
        let mut reader = Reader::from_reader(self.xml_bytes.as_slice());
        reader.config_mut().trim_text(true);

        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) if e.local_name().as_ref() == b"type" => {
                    // Parse type attribute
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"val"
                            && let Ok(value) = attr.decoded_and_normalized_value(
                                XmlVersion::Implicit1_0,
                                reader.decoder(),
                            )
                        {
                            start_type = Start::from_xml(&value).unwrap_or(Start::NewPage);
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
        }

        self.start_type = Some(start_type);
    }
}

/// A collection of sections in a Word document.
///
/// Provides access to all sections in document order, supporting
/// iteration and indexed access.
///
/// # Examples
///
/// ```rust,ignore
/// use litchi_docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
/// let sections = doc.sections()?;
///
/// println!("Document has {} sections", sections.len());
/// for (i, mut section) in sections.iter_mut().enumerate() {
///     println!("Section {}: orientation = {}", i, section.orientation());
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct Sections {
    /// Individual section objects
    sections: Vec<Section>,
}

impl Sections {
    /// Create a new Sections collection from a vector of Section objects.
    #[inline]
    pub fn new(sections: Vec<Section>) -> Self {
        Self { sections }
    }

    /// Get the number of sections in the document.
    #[inline]
    pub fn len(&self) -> usize {
        self.sections.len()
    }

    /// Check if the document has no sections.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.sections.is_empty()
    }

    /// Get a mutable reference to a section by index.
    ///
    /// Returns `None` if the index is out of bounds.
    #[inline]
    pub fn get_mut(&mut self, index: usize) -> Option<&mut Section> {
        self.sections.get_mut(index)
    }

    /// Get a reference to a section by index.
    ///
    /// Returns `None` if the index is out of bounds.
    #[inline]
    pub fn get(&self, index: usize) -> Option<&Section> {
        self.sections.get(index)
    }

    /// Get a mutable iterator over the sections.
    #[inline]
    pub fn iter_mut(&mut self) -> std::slice::IterMut<'_, Section> {
        self.sections.iter_mut()
    }

    /// Get an iterator over the sections.
    #[inline]
    pub fn iter(&self) -> std::slice::Iter<'_, Section> {
        self.sections.iter()
    }
}

impl std::ops::Index<usize> for Sections {
    type Output = Section;

    #[inline]
    fn index(&self, index: usize) -> &Self::Output {
        &self.sections[index]
    }
}

impl std::ops::IndexMut<usize> for Sections {
    #[inline]
    fn index_mut(&mut self, index: usize) -> &mut Self::Output {
        &mut self.sections[index]
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_emu_conversions() {
        let emu = Emu::from_inches(1.0);
        assert_eq!(emu.0, 914_400);
        assert!((emu.to_inches() - 1.0).abs() < 0.0001);

        let emu = Emu::from_twips(1440);
        assert_eq!(emu.0, 914_400);
        assert_eq!(emu.to_twips(), 1440);
    }

    #[test]
    fn test_orientation_default() {
        let orientation = Orientation::default();
        assert_eq!(orientation, Orientation::Portrait);
    }

    #[test]
    fn orientation_xml_and_display_round_trip() {
        assert_eq!(Orientation::Portrait.to_xml(), "portrait");
        assert_eq!(Orientation::Landscape.to_xml(), "landscape");
        assert_eq!(
            Orientation::from_xml("portrait"),
            Some(Orientation::Portrait)
        );
        assert_eq!(
            Orientation::from_xml("landscape"),
            Some(Orientation::Landscape)
        );
        assert_eq!(Orientation::from_xml("invalid"), None);
        assert_eq!(Orientation::Portrait.to_string(), "Portrait");
        assert_eq!(Orientation::Landscape.to_string(), "Landscape");
        assert_eq!(Orientation::Landscape as u8, 1);
    }

    #[test]
    fn section_start_xml_and_display_round_trip() {
        let values = [
            (Start::Continuous, "continuous", "Continuous", 0),
            (Start::NewColumn, "nextColumn", "New Column", 1),
            (Start::NewPage, "nextPage", "New Page", 2),
            (Start::EvenPage, "evenPage", "Even Page", 3),
            (Start::OddPage, "oddPage", "Odd Page", 4),
        ];
        for (value, xml, display, repr) in values {
            assert_eq!(value.to_xml(), xml);
            assert_eq!(Start::from_xml(xml), Some(value));
            assert_eq!(value.to_string(), display);
            assert_eq!(value as u8, repr);
        }
        assert_eq!(Start::from_xml("invalid"), None);
        assert_eq!(Start::default(), Start::NewPage);
    }

    #[test]
    fn test_sections_collection() {
        let sections = Sections::new(vec![]);
        assert_eq!(sections.len(), 0);
        assert!(sections.is_empty());
    }
}
