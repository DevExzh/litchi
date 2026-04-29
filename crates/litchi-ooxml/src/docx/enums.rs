/// Enumerations for Word document elements.
///
/// This module provides various enumerations used throughout the Word document API,
/// matching those found in the VBA API and python-docx.
use std::fmt;

/// Specifies the page layout orientation.
///
/// Corresponds to the VBA `WdOrientation` enumeration.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::enums::WdOrientation;
///
/// let orientation = WdOrientation::Landscape;
/// assert_eq!(orientation.to_xml(), "landscape");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum WdOrientation {
    /// Portrait orientation.
    Portrait = 0,
    /// Landscape orientation.
    Landscape = 1,
}

impl WdOrientation {
    /// Convert the orientation to its XML attribute value.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::enums::WdOrientation;
    ///
    /// assert_eq!(WdOrientation::Portrait.to_xml(), "portrait");
    /// assert_eq!(WdOrientation::Landscape.to_xml(), "landscape");
    /// ```
    #[inline]
    pub const fn to_xml(self) -> &'static str {
        match self {
            Self::Portrait => "portrait",
            Self::Landscape => "landscape",
        }
    }

    /// Parse orientation from XML attribute value.
    ///
    /// Returns `None` if the value is not recognized.
    #[inline]
    pub fn from_xml(s: &str) -> Option<Self> {
        match s {
            "portrait" => Some(Self::Portrait),
            "landscape" => Some(Self::Landscape),
            _ => None,
        }
    }
}

impl Default for WdOrientation {
    #[inline]
    fn default() -> Self {
        Self::Portrait
    }
}

impl fmt::Display for WdOrientation {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Portrait => write!(f, "Portrait"),
            Self::Landscape => write!(f, "Landscape"),
        }
    }
}

/// Specifies the start type of a section break.
///
/// Corresponds to the VBA `WdSectionStart` enumeration.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::enums::WdSectionStart;
///
/// let start_type = WdSectionStart::NewPage;
/// assert_eq!(start_type.to_xml(), "nextPage");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum WdSectionStart {
    /// Continuous section break.
    Continuous = 0,
    /// New column section break.
    NewColumn = 1,
    /// New page section break.
    NewPage = 2,
    /// Even pages section break.
    EvenPage = 3,
    /// Section begins on next odd page.
    OddPage = 4,
}

impl WdSectionStart {
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

    /// Parse section start type from XML attribute value.
    ///
    /// Returns `None` if the value is not recognized.
    #[inline]
    pub fn from_xml(s: &str) -> Option<Self> {
        match s {
            "continuous" => Some(Self::Continuous),
            "nextColumn" => Some(Self::NewColumn),
            "nextPage" => Some(Self::NewPage),
            "evenPage" => Some(Self::EvenPage),
            "oddPage" => Some(Self::OddPage),
            _ => None,
        }
    }
}

impl Default for WdSectionStart {
    #[inline]
    fn default() -> Self {
        Self::NewPage
    }
}

impl fmt::Display for WdSectionStart {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Continuous => write!(f, "Continuous"),
            Self::NewColumn => write!(f, "New Column"),
            Self::NewPage => write!(f, "New Page"),
            Self::EvenPage => write!(f, "Even Page"),
            Self::OddPage => write!(f, "Odd Page"),
        }
    }
}

/// Specifies one of the three possible header/footer definitions for a section.
///
/// Corresponds to the VBA `WdHeaderFooterIndex` enumeration.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::enums::WdHeaderFooter;
///
/// let index = WdHeaderFooter::Primary;
/// assert_eq!(index.to_xml(), "default");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum WdHeaderFooter {
    /// Header/footer for odd pages or all pages if no even header/footer.
    Primary = 1,
    /// Header/footer for first page of section.
    FirstPage = 2,
    /// Header/footer for even pages of recto/verso section.
    EvenPage = 3,
}

impl WdHeaderFooter {
    /// Convert the header/footer index to its XML attribute value.
    #[inline]
    pub const fn to_xml(self) -> &'static str {
        match self {
            Self::Primary => "default",
            Self::FirstPage => "first",
            Self::EvenPage => "even",
        }
    }

    /// Parse header/footer index from XML attribute value.
    ///
    /// Returns `None` if the value is not recognized.
    #[inline]
    pub fn from_xml(s: &str) -> Option<Self> {
        match s {
            "default" => Some(Self::Primary),
            "first" => Some(Self::FirstPage),
            "even" => Some(Self::EvenPage),
            _ => None,
        }
    }
}

impl Default for WdHeaderFooter {
    #[inline]
    fn default() -> Self {
        Self::Primary
    }
}

impl fmt::Display for WdHeaderFooter {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Primary => write!(f, "Primary"),
            Self::FirstPage => write!(f, "First Page"),
            Self::EvenPage => write!(f, "Even Page"),
        }
    }
}

/// Specifies one of the four style types: paragraph, character, list, or table.
///
/// Corresponds to the VBA `WdStyleType` enumeration.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::enums::WdStyleType;
///
/// let style_type = WdStyleType::Paragraph;
/// assert_eq!(style_type.to_xml(), "paragraph");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum WdStyleType {
    /// Paragraph style.
    Paragraph = 1,
    /// Character style.
    Character = 2,
    /// Table style.
    Table = 3,
    /// List (numbering) style.
    List = 4,
}

impl WdStyleType {
    /// Convert the style type to its XML attribute value.
    #[inline]
    pub const fn to_xml(self) -> &'static str {
        match self {
            Self::Paragraph => "paragraph",
            Self::Character => "character",
            Self::Table => "table",
            Self::List => "numbering",
        }
    }

    /// Parse style type from XML attribute value.
    ///
    /// Returns `None` if the value is not recognized.
    #[inline]
    pub fn from_xml(s: &str) -> Option<Self> {
        match s {
            "paragraph" => Some(Self::Paragraph),
            "character" => Some(Self::Character),
            "table" => Some(Self::Table),
            "numbering" => Some(Self::List),
            _ => None,
        }
    }
}

impl Default for WdStyleType {
    #[inline]
    fn default() -> Self {
        // Per spec, if no type is specified, it defaults to paragraph
        Self::Paragraph
    }
}

impl fmt::Display for WdStyleType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Paragraph => write!(f, "Paragraph"),
            Self::Character => write!(f, "Character"),
            Self::Table => write!(f, "Table"),
            Self::List => write!(f, "List"),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_orientation_conversion() {
        assert_eq!(WdOrientation::Portrait.to_xml(), "portrait");
        assert_eq!(WdOrientation::Landscape.to_xml(), "landscape");
        assert_eq!(
            WdOrientation::from_xml("portrait"),
            Some(WdOrientation::Portrait)
        );
        assert_eq!(
            WdOrientation::from_xml("landscape"),
            Some(WdOrientation::Landscape)
        );
        assert_eq!(WdOrientation::from_xml("invalid"), None);
        assert_eq!(WdOrientation::from_xml(""), None);
    }

    #[test]
    fn test_orientation_default() {
        assert_eq!(WdOrientation::default(), WdOrientation::Portrait);
    }

    #[test]
    fn test_orientation_display() {
        assert_eq!(format!("{}", WdOrientation::Portrait), "Portrait");
        assert_eq!(format!("{}", WdOrientation::Landscape), "Landscape");
    }

    #[test]
    fn test_orientation_variants() {
        // Test that all variants can be created
        let portrait = WdOrientation::Portrait;
        let landscape = WdOrientation::Landscape;

        // Test equality and inequality
        assert_eq!(portrait, WdOrientation::Portrait);
        assert_eq!(landscape, WdOrientation::Landscape);
        assert_ne!(portrait, landscape);
    }

    #[test]
    fn test_orientation_debug() {
        let debug_str = format!("{:?}", WdOrientation::Portrait);
        assert!(debug_str.contains("Portrait"));
    }

    #[test]
    fn test_section_start_conversion() {
        assert_eq!(WdSectionStart::Continuous.to_xml(), "continuous");
        assert_eq!(WdSectionStart::NewColumn.to_xml(), "nextColumn");
        assert_eq!(WdSectionStart::NewPage.to_xml(), "nextPage");
        assert_eq!(WdSectionStart::EvenPage.to_xml(), "evenPage");
        assert_eq!(WdSectionStart::OddPage.to_xml(), "oddPage");

        assert_eq!(
            WdSectionStart::from_xml("continuous"),
            Some(WdSectionStart::Continuous)
        );
        assert_eq!(
            WdSectionStart::from_xml("nextColumn"),
            Some(WdSectionStart::NewColumn)
        );
        assert_eq!(
            WdSectionStart::from_xml("nextPage"),
            Some(WdSectionStart::NewPage)
        );
        assert_eq!(
            WdSectionStart::from_xml("evenPage"),
            Some(WdSectionStart::EvenPage)
        );
        assert_eq!(
            WdSectionStart::from_xml("oddPage"),
            Some(WdSectionStart::OddPage)
        );
        assert_eq!(WdSectionStart::from_xml("invalid"), None);
    }

    #[test]
    fn test_section_start_default() {
        assert_eq!(WdSectionStart::default(), WdSectionStart::NewPage);
    }

    #[test]
    fn test_section_start_display() {
        assert_eq!(format!("{}", WdSectionStart::Continuous), "Continuous");
        assert_eq!(format!("{}", WdSectionStart::NewColumn), "New Column");
        assert_eq!(format!("{}", WdSectionStart::NewPage), "New Page");
        assert_eq!(format!("{}", WdSectionStart::EvenPage), "Even Page");
        assert_eq!(format!("{}", WdSectionStart::OddPage), "Odd Page");
    }

    #[test]
    fn test_section_start_variants() {
        assert_eq!(WdSectionStart::Continuous as u8, 0);
        assert_eq!(WdSectionStart::NewColumn as u8, 1);
        assert_eq!(WdSectionStart::NewPage as u8, 2);
        assert_eq!(WdSectionStart::EvenPage as u8, 3);
        assert_eq!(WdSectionStart::OddPage as u8, 4);
    }

    #[test]
    fn test_header_footer_conversion() {
        assert_eq!(WdHeaderFooter::Primary.to_xml(), "default");
        assert_eq!(WdHeaderFooter::FirstPage.to_xml(), "first");
        assert_eq!(WdHeaderFooter::EvenPage.to_xml(), "even");

        assert_eq!(
            WdHeaderFooter::from_xml("default"),
            Some(WdHeaderFooter::Primary)
        );
        assert_eq!(
            WdHeaderFooter::from_xml("first"),
            Some(WdHeaderFooter::FirstPage)
        );
        assert_eq!(
            WdHeaderFooter::from_xml("even"),
            Some(WdHeaderFooter::EvenPage)
        );
        assert_eq!(WdHeaderFooter::from_xml("invalid"), None);
    }

    #[test]
    fn test_header_footer_default() {
        assert_eq!(WdHeaderFooter::default(), WdHeaderFooter::Primary);
    }

    #[test]
    fn test_header_footer_display() {
        assert_eq!(format!("{}", WdHeaderFooter::Primary), "Primary");
        assert_eq!(format!("{}", WdHeaderFooter::FirstPage), "First Page");
        assert_eq!(format!("{}", WdHeaderFooter::EvenPage), "Even Page");
    }

    #[test]
    fn test_header_footer_variants() {
        assert_eq!(WdHeaderFooter::Primary as u8, 1);
        assert_eq!(WdHeaderFooter::FirstPage as u8, 2);
        assert_eq!(WdHeaderFooter::EvenPage as u8, 3);
    }

    #[test]
    fn test_style_type_conversion() {
        assert_eq!(WdStyleType::Paragraph.to_xml(), "paragraph");
        assert_eq!(WdStyleType::Character.to_xml(), "character");
        assert_eq!(WdStyleType::Table.to_xml(), "table");
        assert_eq!(WdStyleType::List.to_xml(), "numbering");

        assert_eq!(
            WdStyleType::from_xml("paragraph"),
            Some(WdStyleType::Paragraph)
        );
        assert_eq!(
            WdStyleType::from_xml("character"),
            Some(WdStyleType::Character)
        );
        assert_eq!(WdStyleType::from_xml("table"), Some(WdStyleType::Table));
        assert_eq!(WdStyleType::from_xml("numbering"), Some(WdStyleType::List));
        assert_eq!(WdStyleType::from_xml("invalid"), None);
    }

    #[test]
    fn test_style_type_default() {
        assert_eq!(WdStyleType::default(), WdStyleType::Paragraph);
    }

    #[test]
    fn test_style_type_display() {
        assert_eq!(format!("{}", WdStyleType::Paragraph), "Paragraph");
        assert_eq!(format!("{}", WdStyleType::Character), "Character");
        assert_eq!(format!("{}", WdStyleType::Table), "Table");
        assert_eq!(format!("{}", WdStyleType::List), "List");
    }

    #[test]
    fn test_style_type_variants() {
        assert_eq!(WdStyleType::Paragraph as u8, 1);
        assert_eq!(WdStyleType::Character as u8, 2);
        assert_eq!(WdStyleType::Table as u8, 3);
        assert_eq!(WdStyleType::List as u8, 4);
    }

    #[test]
    fn test_all_enums_clone_copy() {
        // Test that all enums implement Clone and Copy
        let orientation = WdOrientation::Portrait;
        let orientation2 = orientation;
        assert_eq!(orientation, orientation2);

        let section_start = WdSectionStart::NewPage;
        let section_start2 = section_start;
        assert_eq!(section_start, section_start2);

        let header_footer = WdHeaderFooter::Primary;
        let header_footer2 = header_footer;
        assert_eq!(header_footer, header_footer2);

        let style_type = WdStyleType::Paragraph;
        let style_type2 = style_type;
        assert_eq!(style_type, style_type2);
    }

    #[test]
    fn test_all_enums_hash() {
        use std::collections::HashSet;

        // Test that all enums implement Hash
        let mut orientations = HashSet::new();
        orientations.insert(WdOrientation::Portrait);
        orientations.insert(WdOrientation::Landscape);
        assert_eq!(orientations.len(), 2);

        let mut section_starts = HashSet::new();
        section_starts.insert(WdSectionStart::Continuous);
        section_starts.insert(WdSectionStart::NewPage);
        assert_eq!(section_starts.len(), 2);
    }
}
