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
    }

    #[test]
    fn test_section_start_conversion() {
        assert_eq!(WdSectionStart::NewPage.to_xml(), "nextPage");
        assert_eq!(
            WdSectionStart::from_xml("nextPage"),
            Some(WdSectionStart::NewPage)
        );
    }

    #[test]
    fn test_header_footer_conversion() {
        assert_eq!(WdHeaderFooter::Primary.to_xml(), "default");
        assert_eq!(
            WdHeaderFooter::from_xml("default"),
            Some(WdHeaderFooter::Primary)
        );
    }

    #[test]
    fn test_style_type_conversion() {
        assert_eq!(WdStyleType::Paragraph.to_xml(), "paragraph");
        assert_eq!(
            WdStyleType::from_xml("paragraph"),
            Some(WdStyleType::Paragraph)
        );
    }
}
