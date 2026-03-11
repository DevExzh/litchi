/// Page number format for headers/footers.
#[derive(Debug, Clone, Copy)]
pub enum PageNumberFormat {
    /// Decimal numbers (1, 2, 3, ...)
    Decimal,
    /// Uppercase Roman numerals (I, II, III, ...)
    UpperRoman,
    /// Lowercase Roman numerals (i, ii, iii, ...)
    LowerRoman,
    /// Uppercase letters (A, B, C, ...)
    UpperLetter,
    /// Lowercase letters (a, b, c, ...)
    LowerLetter,
}

impl PageNumberFormat {
    pub(crate) fn as_str(&self) -> &'static str {
        match self {
            Self::Decimal => "decimal",
            Self::UpperRoman => "upperRoman",
            Self::LowerRoman => "lowerRoman",
            Self::UpperLetter => "upperLetter",
            Self::LowerLetter => "lowerLetter",
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
    /// Convert orientation to XML string representation.
    #[allow(dead_code)]
    pub(crate) fn as_str(&self) -> &'static str {
        match self {
            Self::Portrait => "portrait",
            Self::Landscape => "landscape",
        }
    }
}

/// Section properties including page setup and margins.
#[derive(Debug, Clone)]
pub struct SectionProperties {
    /// Page width in twips (twentieth of a point, 1440 = 1 inch)
    pub page_width: u32,
    /// Page height in twips
    pub page_height: u32,
    /// Page orientation
    pub orientation: PageOrientation,
    /// Top margin in twips
    pub margin_top: u32,
    /// Bottom margin in twips
    pub margin_bottom: u32,
    /// Left margin in twips
    pub margin_left: u32,
    /// Right margin in twips
    pub margin_right: u32,
    /// Header distance from top in twips
    pub header_distance: u32,
    /// Footer distance from bottom in twips
    pub footer_distance: u32,
}

impl Default for SectionProperties {
    fn default() -> Self {
        // US Letter size: 8.5" x 11" = 12240 x 15840 twips
        Self {
            page_width: 12240,
            page_height: 15840,
            orientation: PageOrientation::Portrait,
            margin_top: 1440,     // 1 inch
            margin_bottom: 1440,  // 1 inch
            margin_left: 1440,    // 1 inch
            margin_right: 1440,   // 1 inch
            header_distance: 720, // 0.5 inch
            footer_distance: 720, // 0.5 inch
        }
    }
}

impl SectionProperties {
    /// Create A4 page size (210mm x 297mm).
    pub fn a4() -> Self {
        Self {
            page_width: 11906,  // 210mm = 8.27 inches
            page_height: 16838, // 297mm = 11.69 inches
            ..Default::default()
        }
    }

    /// Create US Letter page size (8.5" x 11").
    pub fn letter() -> Self {
        Self::default()
    }

    /// Create Legal page size (8.5" x 14").
    pub fn legal() -> Self {
        Self {
            page_width: 12240,
            page_height: 20160, // 14 inches
            ..Default::default()
        }
    }

    /// Set page to landscape orientation.
    pub fn landscape(mut self) -> Self {
        self.orientation = PageOrientation::Landscape;
        // Swap width and height for landscape
        std::mem::swap(&mut self.page_width, &mut self.page_height);
        self
    }

    /// Set margins (all in inches).
    pub fn margins(mut self, top: f64, bottom: f64, left: f64, right: f64) -> Self {
        self.margin_top = (top * 1440.0) as u32;
        self.margin_bottom = (bottom * 1440.0) as u32;
        self.margin_left = (left * 1440.0) as u32;
        self.margin_right = (right * 1440.0) as u32;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_page_number_format_as_str() {
        assert_eq!(PageNumberFormat::Decimal.as_str(), "decimal");
        assert_eq!(PageNumberFormat::UpperRoman.as_str(), "upperRoman");
        assert_eq!(PageNumberFormat::LowerRoman.as_str(), "lowerRoman");
        assert_eq!(PageNumberFormat::UpperLetter.as_str(), "upperLetter");
        assert_eq!(PageNumberFormat::LowerLetter.as_str(), "lowerLetter");
    }

    #[test]
    fn test_page_orientation_as_str() {
        assert_eq!(PageOrientation::Portrait.as_str(), "portrait");
        assert_eq!(PageOrientation::Landscape.as_str(), "landscape");
    }

    #[test]
    fn test_section_properties_default() {
        let props = SectionProperties::default();
        assert_eq!(props.page_width, 12240); // 8.5 inches
        assert_eq!(props.page_height, 15840); // 11 inches
        assert_eq!(props.margin_top, 1440); // 1 inch
        assert_eq!(props.margin_bottom, 1440);
        assert_eq!(props.margin_left, 1440);
        assert_eq!(props.margin_right, 1440);
        assert_eq!(props.header_distance, 720); // 0.5 inch
        assert_eq!(props.footer_distance, 720);
    }

    #[test]
    fn test_section_properties_a4() {
        let props = SectionProperties::a4();
        assert_eq!(props.page_width, 11906); // ~210mm
        assert_eq!(props.page_height, 16838); // ~297mm
        // Margins should still be default
        assert_eq!(props.margin_top, 1440);
    }

    #[test]
    fn test_section_properties_letter() {
        let props = SectionProperties::letter();
        assert_eq!(props.page_width, 12240);
        assert_eq!(props.page_height, 15840);
    }

    #[test]
    fn test_section_properties_legal() {
        let props = SectionProperties::legal();
        assert_eq!(props.page_width, 12240); // 8.5 inches
        assert_eq!(props.page_height, 20160); // 14 inches
    }

    #[test]
    fn test_section_properties_landscape() {
        let props = SectionProperties::default().landscape();
        assert_eq!(props.orientation, PageOrientation::Landscape);
        // Width and height should be swapped
        assert_eq!(props.page_width, 15840); // Was height
        assert_eq!(props.page_height, 12240); // Was width
    }

    #[test]
    fn test_section_properties_a4_landscape() {
        let props = SectionProperties::a4().landscape();
        assert_eq!(props.orientation, PageOrientation::Landscape);
        // A4 dimensions swapped
        assert_eq!(props.page_width, 16838);
        assert_eq!(props.page_height, 11906);
    }

    #[test]
    fn test_section_properties_margins() {
        let props = SectionProperties::default().margins(1.5, 1.5, 1.0, 1.0);
        assert_eq!(props.margin_top, 2160); // 1.5 * 1440
        assert_eq!(props.margin_bottom, 2160);
        assert_eq!(props.margin_left, 1440); // 1.0 * 1440
        assert_eq!(props.margin_right, 1440);
    }

    #[test]
    fn test_section_properties_chained() {
        let props = SectionProperties::a4()
            .landscape()
            .margins(2.0, 2.0, 1.5, 1.5);
        assert_eq!(props.orientation, PageOrientation::Landscape);
        assert_eq!(props.margin_top, 2880); // 2.0 * 1440
        assert_eq!(props.margin_left, 2160); // 1.5 * 1440
    }

    #[test]
    fn test_page_number_format_debug() {
        let format = PageNumberFormat::UpperRoman;
        let debug_str = format!("{:?}", format);
        assert!(debug_str.contains("UpperRoman"));
    }

    #[test]
    fn test_page_orientation_debug() {
        let orientation = PageOrientation::Landscape;
        let debug_str = format!("{:?}", orientation);
        assert!(debug_str.contains("Landscape"));
    }

    #[test]
    fn test_section_properties_debug() {
        let props = SectionProperties::default();
        let debug_str = format!("{:?}", props);
        assert!(debug_str.contains("12240"));
        assert!(debug_str.contains("Portrait"));
    }
}
