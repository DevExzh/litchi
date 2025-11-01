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
#[derive(Debug, Clone, Copy)]
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
