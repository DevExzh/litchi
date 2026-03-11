/// Header and footer access for Word documents
use super::package::Result;
use super::paragraph::Paragraph;
use super::parts::headers::HeaderFooterType;

/// A header or footer in a Word document
#[derive(Debug, Clone)]
pub struct HeaderFooter {
    /// Type of header/footer
    pub header_footer_type: HeaderFooterType,
    /// Text content
    pub text: String,
    /// Paragraphs in this header/footer
    pub paragraphs: Vec<Paragraph>,
}

impl HeaderFooter {
    /// Create a new header or footer
    pub fn new(header_footer_type: HeaderFooterType, text: String) -> Self {
        Self {
            header_footer_type,
            text,
            paragraphs: Vec::new(),
        }
    }

    /// Get the text content
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Get the paragraphs
    pub fn paragraphs(&self) -> Result<&[Paragraph]> {
        Ok(&self.paragraphs)
    }

    /// Check if this is a header
    pub fn is_header(&self) -> bool {
        self.header_footer_type.is_header()
    }

    /// Check if this is a footer
    pub fn is_footer(&self) -> bool {
        self.header_footer_type.is_footer()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ole::doc::paragraph::Paragraph;

    #[test]
    fn test_header_footer_new() {
        let header = HeaderFooter::new(HeaderFooterType::OddPageHeader, "Test Header".to_string());

        assert_eq!(header.text, "Test Header");
        assert_eq!(header.header_footer_type, HeaderFooterType::OddPageHeader);
        assert!(header.is_header());
        assert!(!header.is_footer());
        assert!(header.paragraphs.is_empty());
    }

    #[test]
    fn test_header_footer_with_paragraphs() {
        let mut header = HeaderFooter::new(
            HeaderFooterType::OddPageHeader,
            "Header with paragraphs".to_string(),
        );

        let para = Paragraph::new("Para text".to_string());
        header.paragraphs.push(para);

        assert_eq!(header.paragraphs.len(), 1);
    }

    #[test]
    fn test_header_footer_text_method() {
        let header = HeaderFooter::new(
            HeaderFooterType::OddPageHeader,
            "Header Content".to_string(),
        );

        assert_eq!(header.text(), "Header Content");
    }

    #[test]
    fn test_header_footer_paragraphs_method() {
        let mut header = HeaderFooter::new(HeaderFooterType::OddPageHeader, "Test".to_string());

        let para = Paragraph::new("Test para".to_string());
        header.paragraphs.push(para);

        let paragraphs = header.paragraphs().unwrap();
        assert_eq!(paragraphs.len(), 1);
    }

    #[test]
    fn test_first_page_header() {
        let header = HeaderFooter::new(
            HeaderFooterType::FirstPageHeader,
            "First Page Header".to_string(),
        );

        assert!(header.is_header());
        assert!(!header.is_footer());
        assert_eq!(header.header_footer_type, HeaderFooterType::FirstPageHeader);
    }

    #[test]
    fn test_first_page_footer() {
        let footer = HeaderFooter::new(
            HeaderFooterType::FirstPageFooter,
            "First Page Footer".to_string(),
        );

        assert!(footer.is_footer());
        assert!(!footer.is_header());
        assert_eq!(footer.header_footer_type, HeaderFooterType::FirstPageFooter);
    }

    #[test]
    fn test_even_page_header() {
        let header = HeaderFooter::new(
            HeaderFooterType::EvenPageHeader,
            "Even Page Header".to_string(),
        );

        assert!(header.is_header());
        assert_eq!(header.header_footer_type, HeaderFooterType::EvenPageHeader);
    }

    #[test]
    fn test_even_page_footer() {
        let footer = HeaderFooter::new(
            HeaderFooterType::EvenPageFooter,
            "Even Page Footer".to_string(),
        );

        assert!(footer.is_footer());
        assert_eq!(footer.header_footer_type, HeaderFooterType::EvenPageFooter);
    }

    #[test]
    fn test_odd_page_header() {
        let header = HeaderFooter::new(
            HeaderFooterType::OddPageHeader,
            "Odd Page Header".to_string(),
        );

        assert!(header.is_header());
        assert_eq!(header.header_footer_type, HeaderFooterType::OddPageHeader);
    }

    #[test]
    fn test_odd_page_footer() {
        let footer = HeaderFooter::new(
            HeaderFooterType::OddPageFooter,
            "Odd Page Footer".to_string(),
        );

        assert!(footer.is_footer());
        assert_eq!(footer.header_footer_type, HeaderFooterType::OddPageFooter);
    }

    #[test]
    fn test_header_footer_clone() {
        let header = HeaderFooter::new(
            HeaderFooterType::OddPageHeader,
            "Clonable Header".to_string(),
        );

        let cloned = header.clone();
        assert_eq!(cloned.text, header.text);
        assert_eq!(cloned.header_footer_type, header.header_footer_type);
    }

    #[test]
    fn test_all_header_types_are_headers() {
        assert!(HeaderFooterType::FirstPageHeader.is_header());
        assert!(HeaderFooterType::EvenPageHeader.is_header());
        assert!(HeaderFooterType::OddPageHeader.is_header());

        assert!(!HeaderFooterType::FirstPageHeader.is_footer());
        assert!(!HeaderFooterType::EvenPageHeader.is_footer());
        assert!(!HeaderFooterType::OddPageHeader.is_footer());
    }

    #[test]
    fn test_all_footer_types_are_footers() {
        assert!(HeaderFooterType::FirstPageFooter.is_footer());
        assert!(HeaderFooterType::EvenPageFooter.is_footer());
        assert!(HeaderFooterType::OddPageFooter.is_footer());

        assert!(!HeaderFooterType::FirstPageFooter.is_header());
        assert!(!HeaderFooterType::EvenPageFooter.is_header());
        assert!(!HeaderFooterType::OddPageFooter.is_header());
    }

    #[test]
    fn test_header_footer_debug() {
        let header = HeaderFooter::new(HeaderFooterType::OddPageHeader, "Debug Test".to_string());

        let debug_str = format!("{:?}", header);
        assert!(debug_str.contains("HeaderFooter"));
        assert!(debug_str.contains("Debug Test"));
    }

    #[test]
    fn test_empty_header() {
        let header = HeaderFooter::new(HeaderFooterType::OddPageHeader, "".to_string());

        assert_eq!(header.text(), "");
        assert!(header.paragraphs.is_empty());
    }

    #[test]
    fn test_header_with_unicode() {
        let header = HeaderFooter::new(
            HeaderFooterType::OddPageHeader,
            "Unicode: 你好世界 🎉".to_string(),
        );

        assert_eq!(header.text(), "Unicode: 你好世界 🎉");
    }

    #[test]
    fn test_header_footer_with_multiple_paragraphs() {
        let mut header = HeaderFooter::new(HeaderFooterType::OddPageHeader, "Header".to_string());

        for i in 0..3 {
            let para_text = format!("Para {}", i);
            header.paragraphs.push(Paragraph::new(para_text));
        }

        assert_eq!(header.paragraphs.len(), 3);
    }
}
