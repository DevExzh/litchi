mod binary;
mod constants;
mod templates;

use crate::formula::ast::MathNode;

/// MTEF parser using proper binary parsing
pub struct MtefParser<'arena> {
    // Arena for lifetime-managed allocations - kept for future use
    #[allow(dead_code)]
    arena: &'arena bumpalo::Bump,
    binary_parser: Option<binary::MtefBinaryParser<'arena>>,
}

impl<'arena> MtefParser<'arena> {
    /// Create a new MTEF parser
    pub fn new(arena: &'arena bumpalo::Bump, data: &'arena [u8]) -> Self {
        let binary_parser = binary::MtefBinaryParser::new(arena, data).ok();
        Self {
            arena,
            binary_parser,
        }
    }

    /// Parse MTEF data into formula nodes
    ///
    /// # Example
    /// ```ignore
    /// let formula = Formula::new();
    /// let parser = MtefParser::new(formula.arena(), mtef_data);
    /// let nodes = parser.parse()?;
    /// ```
    pub fn parse(&mut self) -> Result<Vec<MathNode<'arena>>, MtefError> {
        if let Some(ref mut parser) = self.binary_parser {
            parser.parse()
        } else {
            // Fallback to simple heuristic for invalid MTEF data
            // In a real implementation, this would need access to the actual binary data
            Ok(Vec::new())
        }
    }

    /// Check if the MTEF data is valid and can be parsed
    pub fn is_valid(&self) -> bool {
        self.binary_parser.is_some()
    }

    /// Get MTEF version information if available
    pub fn version_info(&self) -> Option<(u8, u8, u8, u8, u8)> {
        self.binary_parser.as_ref().map(|p| {
            (
                p.mtef_version,
                p.platform,
                p.product,
                p.version,
                p.version_sub,
            )
        })
    }
}

/// Errors that can occur during MTEF parsing
#[derive(Debug)]
pub enum MtefError {
    InvalidFormat(String),
    UnexpectedEof,
    UnknownTag(u8),
    ParseError(String),
}

impl std::fmt::Display for MtefError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            MtefError::InvalidFormat(msg) => write!(f, "Invalid format: {}", msg),
            MtefError::UnexpectedEof => write!(f, "Unexpected end of file"),
            MtefError::UnknownTag(tag) => write!(f, "Unknown tag: {:#x}", tag),
            MtefError::ParseError(msg) => write!(f, "Parse error: {}", msg),
        }
    }
}

impl std::error::Error for MtefError {}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::ast::Formula;
    use smallvec::smallvec;
    use std::borrow::Cow;

    #[test]
    fn test_mtef_parser_creation() {
        let formula = Formula::new();
        let parser = MtefParser::new(formula.arena(), &[0u8; 100]);

        // Should not be valid with random data
        assert!(!parser.is_valid());
    }

    #[test]
    fn test_mtef_parser_with_valid_header() {
        // Create a minimal valid MTEF header with proper structure
        let data = vec![
            // OLE header (28 bytes)
            0x1C, 0x00, // cb_hdr = 28
            0x00, 0x00, 0x02, 0x00, // version = 0x00020000 (little endian)
            0xD3, 0xC2, // format = 0xC2D3
            0x0B, 0x00, 0x00, 0x00, // size = 11 (MTEF header + minimal content)
            0x00, 0x00, 0x00, 0x00, // reserved[0]
            0x00, 0x00, 0x00, 0x00, // reserved[1]
            0x00, 0x00, 0x00, 0x00, // reserved[2]
            0x00, 0x00, 0x00, 0x00, // reserved[3]
            // MTEF header with signature
            0x28, 0x04, 0x6D, 0x74, // signature "(04mt"
            0x05, // version = 5
            0x01, // platform = 1 (Windows)
            0x01, // product = 1 (MathType)
            0x01, // version = 1
            0x00, // version_sub = 0
            0x00, // application_key (empty null-terminated string)
            0x00, // inline = 0
            // Minimal MTEF content (SIZE + END tags)
            0x09, // SIZE tag
            0x00, // END tag
        ];

        let formula = Formula::new();
        let parser = MtefParser::new(formula.arena(), &data);

        // Should be valid with proper headers
        assert!(parser.is_valid());

        if let Some((version, platform, product, ver, sub)) = parser.version_info() {
            assert_eq!(version, 5);
            assert_eq!(platform, 1);
            assert_eq!(product, 1);
            assert_eq!(ver, 1);
            assert_eq!(sub, 0);
        }
    }

    #[test]
    fn test_mtef_parser_invalid_data() {
        let formula = Formula::new();

        // Test with data too short for OLE header
        let parser1 = MtefParser::new(formula.arena(), &[0u8; 10]);
        assert!(!parser1.is_valid());

        // Test with invalid OLE header
        let mut data = vec![0u8; 28];
        data[0] = 0x10; // Invalid cb_hdr
        let parser2 = MtefParser::new(formula.arena(), &data);
        assert!(!parser2.is_valid());
    }

    #[test]
    fn test_character_lookup() {
        // Test Greek letters
        use crate::formula::mtef::binary::charset::lookup_character;

        // Test lowercase alpha (typeface 132, character 97)
        let result = lookup_character(132, 97, 1);
        assert_eq!(result, Some("\\alpha "));

        // Test uppercase Delta (typeface 133, character 68)
        let result = lookup_character(133, 68, 1);
        assert_eq!(result, Some("\\Delta "));

        // Test equals sign (typeface 134, character 61)
        let result = lookup_character(134, 61, 1);
        assert_eq!(result, Some("="));

        // Test non-existent character
        let result = lookup_character(999, 999, 1);
        assert_eq!(result, None);
    }

    #[test]
    fn test_embellishment_templates() {
        use crate::formula::mtef::binary::charset::get_embellishment_template;

        // Test dot embellishment
        let template = get_embellishment_template(2); // embDOT
        assert_eq!(template, "\\dot{%1} ,\\.%1 ");

        // Test hat embellishment
        let template = get_embellishment_template(9); // embHAT
        assert_eq!(template, "\\hat{%1} ,\\^%1 ");

        // Test vector embellishment
        let template = get_embellishment_template(11); // embVEC
        assert_eq!(template, "\\vec{%1} ,%1 ");

        // Test invalid embellishment
        let template = get_embellishment_template(255);
        assert_eq!(template, "");
    }

    #[test]
    fn test_template_parsing() {
        use crate::formula::mtef::templates::TemplateParser;

        // Test template lookup
        let template = TemplateParser::find_template(0, 3); // Fence: angle-both
        assert!(template.is_some());
        let template_def = template.unwrap();
        assert_eq!(template_def.selector, 0);
        assert_eq!(template_def.variation, 3);
        assert!(template_def.template.contains("\\left\\langle"));

        // Test template lookup for fraction
        let template = TemplateParser::find_template(11, 0); // Fraction
        assert!(template.is_some());
        let template_def = template.unwrap();
        assert!(template_def.template.contains("\\frac"));

        // Test non-existent template
        let template = TemplateParser::find_template(255, 0);
        assert!(template.is_none());
    }

    #[test]
    fn test_fence_template_conversion() {
        use crate::formula::mtef::templates::TemplateParser;

        // Test fence type detection
        let fence = TemplateParser::fence_from_selector(1); // Parentheses
        assert_eq!(fence, Some(crate::formula::ast::Fence::Paren));

        let fence = TemplateParser::fence_from_selector(2); // Braces
        assert_eq!(fence, Some(crate::formula::ast::Fence::Brace));

        let fence = TemplateParser::fence_from_selector(3); // Brackets
        assert_eq!(fence, Some(crate::formula::ast::Fence::Bracket));

        let fence = TemplateParser::fence_from_selector(4); // Pipes
        assert_eq!(fence, Some(crate::formula::ast::Fence::Pipe));

        let fence = TemplateParser::fence_from_selector(255); // Invalid
        assert!(fence.is_none());
    }

    #[test]
    fn test_large_operator_conversion() {
        use crate::formula::mtef::templates::TemplateParser;

        // Test large operator type detection
        let op = TemplateParser::large_op_from_selector(15); // Integrals
        assert_eq!(op, Some(crate::formula::ast::LargeOperator::Integral));

        let op = TemplateParser::large_op_from_selector(16); // Sum
        assert_eq!(op, Some(crate::formula::ast::LargeOperator::Sum));

        let op = TemplateParser::large_op_from_selector(17); // Product
        assert_eq!(op, Some(crate::formula::ast::LargeOperator::Product));

        let op = TemplateParser::large_op_from_selector(21); // Integral (single with limits)
        assert_eq!(op, Some(crate::formula::ast::LargeOperator::Integral));

        let op = TemplateParser::large_op_from_selector(255); // Invalid
        assert!(op.is_none());
    }

    #[test]
    fn test_template_ast_parsing() {
        use crate::formula::mtef::templates::TemplateParser;

        // Test fraction template parsing
        let args: smallvec::SmallVec<[smallvec::SmallVec<[crate::formula::ast::MathNode; 8]>; 4]> = smallvec![
            smallvec![crate::formula::ast::MathNode::Number(Cow::Borrowed("1"))],
            smallvec![crate::formula::ast::MathNode::Number(Cow::Borrowed("2"))]
        ];

        let result = TemplateParser::parse_fraction(args[0].to_vec(), args[1].to_vec());

        match result {
            crate::formula::ast::MathNode::Frac {
                numerator,
                denominator,
                ..
            } => {
                assert_eq!(numerator.len(), 1);
                assert_eq!(denominator.len(), 1);
            },
            _ => panic!("Expected fraction node"),
        }
    }

    #[test]
    fn test_charset_attributes() {
        use crate::formula::mtef::binary::charset::get_charset_attributes;

        let attrs = get_charset_attributes(0); // ZERO
        assert_eq!(attrs.math_attr, 1); // Math
        assert!(attrs.do_lookup);
        assert!(attrs.use_codepoint);

        let attrs = get_charset_attributes(1); // TEXT
        assert_eq!(attrs.math_attr, 2); // Force text
        assert!(attrs.do_lookup);
        assert!(attrs.use_codepoint);

        // Test out of bounds
        let attrs = get_charset_attributes(100);
        assert_eq!(attrs.math_attr, 3); // Default to force math
        assert!(attrs.do_lookup);
        assert!(attrs.use_codepoint);
    }
}
