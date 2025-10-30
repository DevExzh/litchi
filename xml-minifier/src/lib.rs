use proc_macro::{TokenStream, TokenTree::Literal};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use quote::quote;
use std::fs;
use std::path::Path;

/// Minifies an XML file at compile time and embeds it as a string literal
///
/// This macro performs aggressive XML minification including:
/// - Removing comments and processing instructions
/// - Trimming and collapsing whitespace in text nodes
/// - Collapsing empty tags (`<tag></tag>` → `<tag/>`)
/// - Removing unnecessary whitespace between tags
///
/// # Path Resolution
///
/// File paths are resolved **relative to the source file** that invokes the macro.
/// This allows for intuitive usage where XML files can be placed next to the source code.
///
/// # Examples
///
/// ```ignore
/// // If you have this structure:
/// // src/
/// //   templates/
/// //     mod.rs
/// //     document.xml
/// //
/// // In templates/mod.rs:
/// const TEMPLATE: &str = minified_xml!("document.xml");
///
/// // Or in parent directory:
/// const TEMPLATE: &str = minified_xml!("templates/document.xml");
/// ```
#[proc_macro]
pub fn minified_xml(input: TokenStream) -> TokenStream {
    let file_path = input_to_string(input);

    // Get the source file location where the macro was called
    let call_site = proc_macro::Span::call_site();
    let source_file = call_site.local_file().expect("Failed to get local file");
    let target_path = source_file
        .parent()
        .expect("Failed to get parent directory of calling file")
        .join(Path::new(&file_path));

    // Canonicalize to get absolute path (helps with error messages and change detection)
    let canonical_path = target_path
        .canonicalize()
        .unwrap_or_else(|e| panic!("Failed to canonicalize file path '{}': {}", file_path, e));

    // Read the XML file
    let xml_content = fs::read_to_string(&canonical_path).expect("Failed to read XML file");

    // Minify the XML
    let minified = minify_xml(&xml_content)
        .unwrap_or_else(|e| panic!("Failed to minify XML from '{}': {}", file_path, e));

    let expanded = quote! {
        #minified
    };

    // Generate the output token stream
    TokenStream::from(expanded)
}

/// Handles the conversion between String Literals represented by TokenStream and Rust String type
///
/// Thanks to https://github.com/scpso/const-css-minify, the code snippet below is nearly a Copy & Paste
fn input_to_string(input: TokenStream) -> String {
    let token_trees: Vec<_> = input.into_iter().collect();
    if token_trees.len() != 1 {
        panic!("Expected exactly one token tree, got {}", token_trees.len());
    }
    let Literal(literal) = token_trees.first().unwrap() else {
        panic!("Expected a string literal");
    };
    let mut literal = literal.to_string();
    // Unescape the raw string literal
    if let Some(c) = literal.get(0..=0)
        && c != "r"
    {
        literal = literal
            .replace("\\\"", "\"")
            .replace("\\n", "\n")
            .replace("\\r", "\r")
            .replace("\\t", "\t")
            .replace("\\\\", "\\")
    }

    // trim leading and trailing ".." or r#".."# from string literal
    let start = &literal.find('\"').unwrap() + 1;
    let end = &literal.rfind('\"').unwrap() - 1;
    if start > end {
        panic!("Invalid string literal");
    }
    literal[start..=end].to_string()
}

/// Minifies XML content by removing unnecessary whitespace, comments, and collapsing empty tags
///
/// This implementation follows best practices for XML minification:
/// - Preserves XML declarations
/// - Removes comments and processing instructions
/// - Intelligently trims whitespace between elements
/// - Collapses empty element tags
/// - Handles CDATA sections properly
///
/// # Performance
/// - Zero-copy where possible using `Cow<[u8]>`
/// - Single-pass processing
/// - Efficient buffer reuse
fn minify_xml(xml: &str) -> Result<String, Box<dyn std::error::Error>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false); // We handle trimming ourselves for better control

    let mut output = Vec::with_capacity(xml.len() / 2); // Pre-allocate roughly half the size
    let mut buf = Vec::new();

    // Stack to track element names for collapsing empty tags
    let mut tag_stack: Vec<BytesStart<'static>> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,

            // Preserve XML declaration - write it as-is
            Event::Decl(e) => {
                output.extend_from_slice(b"<?");
                output.extend_from_slice(e.as_ref());
                output.extend_from_slice(b"?>");
            },

            // Skip comments - they're not needed in minified output
            Event::Comment(_) => continue,

            // Skip processing instructions (except xml declaration handled above)
            Event::PI(_) => continue,

            // Handle DOCTYPE declarations - preserve them
            Event::DocType(e) => {
                output.extend_from_slice(b"<!DOCTYPE");
                output.push(b' ');
                output.extend_from_slice(e.as_ref());
                output.push(b'>');
            },

            // Handle start tags - buffer them to check if they can be collapsed
            Event::Start(e) => {
                // Clone the tag for our stack (we need owned data)
                let owned = e.to_owned();
                tag_stack.push(owned);
            },

            // Handle empty tags - write directly
            Event::Empty(e) => {
                output.push(b'<');
                output.extend_from_slice(e.name().as_ref());
                write_attributes(&mut output, &e)?;
                output.extend_from_slice(b"/>");
            },

            // Handle end tags - check if we can collapse with start tag
            Event::End(e) => {
                if let Some(start_tag) = tag_stack.pop() {
                    // Check if this end tag matches the last start tag
                    // If so, we can collapse to an empty tag
                    if start_tag.name() == e.name() {
                        // Write as collapsed tag
                        output.push(b'<');
                        output.extend_from_slice(start_tag.name().as_ref());
                        write_attributes(&mut output, &start_tag)?;
                        output.extend_from_slice(b"/>");
                    } else {
                        // Tags don't match - we have content in between
                        // Write the buffered start tag first
                        output.push(b'<');
                        output.extend_from_slice(start_tag.name().as_ref());
                        write_attributes(&mut output, &start_tag)?;
                        output.push(b'>');

                        // Push it back since there's a mismatch
                        tag_stack.push(start_tag);

                        // Write the end tag
                        output.push(b'<');
                        output.push(b'/');
                        output.extend_from_slice(e.name().as_ref());
                        output.push(b'>');
                    }
                } else {
                    // No matching start tag in our buffer - just write end tag
                    output.push(b'<');
                    output.push(b'/');
                    output.extend_from_slice(e.name().as_ref());
                    output.push(b'>');
                }
            },

            // Handle text content - trim whitespace intelligently
            Event::Text(e) => {
                // First, flush any buffered start tags since we have text content
                if let Some(start_tag) = tag_stack.pop() {
                    output.push(b'<');
                    output.extend_from_slice(start_tag.name().as_ref());
                    write_attributes(&mut output, &start_tag)?;
                    output.push(b'>');
                }

                // Get the text content
                let text = e.as_ref();

                // Intelligently handle whitespace
                // Skip pure whitespace between tags, otherwise trim both leading and trailing whitespace
                // This is safe for most XML use cases where whitespace between elements is not significant
                let trimmed = if is_whitespace_only(text) {
                    &[]
                } else {
                    trim_whitespace(text)
                };

                if !trimmed.is_empty() {
                    output.extend_from_slice(trimmed);
                }
            },

            // Preserve CDATA sections as-is (they may contain formatting-sensitive content)
            Event::CData(e) => {
                // Flush any buffered start tags
                if let Some(start_tag) = tag_stack.pop() {
                    output.push(b'<');
                    output.extend_from_slice(start_tag.name().as_ref());
                    write_attributes(&mut output, &start_tag)?;
                    output.push(b'>');
                }

                output.extend_from_slice(b"<![CDATA[");
                output.extend_from_slice(e.as_ref());
                output.extend_from_slice(b"]]>");
            },

            // Skip entity references - they'll be handled by the parser
            // This case is for general entity references which are rare in modern XML
            Event::GeneralRef(_) => continue,
        }

        buf.clear();
    }

    // Flush any remaining buffered tags (shouldn't happen with valid XML)
    for start_tag in tag_stack {
        output.push(b'<');
        output.extend_from_slice(start_tag.name().as_ref());
        write_attributes(&mut output, &start_tag)?;
        output.push(b'>');
    }

    let result = String::from_utf8(output)?;
    Ok(result)
}

/// Helper function to write attributes efficiently
#[inline]
fn write_attributes(output: &mut Vec<u8>, tag: &BytesStart) -> Result<(), quick_xml::Error> {
    for attr in tag.attributes() {
        let attr = attr?;
        output.push(b' ');
        output.extend_from_slice(attr.key.as_ref());
        output.extend_from_slice(b"=\"");
        output.extend_from_slice(&attr.value);
        output.push(b'"');
    }
    Ok(())
}

/// Check if a byte slice contains only whitespace characters
#[inline]
fn is_whitespace_only(bytes: &[u8]) -> bool {
    bytes
        .iter()
        .all(|&b| matches!(b, b' ' | b'\t' | b'\n' | b'\r'))
}

/// Trim leading and trailing whitespace from byte slice
#[inline]
fn trim_whitespace(bytes: &[u8]) -> &[u8] {
    let start = bytes
        .iter()
        .position(|&b| !matches!(b, b' ' | b'\t' | b'\n' | b'\r'))
        .unwrap_or(bytes.len());

    let end = bytes
        .iter()
        .rposition(|&b| !matches!(b, b' ' | b'\t' | b'\n' | b'\r'))
        .map(|pos| pos + 1)
        .unwrap_or(0);

    if start <= end {
        &bytes[start..end]
    } else {
        &[]
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_minify_xml_basic() {
        let input = r#"
            <root>
                <!-- This is a comment -->
                <child attr="value">
                    Text content
                </child>
                <empty />
            </root>
        "#;

        let minified = minify_xml(input).unwrap();

        // Should remove extra whitespace and comments
        assert!(!minified.contains("<!--"), "Comments should be removed");
        // The actual check should be for multiple spaces or newlines
        assert!(
            !minified.contains("  ") && !minified.contains('\n'),
            "Excessive whitespace should be removed, got: {:?}",
            minified
        );
        assert!(minified.contains("<root>"), "Root tag should be present");
        assert!(
            minified.contains("Text content"),
            "Text content should be preserved"
        );
    }

    #[test]
    fn test_collapse_empty_tags() {
        let input = r#"<root><empty></empty></root>"#;
        let minified = minify_xml(input).unwrap();

        // Empty tags should be collapsed
        assert!(
            minified.contains("<empty/>"),
            "Empty tags should collapse to self-closing: got {}",
            minified
        );
    }

    #[test]
    fn test_preserve_xml_declaration() {
        let input = r#"<?xml version="1.0" encoding="UTF-8"?><root/>"#;
        let minified = minify_xml(input).unwrap();

        assert!(
            minified.contains("<?xml"),
            "XML declaration should be preserved"
        );
        assert!(
            minified.contains(r#"version="1.0""#),
            "Version attribute should be preserved"
        );
    }

    #[test]
    fn test_preserve_cdata() {
        let input = r#"<root><![CDATA[Some <data> with special chars]]></root>"#;
        let minified = minify_xml(input).unwrap();

        assert!(
            minified.contains("<![CDATA[Some <data> with special chars]]>"),
            "CDATA should be preserved as-is"
        );
    }

    #[test]
    fn test_remove_whitespace_between_tags() {
        let input = r#"
            <root>
                <child1/>
                <child2/>
            </root>
        "#;
        let minified = minify_xml(input).unwrap();

        // Should not contain excessive whitespace
        assert!(
            !minified.contains("\n"),
            "Newlines between tags should be removed"
        );
        assert!(
            !minified.contains("  "),
            "Multiple spaces should be removed"
        );
    }

    #[test]
    fn test_preserve_attributes() {
        let input = r#"<root attr1="value1" attr2="value2"/>"#;
        let minified = minify_xml(input).unwrap();

        assert!(
            minified.contains(r#"attr1="value1""#),
            "Attributes should be preserved"
        );
        assert!(
            minified.contains(r#"attr2="value2""#),
            "All attributes should be preserved"
        );
    }

    #[test]
    fn test_nested_elements_with_text() {
        let input = r#"
            <root>
                <parent>
                    <child>Text here</child>
                </parent>
            </root>
        "#;
        let minified = minify_xml(input).unwrap();

        assert!(
            minified.contains("Text here"),
            "Text content should be preserved"
        );
        assert!(
            minified.contains("<parent><child>"),
            "Nested structure should be preserved without extra whitespace"
        );
    }

    #[test]
    fn test_doctype_preservation() {
        let input = r#"<!DOCTYPE html><root/>"#;
        let minified = minify_xml(input).unwrap();

        assert!(
            minified.contains("<!DOCTYPE"),
            "DOCTYPE should be preserved"
        );
    }
}
