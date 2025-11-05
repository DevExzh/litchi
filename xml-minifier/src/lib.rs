use proc_macro::{TokenStream, TokenTree};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use quote::quote;
use std::fs;
use std::path::Path;

/// Minifies an XML string literal at compile time
///
/// This macro performs aggressive XML minification including:
/// - Removing comments and processing instructions
/// - Trimming and collapsing whitespace in text nodes
/// - Collapsing empty tags (`<tag></tag>` → `<tag/>`)
/// - Removing unnecessary whitespace between tags
///
/// Unlike [`minified_xml!`], this macro takes an XML string literal directly
/// instead of reading from a file.
///
/// # Examples
///
/// ```ignore
/// const TEMPLATE: &str = minified_xml_str!(r#"
///     <?xml version="1.0"?>
///     <root>
///         <!-- This comment will be removed -->
///         <child attr="value">
///             Some text content
///         </child>
///         <empty></empty>
///     </root>
/// "#);
/// // Result: <?xml version="1.0"?><root><child attr="value">Some text content</child><empty/></root>
/// ```
#[proc_macro]
pub fn minified_xml_str(input: TokenStream) -> TokenStream {
    let xml_content = input_to_string(input);

    // Minify the XML
    let minified = minify_xml(&xml_content)
        .unwrap_or_else(|e| panic!("Failed to minify XML string literal: {}", e));

    let expanded = quote! {
        #minified
    };

    // Generate the output token stream
    TokenStream::from(expanded)
}

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
/// For minifying XML string literals directly, see [`minified_xml_str!`].
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

/// Minifies an XML template and formats it with arguments at runtime, with compile-time optimizations
///
/// This macro combines XML minification with optimized string formatting:
/// - Minifies the XML template at compile time
/// - Pre-calculates sizes of static parts
/// - Pre-allocates exact memory needed
/// - Avoids format! macro overhead through direct string building
///
/// The syntax is similar to `format!`, but the template is minified first.
///
/// # Formatting Syntax
///
/// - `{}` - Positional argument (uses `Display` trait)
/// - `{0}`, `{1}`, ... - Indexed positional argument
/// - `{name}` - Named argument
///
/// # Examples
///
/// ```ignore
/// // Basic usage with positional arguments
/// let name = "document";
/// let version = "1.0";
/// let xml = minified_xml_format!(r#"
///     <?xml version="{}"?>
///     <root>
///         <!-- Comment removed -->
///         <name>{}</name>
///     </root>
/// "#, version, name);
/// // Result: <?xml version="1.0"?><root><name>document</name></root>
///
/// // With named arguments
/// let xml = minified_xml_format!(
///     r#"<person><name>{name}</name><age>{age}</age></person>"#,
///     name = "Alice",
///     age = 30
/// );
/// // Result: <person><name>Alice</name><age>30</age></person>
/// ```
#[proc_macro]
pub fn minified_xml_format(input: TokenStream) -> TokenStream {
    // Parse the input tokens
    let tokens: Vec<TokenTree> = input.into_iter().collect();
    
    if tokens.is_empty() {
        panic!("minified_xml_format! requires at least a format string");
    }
    
    // Extract the format string (first argument)
    let format_str_literal = &tokens[0];
    let TokenTree::Literal(lit) = format_str_literal else {
        panic!("First argument must be a string literal");
    };
    
    let template = literal_to_string(lit.to_string());
    
    // Replace format placeholders with temporary markers before minification
    // This prevents the XML parser from being confused by {} characters
    let (template_with_markers, placeholder_map) = replace_placeholders_with_markers(&template);
    
    // Minify the XML template
    let minified = minify_xml(&template_with_markers)
        .unwrap_or_else(|e| panic!("Failed to minify XML template: {}", e));
    
    // Restore the placeholders
    let minified_with_placeholders = restore_placeholders_from_markers(&minified, &placeholder_map);
    
    // Parse the remaining arguments
    let args = if tokens.len() > 1 {
        // Skip the first token (format string) and the comma
        let mut arg_tokens = Vec::new();
        let mut i = 1;
        
        // Skip comma after format string
        if let Some(TokenTree::Punct(p)) = tokens.get(i)
            && p.as_char() == ','
        {
            i += 1;
        }
        
        while i < tokens.len() {
            arg_tokens.push(tokens[i].clone());
            i += 1;
        }
        
        TokenStream::from_iter(arg_tokens)
    } else {
        TokenStream::new()
    };
    
    // Parse the minified template to find format placeholders and static parts
    let parts = parse_format_string(&minified_with_placeholders);
    
    // Generate optimized code
    generate_format_code(&parts, args)
}

/// Replace format placeholders with unique markers that won't confuse the XML parser
/// Returns the modified string and a map of marker -> placeholder
fn replace_placeholders_with_markers(template: &str) -> (String, Vec<String>) {
    let mut result = String::with_capacity(template.len());
    let mut placeholders = Vec::new();
    let mut chars = template.chars().peekable();
    
    while let Some(ch) = chars.next() {
        if ch == '{' {
            // Check for escaped brace {{
            if chars.peek() == Some(&'{') {
                chars.next();
                result.push_str("{{");
                continue;
            }
            
            // Parse the placeholder content
            let mut placeholder_content = String::new();
            placeholder_content.push('{');
            
            loop {
                match chars.next() {
                    Some('}') => {
                        placeholder_content.push('}');
                        break;
                    }
                    Some(ch) => placeholder_content.push(ch),
                    None => {
                        // Unclosed placeholder, just add what we have
                        result.push_str(&placeholder_content);
                        return (result, placeholders);
                    }
                }
            }
            
            // Create a unique marker
            let marker = format!("__PLACEHOLDER_{}__", placeholders.len());
            placeholders.push(placeholder_content);
            result.push_str(&marker);
        } else if ch == '}' {
            // Check for escaped brace }}
            if chars.peek() == Some(&'}') {
                chars.next();
                result.push_str("}}");
            } else {
                result.push('}');
            }
        } else {
            result.push(ch);
        }
    }
    
    (result, placeholders)
}

/// Restore the original placeholders from markers
fn restore_placeholders_from_markers(minified: &str, placeholders: &[String]) -> String {
    let mut result = minified.to_string();
    
    // Replace markers back with original placeholders
    for (idx, placeholder) in placeholders.iter().enumerate() {
        let marker = format!("__PLACEHOLDER_{}__", idx);
        result = result.replace(&marker, placeholder);
    }
    
    result
}

/// Represents a part of a format string
#[derive(Debug, Clone)]
enum FormatPart {
    /// Static text that doesn't need formatting
    Static(String),
    /// A format placeholder (either positional index or named argument)
    Placeholder(PlaceholderType),
}

/// Type of format placeholder
#[derive(Debug, Clone)]
enum PlaceholderType {
    /// Positional argument by index (e.g., {0}, {1})
    Positional(usize),
    /// Named argument (e.g., {name})
    Named(String),
    /// Next positional argument (e.g., {})
    NextPositional,
}

/// Parse a format string into static parts and placeholders
fn parse_format_string(template: &str) -> Vec<FormatPart> {
    let mut parts = Vec::new();
    let mut current_static = String::new();
    let mut chars = template.chars().peekable();
    
    while let Some(ch) = chars.next() {
        if ch == '{' {
            // Check for escaped brace {{
            if chars.peek() == Some(&'{') {
                chars.next();
                current_static.push('{');
                continue;
            }
            
            // Save any accumulated static text
            if !current_static.is_empty() {
                parts.push(FormatPart::Static(current_static.clone()));
                current_static.clear();
            }
            
            // Parse the placeholder content
            let mut placeholder_content = String::new();
            loop {
                match chars.next() {
                    Some('}') => break,
                    Some(ch) => placeholder_content.push(ch),
                    None => panic!("Unclosed format placeholder in template"),
                }
            }
            
            // Determine placeholder type
            let placeholder = if placeholder_content.is_empty() {
                PlaceholderType::NextPositional
            } else if placeholder_content.chars().all(|c| c.is_ascii_digit()) {
                PlaceholderType::Positional(
                    placeholder_content.parse().expect("Invalid positional index")
                )
            } else {
                PlaceholderType::Named(placeholder_content)
            };
            
            parts.push(FormatPart::Placeholder(placeholder));
        } else if ch == '}' {
            // Check for escaped brace }}
            if chars.peek() == Some(&'}') {
                chars.next();
                current_static.push('}');
            } else {
                panic!("Unmatched }} in format string");
            }
        } else {
            current_static.push(ch);
        }
    }
    
    // Add any remaining static text
    if !current_static.is_empty() {
        parts.push(FormatPart::Static(current_static));
    }
    
    parts
}

/// Generate optimized formatting code
fn generate_format_code(parts: &[FormatPart], args: TokenStream) -> TokenStream {
    use proc_macro::TokenTree as TT;
    
    // Parse arguments into positional and named
    let mut positional_args = Vec::new();
    let mut named_args = std::collections::HashMap::new();
    
    let arg_tokens: Vec<TT> = args.into_iter().collect();
    let mut i = 0;
    
    while i < arg_tokens.len() {
        // Check if this is a named argument (ident = value)
        if let Some(TT::Ident(name)) = arg_tokens.get(i)
            && let Some(TT::Punct(punct)) = arg_tokens.get(i + 1)
            && punct.as_char() == '='
        {
            // Named argument
            let name_str = name.to_string();
            let mut value_tokens = Vec::new();
            i += 2; // Skip name and =
            
            // Collect value tokens until comma or end
            while i < arg_tokens.len() {
                if let TT::Punct(p) = &arg_tokens[i]
                    && p.as_char() == ','
                {
                    i += 1;
                    break;
                }
                value_tokens.push(arg_tokens[i].clone());
                i += 1;
            }
            
            named_args.insert(name_str, value_tokens);
            continue;
        }
        
        // Positional argument
        let mut value_tokens = Vec::new();
        while i < arg_tokens.len() {
            if let TT::Punct(p) = &arg_tokens[i]
                && p.as_char() == ','
            {
                i += 1;
                break;
            }
            value_tokens.push(arg_tokens[i].clone());
            i += 1;
        }
        
        if !value_tokens.is_empty() {
            positional_args.push(value_tokens);
        }
    }
    
    // Calculate static size
    let static_size: usize = parts.iter()
        .filter_map(|p| match p {
            FormatPart::Static(s) => Some(s.len()),
            _ => None,
        })
        .sum();
    
    // Generate code to build the string - build it as a string to avoid ToTokens issues
    let mut code = format!(
        "{{ let mut __result = ::std::string::String::with_capacity({} + 32);",
        static_size
    );
    
    let mut next_positional_idx = 0;
    
    for part in parts {
        match part {
            FormatPart::Static(text) => {
                code.push_str(&format!("__result.push_str({:?});", text));
            }
            FormatPart::Placeholder(placeholder) => {
                let arg_tokens = match placeholder {
                    PlaceholderType::NextPositional => {
                        if let Some(arg) = positional_args.get(next_positional_idx) {
                            next_positional_idx += 1;
                            arg
                        } else {
                            panic!("Not enough positional arguments");
                        }
                    }
                    PlaceholderType::Positional(idx) => {
                        if let Some(arg) = positional_args.get(*idx) {
                            arg
                        } else {
                            panic!("Positional argument {} not found", idx);
                        }
                    }
                    PlaceholderType::Named(name) => {
                        if let Some(arg) = named_args.get(name) {
                            arg
                        } else {
                            panic!("Named argument '{}' not found", name);
                        }
                    }
                };
                
                // Convert the token trees to a string representation
                let arg_str: String = arg_tokens.iter()
                    .map(|tt| tt.to_string())
                    .collect::<Vec<_>>()
                    .join("");
                
                code.push_str(&format!(
                    "{{ use ::std::fmt::Write; let _ = write!(&mut __result, \"{{}}\", {}); }}",
                    arg_str
                ));
            }
        }
    }
    
    code.push_str("__result }");
    
    // Parse the generated code string back into a TokenStream
    code.parse().expect("Failed to parse generated code")
}

/// Handles the conversion between String Literals represented by TokenStream and Rust String type
///
/// Thanks to https://github.com/scpso/const-css-minify, the code snippet below is nearly a Copy & Paste
fn input_to_string(input: TokenStream) -> String {
    let token_trees: Vec<_> = input.into_iter().collect();
    if token_trees.len() != 1 {
        panic!("Expected exactly one token tree, got {}", token_trees.len());
    }
    let TokenTree::Literal(literal) = token_trees.first().unwrap() else {
        panic!("Expected a string literal");
    };
    literal_to_string(literal.to_string())
}

/// Convert a literal token string to its actual string value
fn literal_to_string(mut literal: String) -> String {
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
                        // Before writing the collapsed tag, flush all other buffered start tags
                        // This ensures proper nesting: <root><a><b/></a></root> not <b/><a/><root/>
                        let remaining_tags = std::mem::take(&mut tag_stack);
                        for buffered_tag in remaining_tags {
                            output.push(b'<');
                            output.extend_from_slice(buffered_tag.name().as_ref());
                            write_attributes(&mut output, &buffered_tag)?;
                            output.push(b'>');
                        }
                        
                        // Now write the collapsed tag
                        output.push(b'<');
                        output.extend_from_slice(start_tag.name().as_ref());
                        write_attributes(&mut output, &start_tag)?;
                        output.extend_from_slice(b"/>");
                    } else {
                        // Tags don't match - we have content in between
                        // Flush all buffered tags
                        let mut all_tags = std::mem::take(&mut tag_stack);
                        all_tags.push(start_tag);
                        
                        for buffered_tag in all_tags {
                            output.push(b'<');
                            output.extend_from_slice(buffered_tag.name().as_ref());
                            write_attributes(&mut output, &buffered_tag)?;
                            output.push(b'>');
                        }

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

                // Only flush buffered start tags if we have non-whitespace text content
                if !trimmed.is_empty() {
                    // Flush ALL buffered start tags since we have text content
                    // Use mem::take to efficiently move all elements out of the stack
                    let tags_to_flush = std::mem::take(&mut tag_stack);
                    for start_tag in tags_to_flush {
                        output.push(b'<');
                        output.extend_from_slice(start_tag.name().as_ref());
                        write_attributes(&mut output, &start_tag)?;
                        output.push(b'>');
                    }
                    
                    output.extend_from_slice(trimmed);
                }
            },

            // Preserve CDATA sections as-is (they may contain formatting-sensitive content)
            Event::CData(e) => {
                // Flush ALL buffered start tags in correct order
                let tags_to_flush = std::mem::take(&mut tag_stack);
                for start_tag in tags_to_flush {
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

    #[test]
    fn test_minify_xml_str_with_multiple_features() {
        let input = r#"
            <?xml version="1.0"?>
            <root>
                <!-- This comment should be removed -->
                <child attr="value">
                    Some text content
                </child>
                <empty></empty>
                <nested>
                    <deep>
                        <![CDATA[Some <data> here]]>
                    </deep>
                </nested>
            </root>
        "#;
        let minified = minify_xml(input).unwrap();

        // Verify comment removal
        assert!(!minified.contains("<!--"), "Comments should be removed");
        
        // Verify XML declaration preserved
        assert!(minified.contains("<?xml"), "XML declaration should be preserved");
        
        // Verify empty tag collapse
        assert!(minified.contains("<empty/>"), "Empty tags should collapse");
        
        // Verify CDATA preservation
        assert!(minified.contains("<![CDATA[Some <data> here]]>"), "CDATA should be preserved");
        
        // Verify text content preserved
        assert!(minified.contains("Some text content"), "Text content should be preserved");
        
        // Verify whitespace removal between tags
        assert!(!minified.contains("\n"), "Newlines should be removed");
    }

    #[test]
    fn test_parse_format_string_empty() {
        let parts = parse_format_string("hello world");
        assert_eq!(parts.len(), 1);
        assert!(matches!(parts[0], FormatPart::Static(ref s) if s == "hello world"));
    }

    #[test]
    fn test_parse_format_string_simple_placeholder() {
        let parts = parse_format_string("hello {}");
        assert_eq!(parts.len(), 2);
        assert!(matches!(parts[0], FormatPart::Static(ref s) if s == "hello "));
        assert!(matches!(parts[1], FormatPart::Placeholder(PlaceholderType::NextPositional)));
    }

    #[test]
    fn test_parse_format_string_indexed_placeholder() {
        let parts = parse_format_string("{0} and {1}");
        assert_eq!(parts.len(), 3);
        assert!(matches!(parts[0], FormatPart::Placeholder(PlaceholderType::Positional(0))));
        assert!(matches!(parts[1], FormatPart::Static(ref s) if s == " and "));
        assert!(matches!(parts[2], FormatPart::Placeholder(PlaceholderType::Positional(1))));
    }

    #[test]
    fn test_parse_format_string_named_placeholder() {
        let parts = parse_format_string("Hello {name}!");
        assert_eq!(parts.len(), 3);
        assert!(matches!(parts[0], FormatPart::Static(ref s) if s == "Hello "));
        assert!(matches!(parts[1], FormatPart::Placeholder(PlaceholderType::Named(ref n)) if n == "name"));
        assert!(matches!(parts[2], FormatPart::Static(ref s) if s == "!"));
    }

    #[test]
    fn test_parse_format_string_escaped_braces() {
        let parts = parse_format_string("{{escaped}} and {} normal");
        assert_eq!(parts.len(), 3);
        assert!(matches!(parts[0], FormatPart::Static(ref s) if s == "{escaped} and "));
        assert!(matches!(parts[1], FormatPart::Placeholder(PlaceholderType::NextPositional)));
        assert!(matches!(parts[2], FormatPart::Static(ref s) if s == " normal"));
    }

    #[test]
    fn test_parse_format_string_mixed() {
        let parts = parse_format_string("<root><name>{}</name><age>{age}</age></root>");
        assert_eq!(parts.len(), 5);
        assert!(matches!(parts[0], FormatPart::Static(ref s) if s == "<root><name>"));
        assert!(matches!(parts[1], FormatPart::Placeholder(PlaceholderType::NextPositional)));
        assert!(matches!(parts[2], FormatPart::Static(ref s) if s == "</name><age>"));
        assert!(matches!(parts[3], FormatPart::Placeholder(PlaceholderType::Named(ref n)) if n == "age"));
        assert!(matches!(parts[4], FormatPart::Static(ref s) if s == "</age></root>"));
    }

    #[test]
    fn test_replace_placeholders_with_markers_simple() {
        let (result, placeholders) = replace_placeholders_with_markers("<root>{}</root>");
        assert_eq!(result, "<root>__PLACEHOLDER_0__</root>");
        assert_eq!(placeholders, vec!["{}"]);
    }

    #[test]
    fn test_replace_placeholders_with_markers_multiple() {
        let (result, placeholders) = replace_placeholders_with_markers("<root><a>{}</a><b>{name}</b></root>");
        assert_eq!(result, "<root><a>__PLACEHOLDER_0__</a><b>__PLACEHOLDER_1__</b></root>");
        assert_eq!(placeholders, vec!["{}", "{name}"]);
    }

    #[test]
    fn test_replace_placeholders_with_markers_escaped() {
        let (result, placeholders) = replace_placeholders_with_markers("<root>{{escaped}}</root>");
        assert_eq!(result, "<root>{{escaped}}</root>");
        assert_eq!(placeholders.len(), 0);
    }

    #[test]
    fn test_restore_placeholders_from_markers() {
        let placeholders = vec![String::from("{}"), String::from("{name}")];
        let result = restore_placeholders_from_markers(
            "<root><a>__PLACEHOLDER_0__</a><b>__PLACEHOLDER_1__</b></root>",
            &placeholders
        );
        assert_eq!(result, "<root><a>{}</a><b>{name}</b></root>");
    }

    #[test]
    fn test_minify_xml_nested_structure() {
        let input = r#"
            <root>
                <level1>
                    <level2>
                        <level3>text</level3>
                    </level2>
                </level1>
            </root>
        "#;
        let minified = minify_xml(input).unwrap();
        assert_eq!(minified, "<root><level1><level2><level3>text</level3></level2></level1></root>");
    }

    #[test]
    fn test_minify_xml_siblings() {
        let input = r#"<root><child1>a</child1><child2>b</child2><child3>c</child3></root>"#;
        let minified = minify_xml(input).unwrap();
        assert_eq!(minified, "<root><child1>a</child1><child2>b</child2><child3>c</child3></root>");
    }
}
