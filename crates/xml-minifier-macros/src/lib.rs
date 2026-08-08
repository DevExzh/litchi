//! Compile-time producer-template macros re-exported by `xml-minifier`.

use proc_macro::{TokenStream, TokenTree};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use quote::quote;
use std::fmt::Write;
use std::fs;
use std::path::Path;

/// Represents a part of a format string.
#[derive(Debug, Clone)]
enum FormatPart {
    /// Static text that doesn't need formatting.
    Static(String),
    /// A format placeholder (either positional index or named argument).
    Placeholder(PlaceholderType),
}

/// Type of format placeholder.
#[derive(Debug, Clone)]
enum PlaceholderType {
    /// Positional argument by index (e.g., {0}, {1}).
    Positional(usize),
    /// Named argument (e.g., {name}).
    Named(String),
    /// Next positional argument (e.g., {}).
    NextPositional,
}

/// Minifies an XML string literal at compile time
///
/// This macro performs deterministic producer-template normalization including:
/// - Removing authoring comments while preserving processing instructions and meaningful text
/// - Rejecting CR/LF/tab-only structural text outside inherited `xml:space="preserve"`
/// - Preserving every accepted character-data event byte-for-byte
/// - Collapsing empty tags (`<tag></tag>` → `<tag/>`)
/// - Requiring source text to contain no non-semantic formatting whitespace
///
/// Unlike [`minified_xml!`], this macro takes an XML string literal directly
/// instead of reading from a file.
///
/// # Examples
///
/// ```ignore
/// const TEMPLATE: &str = minified_xml_str!(r#"<?xml version="1.0"?><root><!-- This comment will be removed --><child attr="value">Some text content</child><empty></empty></root>"#);
/// // Result: <?xml version="1.0"?><root><child attr="value">Some text content</child><empty/></root>
/// ```
#[proc_macro]
pub fn minified_xml_str(input: TokenStream) -> TokenStream {
    match expand_minified_xml_str(input) {
        Ok(tokens) => tokens,
        Err(message) => compile_error(&message),
    }
}

fn expand_minified_xml_str(input: TokenStream) -> Result<TokenStream, String> {
    let xml_content = input_to_string(input)?;

    let minified = minify_xml(&xml_content)
        .map_err(|error| format!("failed to transform XML string literal: {error}"))?;

    let expanded = quote! {
        #minified
    };

    // Generate the output token stream
    Ok(TokenStream::from(expanded))
}

/// Minifies an XML file at compile time and embeds it as a string literal
///
/// This macro performs deterministic producer-template normalization including:
/// - Removing authoring comments while preserving processing instructions and meaningful text
/// - Preserving every character-data event byte-for-byte
/// - Collapsing empty tags (`<tag></tag>` → `<tag/>`)
/// - Requiring source text whitespace to already be intentional and compact
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
    match expand_minified_xml(input) {
        Ok(tokens) => tokens,
        Err(message) => compile_error(&message),
    }
}

fn expand_minified_xml(input: TokenStream) -> Result<TokenStream, String> {
    let file_path = input_to_string(input)?;

    // Get the source file location where the macro was called
    let call_site = proc_macro::Span::call_site();
    let source_file = call_site
        .local_file()
        .ok_or_else(|| "failed to locate the macro call-site file".to_owned())?;
    let parent = source_file
        .parent()
        .ok_or_else(|| "macro call-site file has no parent directory".to_owned())?;
    let target_path = parent.join(Path::new(&file_path));

    // Canonicalize to get absolute path (helps with error messages and change detection)
    let canonical_path = target_path
        .canonicalize()
        .map_err(|error| format!("failed to canonicalize XML file '{file_path}': {error}"))?;

    // Read the XML file
    let xml_content = fs::read_to_string(&canonical_path)
        .map_err(|error| format!("failed to read XML file '{file_path}': {error}"))?;

    // Minify the XML
    let minified = minify_xml(&xml_content)
        .map_err(|error| format!("failed to transform XML file '{file_path}': {error}"))?;

    let tracked_path = canonical_path.to_string_lossy().into_owned();
    let expanded = quote! {{
        // Keep the source XML in rustc's dependency graph so changing it
        // invalidates this expansion instead of silently reusing stale bytes.
        const _: &str = include_str!(#tracked_path);
        #minified
    }};

    // Generate the output token stream
    Ok(TokenStream::from(expanded))
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
/// let xml = minified_xml_format!(r#"<?xml version="{}"?><root><!-- Comment removed --><name>{}</name></root>"#, version, name);
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
    match expand_minified_xml_format(input) {
        Ok(tokens) => tokens,
        Err(message) => compile_error(&message),
    }
}

fn expand_minified_xml_format(input: TokenStream) -> Result<TokenStream, String> {
    // Parse the input tokens
    let tokens: Vec<TokenTree> = input.into_iter().collect();

    if tokens.is_empty() {
        return Err("minified_xml_format! requires at least a format string".to_owned());
    }

    // Extract the format string (first argument)
    let format_str_literal = &tokens[0];
    let TokenTree::Literal(lit) = format_str_literal else {
        return Err("first argument must be a string literal".to_owned());
    };

    let template = syn::parse_str::<syn::LitStr>(&lit.to_string())
        .map_err(|error| format!("invalid format string literal: {error}"))?
        .value();

    // Replace format placeholders with temporary markers before minification
    // This prevents the XML parser from being confused by {} characters
    let (template_with_markers, placeholder_map) = replace_placeholders_with_markers(&template);

    // Minify the XML template
    let minified = minify_xml(&template_with_markers)
        .map_err(|error| format!("failed to transform XML template: {error}"))?;

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
    let parts = parse_format_string(&minified_with_placeholders)?;

    // Generate optimized code
    generate_format_code(&parts, args)
}

fn compile_error(message: &str) -> TokenStream {
    TokenStream::from(quote! { compile_error!(#message) })
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
                    },
                    Some(character) => placeholder_content.push(character),
                    None => {
                        // Unclosed placeholder, just add what we have
                        result.push_str(&placeholder_content);
                        return (result, placeholders);
                    },
                }
            }

            // Create a unique marker
            let index = placeholders.len();
            let marker = format!("__PLACEHOLDER_{index}__");
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
        let marker = format!("__PLACEHOLDER_{idx}__");
        result = result.replace(&marker, placeholder);
    }

    result
}

/// Parse a format string into static parts and placeholders
fn parse_format_string(template: &str) -> Result<Vec<FormatPart>, String> {
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
                    Some(character) => placeholder_content.push(character),
                    None => return Err("unclosed format placeholder in template".to_owned()),
                }
            }

            // Determine placeholder type
            let placeholder = if placeholder_content.is_empty() {
                PlaceholderType::NextPositional
            } else if placeholder_content.chars().all(|c| c.is_ascii_digit()) {
                let index = placeholder_content
                    .parse()
                    .map_err(|error| format!("invalid positional index: {error}"))?;
                PlaceholderType::Positional(index)
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
                return Err("unmatched '}' in format string".to_owned());
            }
        } else {
            current_static.push(ch);
        }
    }

    // Add any remaining static text
    if !current_static.is_empty() {
        parts.push(FormatPart::Static(current_static));
    }

    Ok(parts)
}

/// Generate optimized formatting code
fn generate_format_code(parts: &[FormatPart], args: TokenStream) -> Result<TokenStream, String> {
    use proc_macro::TokenTree as TT;

    // Parse arguments into positional and named
    let mut positional_args = Vec::new();
    let mut named_args = std::collections::HashMap::new();

    let input_tokens: Vec<TT> = args.into_iter().collect();
    let mut i = 0;

    while i < input_tokens.len() {
        // Check if this is a named argument (ident = value)
        if let Some(TT::Ident(name)) = input_tokens.get(i)
            && let Some(TT::Punct(punct)) = input_tokens.get(i + 1)
            && punct.as_char() == '='
        {
            // Named argument
            let name_str = name.to_string();
            let mut value_tokens = Vec::new();
            i += 2; // Skip name and =

            // Collect value tokens until comma or end
            while i < input_tokens.len() {
                if let TT::Punct(p) = &input_tokens[i]
                    && p.as_char() == ','
                {
                    i += 1;
                    break;
                }
                value_tokens.push(input_tokens[i].clone());
                i += 1;
            }

            named_args.insert(name_str, value_tokens);
            continue;
        }

        // Positional argument
        let mut value_tokens = Vec::new();
        while i < input_tokens.len() {
            if let TT::Punct(p) = &input_tokens[i]
                && p.as_char() == ','
            {
                i += 1;
                break;
            }
            value_tokens.push(input_tokens[i].clone());
            i += 1;
        }

        if !value_tokens.is_empty() {
            positional_args.push(value_tokens);
        }
    }

    // Calculate static size
    let static_size: usize = parts
        .iter()
        .filter_map(|p| match p {
            FormatPart::Static(s) => Some(s.len()),
            FormatPart::Placeholder(_) => None,
        })
        .sum();

    // Generate code to build the string - build it as a string to avoid ToTokens issues
    let mut code =
        format!("{{ let mut __result = ::std::string::String::with_capacity({static_size} + 32);");

    let mut next_positional_idx = 0;

    for part in parts {
        match part {
            FormatPart::Static(text) => {
                write!(&mut code, "__result.push_str({text:?});")
                    .map_err(|_| "failed to build static formatter code".to_owned())?;
            },
            FormatPart::Placeholder(placeholder) => {
                let selected_tokens = match placeholder {
                    PlaceholderType::NextPositional => {
                        if let Some(arg) = positional_args.get(next_positional_idx) {
                            next_positional_idx += 1;
                            arg
                        } else {
                            return Err("not enough positional arguments".to_owned());
                        }
                    },
                    PlaceholderType::Positional(idx) => {
                        if let Some(arg) = positional_args.get(*idx) {
                            arg
                        } else {
                            return Err(format!("positional argument {idx} not found"));
                        }
                    },
                    PlaceholderType::Named(name) => {
                        if let Some(arg) = named_args.get(name) {
                            arg
                        } else {
                            return Err(format!("named argument '{name}' not found"));
                        }
                    },
                };

                // Convert the token trees to a string representation
                let arg_str: String = selected_tokens.iter().map(ToString::to_string).collect();

                write!(
                    &mut code,
                    "{{ use ::std::fmt::Write; let _ = write!(&mut __result, \"{{}}\", {arg_str}); }}"
                )
                .map_err(|_| "failed to build placeholder formatter code".to_owned())?;
            },
        }
    }

    code.push_str("__result }");

    // Parse the generated code string back into a TokenStream
    code.parse()
        .map_err(|error| format!("failed to parse generated formatter: {error}"))
}

/// Converts a string literal represented by a `TokenStream` to a Rust string.
///
fn input_to_string(input: TokenStream) -> Result<String, String> {
    syn::parse::<syn::LitStr>(input)
        .map(|literal| literal.value())
        .map_err(|error| format!("expected exactly one string literal: {error}"))
}

/// Normalizes XML tag syntax without changing character-data events.
///
/// This implementation follows the repository's producer-template contract:
/// - Preserves XML declarations
/// - Removes authoring comments while preserving processing instructions and character data
/// - Rejects CR/LF/tab-only structural text outside inherited `xml:space="preserve"`
/// - Preserves every accepted character-data event byte-for-byte
/// - Collapses empty element tags
/// - Handles CDATA sections properly
///
/// # Performance
/// - Borrowed parser events where possible
/// - Single-pass processing
/// - Efficient buffer reuse
fn minify_xml(xml: &str) -> Result<String, Box<dyn std::error::Error>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false); // Text events must remain byte-exact.

    let mut output = Vec::with_capacity(xml.len());
    let mut buf = Vec::new();

    // Stack to track element names for collapsing empty tags
    let mut tag_stack: Vec<BytesStart<'static>> = Vec::new();
    let mut preserve_space = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => {
                if !tag_stack.is_empty() || !preserve_space.is_empty() {
                    return Err(std::io::Error::new(
                        std::io::ErrorKind::InvalidData,
                        "unclosed XML element",
                    )
                    .into());
                }
                break;
            },

            // Preserve XML declaration - write it as-is
            Event::Decl(e) => {
                output.extend_from_slice(b"<?");
                output.extend_from_slice(e.as_ref());
                output.extend_from_slice(b"?>");
            },

            Event::Comment(_) => continue,

            Event::PI(e) => {
                flush_tags(&mut output, &mut tag_stack)?;
                output.extend_from_slice(b"<?");
                output.extend_from_slice(e.as_ref());
                output.extend_from_slice(b"?>");
            },

            // Handle DOCTYPE declarations - preserve them
            Event::DocType(e) => {
                flush_tags(&mut output, &mut tag_stack)?;
                output.extend_from_slice(b"<!DOCTYPE");
                output.push(b' ');
                output.extend_from_slice(e.as_ref());
                output.push(b'>');
            },

            // Handle start tags - buffer them to check if they can be collapsed
            Event::Start(e) => {
                let inherited = preserve_space.last().is_some_and(|value| *value);
                let mut current = inherited;
                for attribute_result in e.attributes() {
                    let attribute = attribute_result?;
                    if attribute.key.as_ref() == b"xml:space" {
                        let value = attribute.decoded_and_normalized_value(
                            quick_xml::XmlVersion::Implicit1_0,
                            reader.decoder(),
                        )?;
                        current = match value.as_ref() {
                            "preserve" => true,
                            "default" => false,
                            _ => inherited,
                        };
                    }
                }
                preserve_space.push(current);
                // Clone the tag for our stack (we need owned data)
                let owned = e.to_owned();
                tag_stack.push(owned);
            },

            // Handle empty tags - flush buffered tags first, then write
            Event::Empty(e) => {
                // Flush all buffered start tags since we have an empty element
                let tags_to_flush = std::mem::take(&mut tag_stack);
                for start_tag in tags_to_flush {
                    output.push(b'<');
                    output.extend_from_slice(start_tag.name().as_ref());
                    write_attributes(&mut output, &start_tag)?;
                    output.push(b'>');
                }

                // Now write the empty tag
                output.push(b'<');
                output.extend_from_slice(e.name().as_ref());
                write_attributes(&mut output, &e)?;
                output.extend_from_slice(b"/>");
            },

            // Handle end tags - check if we can collapse with start tag
            Event::End(e) => {
                preserve_space.pop().ok_or_else(|| {
                    std::io::Error::new(
                        std::io::ErrorKind::InvalidData,
                        "unexpected XML end element",
                    )
                })?;
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

            // Character data is semantic in generic XML. Source templates
            // must already be compact: formatting-only CR/LF/tab text is
            // rejected, never deleted. Accepted text remains byte-exact.
            Event::Text(e) => {
                if !preserve_space.last().is_some_and(|value| *value)
                    && is_formatting_only_text(e.as_ref())
                {
                    return Err(std::io::Error::new(
                        std::io::ErrorKind::InvalidData,
                        "formatting-only XML text containing CR, LF, or tab is not allowed outside xml:space=\"preserve\"; compact the source XML",
                    )
                    .into());
                }
                flush_tags(&mut output, &mut tag_stack)?;
                output.extend_from_slice(e.as_ref());
            },

            // Preserve CDATA sections as-is (they may contain formatting-sensitive content)
            Event::CData(e) => {
                flush_tags(&mut output, &mut tag_stack)?;

                output.extend_from_slice(b"<![CDATA[");
                output.extend_from_slice(e.as_ref());
                output.extend_from_slice(b"]]>");
            },

            Event::GeneralRef(e) => {
                flush_tags(&mut output, &mut tag_stack)?;
                output.push(b'&');
                output.extend_from_slice(e.as_ref());
                output.push(b';');
            },
        }

        buf.clear();
    }

    let result = String::from_utf8(output)?;
    Ok(result)
}

fn flush_tags(
    output: &mut Vec<u8>,
    tags: &mut Vec<BytesStart<'static>>,
) -> Result<(), Box<dyn std::error::Error>> {
    for tag in std::mem::take(tags) {
        output.push(b'<');
        output.extend_from_slice(tag.name().as_ref());
        write_attributes(output, &tag)?;
        output.push(b'>');
    }
    Ok(())
}

/// Helper function to write attributes efficiently
#[inline]
fn write_attributes(
    output: &mut Vec<u8>,
    tag: &BytesStart<'_>,
) -> Result<(), Box<dyn std::error::Error>> {
    let raw = tag.as_ref();
    let mut cursor = tag.name().as_ref().len();
    while cursor < raw.len() {
        while cursor < raw.len() && is_xml_space(raw[cursor]) {
            cursor += 1;
        }
        if cursor == raw.len() {
            break;
        }

        let name_start = cursor;
        while cursor < raw.len() && !is_xml_space(raw[cursor]) && raw[cursor] != b'=' {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < raw.len() && is_xml_space(raw[cursor]) {
            cursor += 1;
        }
        if name_start == name_end || raw.get(cursor) != Some(&b'=') {
            return Err(invalid_attribute().into());
        }
        cursor += 1;
        while cursor < raw.len() && is_xml_space(raw[cursor]) {
            cursor += 1;
        }
        let Some(&quote) = raw.get(cursor) else {
            return Err(invalid_attribute().into());
        };
        if quote != b'\'' && quote != b'"' {
            return Err(invalid_attribute().into());
        }
        let value_start = cursor;
        cursor += 1;
        while cursor < raw.len() && raw[cursor] != quote {
            cursor += 1;
        }
        if cursor == raw.len() {
            return Err(invalid_attribute().into());
        }
        cursor += 1;

        output.push(b' ');
        output.extend_from_slice(&raw[name_start..name_end]);
        output.push(b'=');
        output.extend_from_slice(&raw[value_start..cursor]);
    }
    Ok(())
}

fn invalid_attribute() -> std::io::Error {
    std::io::Error::new(std::io::ErrorKind::InvalidData, "invalid XML attribute")
}

#[inline]
const fn is_xml_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\n' | b'\r')
}

fn is_formatting_only_text(text: &[u8]) -> bool {
    text.iter().copied().all(is_xml_space)
        && text
            .iter()
            .any(|byte| matches!(byte, b'\t' | b'\n' | b'\r'))
}

/// Check if a byte slice contains only whitespace characters
#[cfg(test)]
mod tests {
    use super::*;

    fn parsed(template: &str) -> Vec<FormatPart> {
        parse_format_string(template).unwrap_or_else(|error| panic!("{error}"))
    }

    fn transformed(input: &str) -> String {
        minify_xml(input).unwrap_or_else(|error| panic!("{error}"))
    }

    fn rejected(input: &str) -> String {
        match minify_xml(input) {
            Ok(output) => panic!("expected XML source to be rejected, got {output:?}"),
            Err(error) => error.to_string(),
        }
    }

    #[test]
    fn test_minify_xml_basic() {
        let input = r#"<root><!-- This is a comment --><child attr="value">Text content</child><empty /></root>"#;

        let minified = transformed(input);

        assert!(
            !minified.contains("<!--"),
            "authoring comments should be removed"
        );
        assert!(minified.contains("<root>"), "Root tag should be present");
        assert!(
            minified.contains("Text content"),
            "Text content should be preserved"
        );
    }

    #[test]
    fn test_collapse_empty_tags() {
        let input = r"<root><empty></empty></root>";
        let minified = transformed(input);

        // Empty tags should be collapsed
        assert!(
            minified.contains("<empty/>"),
            "Empty tags should collapse to self-closing: got {minified}"
        );
    }

    #[test]
    fn test_preserve_xml_declaration() {
        let input = r#"<?xml version="1.0" encoding="UTF-8"?><root/>"#;
        let minified = transformed(input);

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
        let input = r"<root><![CDATA[Some <data> with special chars]]></root>";
        let minified = transformed(input);

        assert!(
            minified.contains("<![CDATA[Some <data> with special chars]]>"),
            "CDATA should be preserved as-is"
        );
    }

    #[test]
    fn test_reject_pretty_source_with_compile_error_message() {
        let input = "<root>\n  <child1/>\n  <child2/>\n</root>";
        let error = rejected(input);

        assert!(error.contains("formatting-only XML text"), "{error}");
        assert!(error.contains("xml:space=\"preserve\""), "{error}");
    }

    #[test]
    fn test_preserve_attributes() {
        let input = r#"<root attr1="value1" attr2="value2"/>"#;
        let minified = transformed(input);

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
        let input = r"<root><parent><child>Text here</child></parent></root>";
        let minified = transformed(input);

        assert!(
            minified.contains("Text here"),
            "Text content should be preserved"
        );
        assert_eq!(minified, input);
    }

    #[test]
    fn test_doctype_preservation() {
        let input = r"<!DOCTYPE html><root/>";
        let minified = transformed(input);

        assert!(
            minified.contains("<!DOCTYPE"),
            "DOCTYPE should be preserved"
        );
    }

    #[test]
    fn test_minify_xml_str_with_multiple_features() {
        let input = r#"<?xml version="1.0"?><root><!-- This comment should be removed --><child attr="value">Some text content</child><empty></empty><nested><deep><![CDATA[Some <data> here]]></deep></nested></root>"#;
        let minified = transformed(input);

        assert!(
            !minified.contains("<!--"),
            "authoring comments should be removed"
        );

        // Verify XML declaration preserved
        assert!(
            minified.contains("<?xml"),
            "XML declaration should be preserved"
        );

        // Verify empty tag collapse
        assert!(minified.contains("<empty/>"), "Empty tags should collapse");

        // Verify CDATA preservation
        assert!(
            minified.contains("<![CDATA[Some <data> here]]>"),
            "CDATA should be preserved"
        );

        // Verify text content preserved
        assert!(
            minified.contains("Some text content"),
            "Text content should be preserved"
        );
    }

    #[test]
    fn test_parse_format_string_empty() {
        let parts = parsed("hello world");
        assert_eq!(parts.len(), 1);
        assert!(matches!(parts[0], FormatPart::Static(ref s) if s == "hello world"));
    }

    #[test]
    fn test_parse_format_string_simple_placeholder() {
        let parts = parsed("hello {}");
        assert_eq!(parts.len(), 2);
        assert!(matches!(parts[0], FormatPart::Static(ref s) if s == "hello "));
        assert!(matches!(
            parts[1],
            FormatPart::Placeholder(PlaceholderType::NextPositional)
        ));
    }

    #[test]
    fn test_parse_format_string_indexed_placeholder() {
        let parts = parsed("{0} and {1}");
        assert_eq!(parts.len(), 3);
        assert!(matches!(
            parts[0],
            FormatPart::Placeholder(PlaceholderType::Positional(0))
        ));
        assert!(matches!(parts[1], FormatPart::Static(ref s) if s == " and "));
        assert!(matches!(
            parts[2],
            FormatPart::Placeholder(PlaceholderType::Positional(1))
        ));
    }

    #[test]
    fn test_parse_format_string_named_placeholder() {
        let parts = parsed("Hello {name}!");
        assert_eq!(parts.len(), 3);
        assert!(matches!(parts[0], FormatPart::Static(ref s) if s == "Hello "));
        assert!(
            matches!(parts[1], FormatPart::Placeholder(PlaceholderType::Named(ref n)) if n == "name")
        );
        assert!(matches!(parts[2], FormatPart::Static(ref s) if s == "!"));
    }

    #[test]
    fn test_parse_format_string_escaped_braces() {
        let parts = parsed("{{escaped}} and {} normal");
        assert_eq!(parts.len(), 3);
        assert!(matches!(parts[0], FormatPart::Static(ref s) if s == "{escaped} and "));
        assert!(matches!(
            parts[1],
            FormatPart::Placeholder(PlaceholderType::NextPositional)
        ));
        assert!(matches!(parts[2], FormatPart::Static(ref s) if s == " normal"));
    }

    #[test]
    fn test_parse_format_string_mixed() {
        let parts = parsed("<root><name>{}</name><age>{age}</age></root>");
        assert_eq!(parts.len(), 5);
        assert!(matches!(parts[0], FormatPart::Static(ref s) if s == "<root><name>"));
        assert!(matches!(
            parts[1],
            FormatPart::Placeholder(PlaceholderType::NextPositional)
        ));
        assert!(matches!(parts[2], FormatPart::Static(ref s) if s == "</name><age>"));
        assert!(
            matches!(parts[3], FormatPart::Placeholder(PlaceholderType::Named(ref n)) if n == "age")
        );
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
        let (result, placeholders) =
            replace_placeholders_with_markers("<root><a>{}</a><b>{name}</b></root>");
        assert_eq!(
            result,
            "<root><a>__PLACEHOLDER_0__</a><b>__PLACEHOLDER_1__</b></root>"
        );
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
            &placeholders,
        );
        assert_eq!(result, "<root><a>{}</a><b>{name}</b></root>");
    }

    #[test]
    fn test_minify_xml_nested_structure() {
        let input = "<root><level1><level2><level3>text</level3></level2></level1></root>";
        let minified = transformed(input);
        assert_eq!(
            minified,
            "<root><level1><level2><level3>text</level3></level2></level1></root>"
        );
    }

    #[test]
    fn test_minify_xml_siblings() {
        let input = r"<root><child1>a</child1><child2>b</child2><child3>c</child3></root>";
        let minified = transformed(input);
        assert_eq!(
            minified,
            "<root><child1>a</child1><child2>b</child2><child3>c</child3></root>"
        );
    }

    #[test]
    fn test_preserve_mixed_content_and_inert_events() {
        let input = "<p>Hello <b>world</b> !<?keep x?>&amp;<![CDATA[  ]]></p>";
        let minified = transformed(input);
        assert_eq!(minified, input);
    }

    #[test]
    fn test_preserve_plain_space_nodes_and_mixed_boundaries() {
        for input in ["<p><b>a</b> <i>b</i></p>", "<p>   </p>", "<p>a  b</p>"] {
            let minified = transformed(input);
            assert_eq!(minified, input);
        }
    }

    #[test]
    fn test_reject_newline_only_mixed_content() {
        let input = "<p><b>a</b>\n\t<i>b</i></p>";
        let error = rejected(input);
        assert!(error.contains("formatting-only XML text"), "{error}");
    }

    #[test]
    fn test_preserve_semantic_mixed_content_with_newline() {
        let input = "<p>Hello\n<b>world</b></p>";
        let minified = transformed(input);
        assert_eq!(minified, input);
    }

    #[test]
    fn test_preserve_inherited_xml_space() {
        let input = "<a xml:space=\"preserve\">\n<b><c> </c></b>\t</a>";
        let minified = transformed(input);
        assert_eq!(minified, input);
    }

    #[test]
    fn test_xml_space_default_resets_inherited_preservation() {
        let input = "<a xml:space=\"preserve\"><b xml:space=\"default\">\n</b></a>";
        let error = rejected(input);
        assert!(error.contains("formatting-only XML text"), "{error}");
    }

    #[test]
    fn test_preserve_single_quoted_attribute_containing_double_quote() {
        let input = "<root value='a\"b'/>";
        let minified = transformed(input);
        assert_eq!(minified, input);
    }
}
