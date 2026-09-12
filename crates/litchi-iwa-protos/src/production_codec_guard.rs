//! Shared source ratchet for focused codec production boundaries.

use std::borrow::Cow;

/// Generated-Prost decode/encode operations are retained only in the
/// test-only differential oracles. Keep this list deliberately narrower
/// than a blanket `encode` ban: the archive-header codec has a private Buffa
/// encoder whose `try_encode_bounded` path is production-authorized.
pub(crate) const FORBIDDEN_PROST_CODEC_MARKERS: &[&str] = &[
    "prost::",
    "prost ::",
    "prost as",
    "Message::decode",
    "Message :: decode",
    "Message::merge",
    "Message :: merge",
    "decode_length_delimited",
    "encode_to_vec",
    "try_encode_to_vec",
    "encode_length_delimited",
    "Message::encode",
    "Message :: encode",
    // These are lower-level Prost trait operations. They are not used by
    // the focused codecs today, but keeping their qualified forms here
    // prevents a production call from hiding behind an otherwise innocuous
    // helper name.
    "Message::encode_raw",
    "Message :: encode_raw",
    "Message::encoded_len",
    "Message :: encoded_len",
    "Message::merge_field",
    "Message :: merge_field",
    "Message::merge_length_delimited",
    "Message :: merge_length_delimited",
    "merge_length_delimited",
];

/// Generated Buffa's owned-view surface is not an ingress boundary. The
/// focused codecs may force a borrowed lazy or eager view after their own
/// strict preflight, but they must not retain a
/// `Bytes`-backed `OwnedView`, convert it into an owned generated message, or
/// ask a generated view to allocate an encoded byte container. Keep these
/// markers specific to Buffa's ownership helpers instead of banning every
/// `encode` call: `try_encode_bounded` is the archive-header codec's private,
/// caller-buffered encoder and remains production-authorized.
pub(crate) const FORBIDDEN_BUFFA_OWNERSHIP_MARKERS: &[&str] = &[
    "OwnedView",
    "from_owned",
    "to_owned_message",
    "to_owned_from_source",
    "decode_view_handle",
    "encode_to_bytes",
    "try_encode_to_bytes",
    // Owned-message/eager convenience entry points are not an ingress
    // boundary. A focused codec may force a borrowed `decode_view` only after
    // strict preflight; it must not materialize an owned message or bypass
    // explicit view validation through one of these helpers.
    "decode_from_slice",
    "merge_from_slice",
    "decode_reader",
    "decode_length_delimited_reader",
    "decode_view_with_options",
    "decode_with_options",
    // Keep the generic eager message merge/decode entry points covered too;
    // the existing `decode_length_delimited` marker covers that companion
    // API while these two methods have distinct names. Direct view decoding
    // remains separately audited by the archive ingress test because the
    // private encoder's source validation may use a view without owning it.
    "decode::<",
    "merge::<",
    // Focused codecs consistently construct Buffa options through their
    // private `buffa()` adapter; keep inferred generic calls covered even
    // when the caller omits a turbofish.
    "buffa().decode(",
    "buffa().merge(",
];

/// Return whether the production portion of a codec contains a forbidden
/// generated-code ingress marker.
///
/// Search code bytes only.  A marker in a documentation comment or a fixture
/// string must not turn an otherwise valid production codec into a false
/// positive, while comments/strings are replaced with spaces so a marker
/// cannot be assembled across a removed span.  The build script uses this
/// same predicate for the focused Pages hidden-state boundary.
pub(crate) fn has_forbidden_codec_marker(source: &str) -> bool {
    let production = production_codec_source(source);
    let code = rust_code_only(production.as_ref());
    FORBIDDEN_PROST_CODEC_MARKERS
        .iter()
        .chain(FORBIDDEN_BUFFA_OWNERSHIP_MARKERS)
        .any(|marker| {
            code.windows(marker.len())
                .any(|window| window == marker.as_bytes())
        })
}

/// Replace Rust comments and literals with spaces while preserving all code
/// bytes.  This is intentionally a small lexical filter rather than a line
/// or declaration-order heuristic: the ratchet only needs to know whether a
/// marker occurs in executable/source code, and must remain stable when items
/// are reordered or reformatted.
fn rust_code_only(source: &str) -> Vec<u8> {
    let bytes = source.as_bytes();
    let mut code = bytes.to_vec();
    let mut cursor = 0usize;
    while cursor < bytes.len() {
        if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'/') {
            let start = cursor;
            cursor += 2;
            while cursor < bytes.len() && bytes[cursor] != b'\n' {
                cursor += 1;
            }
            blank_bytes(&mut code, start..cursor);
            continue;
        }
        if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'*') {
            let start = cursor;
            cursor = skip_block_comment(bytes, cursor);
            blank_bytes(&mut code, start..cursor);
            continue;
        }
        if let Some(end) = rust_literal_end(bytes, cursor) {
            blank_bytes(&mut code, cursor..end);
            cursor = end;
            continue;
        }
        cursor += 1;
    }
    code
}

fn blank_bytes(bytes: &mut [u8], range: std::ops::Range<usize>) {
    for byte in bytes.get_mut(range).into_iter().flatten() {
        if *byte != b'\n' && *byte != b'\r' {
            *byte = b' ';
        }
    }
}

fn skip_block_comment(bytes: &[u8], mut cursor: usize) -> usize {
    let mut depth = 0usize;
    while cursor < bytes.len() {
        if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'*') {
            depth = depth.saturating_add(1);
            cursor += 2;
        } else if bytes.get(cursor) == Some(&b'*') && bytes.get(cursor + 1) == Some(&b'/') {
            depth = depth.saturating_sub(1);
            cursor += 2;
            if depth == 0 {
                return cursor;
            }
        } else {
            cursor += 1;
        }
    }
    cursor
}

/// Return the end of a Rust literal beginning at `cursor`.
fn rust_literal_end(bytes: &[u8], cursor: usize) -> Option<usize> {
    let (quote, hashes) = if bytes.get(cursor) == Some(&b'r')
        || (bytes.get(cursor) == Some(&b'b') && bytes.get(cursor + 1) == Some(&b'r'))
    {
        let prefix = if bytes[cursor] == b'b' {
            cursor + 2
        } else {
            cursor + 1
        };
        let mut quote = prefix;
        while bytes.get(quote) == Some(&b'#') {
            quote += 1;
        }
        (bytes.get(quote) == Some(&b'"')).then_some((quote, quote - prefix))?
    } else if bytes.get(cursor) == Some(&b'"') {
        (cursor, 0)
    } else if bytes.get(cursor) == Some(&b'b') && bytes.get(cursor + 1) == Some(&b'"') {
        (cursor + 1, 0)
    } else if bytes.get(cursor) == Some(&b'\'') {
        let end = char_literal_end(bytes, cursor)?;
        return Some(end + 1);
    } else if bytes.get(cursor) == Some(&b'b') && bytes.get(cursor + 1) == Some(&b'\'') {
        let end = char_literal_end(bytes, cursor + 1)?;
        return Some(end + 2);
    } else {
        return None;
    };

    if hashes != 0 || bytes.get(quote) == Some(&b'"') && cursor != quote {
        let mut probe = quote + 1;
        while probe < bytes.len() {
            if bytes[probe] == b'"'
                && bytes
                    .get(probe + 1..probe + 1 + hashes)
                    .is_some_and(|tail| tail.iter().all(|byte| *byte == b'#'))
            {
                return Some(probe + 1 + hashes);
            }
            probe += 1;
        }
        return Some(bytes.len());
    }

    let mut probe = quote + 1;
    while probe < bytes.len() {
        if bytes[probe] == b'\\' {
            probe = probe.saturating_add(2);
        } else if bytes[probe] == b'"' {
            return Some(probe + 1);
        } else {
            probe += 1;
        }
    }
    Some(bytes.len())
}

/// Return the first byte of the item controlled by a test-only `cfg`
/// attribute that starts on `cursor`'s source line.
///
/// The matcher accepts equivalent token spacing and predicates such as
/// `cfg(any(test))` or `cfg(all(test, feature = "..."))`, but deliberately
/// keeps `cfg(any(test, feature = "..."))` in production because that item can
/// be active outside tests. Attribute and comment scanning is lexical, so a
/// bracket in a string or comment cannot terminate the attribute early.
fn cfg_test_attribute_item_start(source: &str, cursor: usize) -> Option<usize> {
    let bytes = source.as_bytes();
    let mut attribute_start = cursor;
    while bytes
        .get(attribute_start)
        .is_some_and(|byte| matches!(byte, b' ' | b'\t' | b'\r'))
    {
        attribute_start += 1;
    }
    if bytes
        .get(attribute_start..attribute_start.saturating_add(2))
        .is_none_or(|prefix| prefix != b"#[")
    {
        return None;
    }
    let attribute_end = rust_attribute_end(bytes, attribute_start)?;
    if !cfg_attribute_is_test_only(&source[attribute_start..attribute_end]) {
        return None;
    }
    Some(skip_rust_trivia(source, attribute_end))
}

fn rust_attribute_end(bytes: &[u8], opening: usize) -> Option<usize> {
    let mut depth = 1usize;
    let mut cursor = opening.saturating_add(2);
    while cursor < bytes.len() {
        if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'/') {
            return None;
        }
        if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'*') {
            cursor = rust_block_comment_end(bytes, cursor);
            if cursor >= bytes.len() {
                return None;
            }
            continue;
        }
        if let Some(end) = rust_literal_end(bytes, cursor) {
            cursor = end;
            continue;
        }
        match bytes[cursor] {
            b'[' => depth = depth.checked_add(1)?,
            b']' => {
                depth = depth.checked_sub(1)?;
                if depth == 0 {
                    return cursor.checked_add(1);
                }
            },
            _ => {},
        }
        cursor += 1;
    }
    None
}

fn cfg_attribute_is_test_only(attribute: &str) -> bool {
    let Some(compact) = compact_cfg_attribute(attribute) else {
        return false;
    };
    let Some(expression) = compact
        .strip_prefix(b"#[cfg(")
        .and_then(|value| value.strip_suffix(b")]"))
    else {
        return false;
    };
    cfg_expression_requires_test(expression)
}

fn compact_cfg_attribute(attribute: &str) -> Option<Vec<u8>> {
    let bytes = attribute.as_bytes();
    let mut compact = Vec::with_capacity(bytes.len());
    let mut cursor = 0usize;
    while cursor < bytes.len() {
        if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'/') {
            return None;
        }
        if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'*') {
            cursor = rust_block_comment_end(bytes, cursor);
            if cursor >= bytes.len() {
                return None;
            }
            continue;
        }
        if let Some(end) = rust_literal_end(bytes, cursor) {
            compact.extend_from_slice(&bytes[cursor..end]);
            cursor = end;
            continue;
        }
        if !bytes[cursor].is_ascii_whitespace() {
            compact.push(bytes[cursor]);
        }
        cursor += 1;
    }
    Some(compact)
}

/// Return whether a cfg expression necessarily implies the `test` cfg.
/// Unknown predicates are treated as ordinary production conditions. This
/// gives `all(test, ...)` the useful test-only behavior while preventing an
/// `any(test, feature = ...)` oracle from being masked.
fn cfg_expression_requires_test(expression: &[u8]) -> bool {
    let mut parser = CfgExpressionParser {
        bytes: expression,
        cursor: 0,
    };
    parser
        .parse_expression()
        .filter(|_| parser.cursor == expression.len())
        .unwrap_or(false)
}

struct CfgExpressionParser<'source> {
    bytes: &'source [u8],
    cursor: usize,
}

impl CfgExpressionParser<'_> {
    fn parse_expression(&mut self) -> Option<bool> {
        let start = self.cursor;
        while self
            .bytes
            .get(self.cursor)
            .is_some_and(|byte| byte.is_ascii_alphanumeric() || *byte == b'_')
        {
            self.cursor += 1;
        }
        if start == self.cursor {
            return None;
        }
        let name = &self.bytes[start..self.cursor];
        if self.bytes.get(self.cursor) == Some(&b'=') {
            self.cursor += 1;
            self.skip_cfg_value()?;
            return Some(false);
        }
        if self.bytes.get(self.cursor) != Some(&b'(') {
            return Some(name == b"test");
        }

        self.cursor += 1;
        let mut saw_argument = false;
        let mut any_requires_test = true;
        let mut all_requires_test = false;
        loop {
            if self.bytes.get(self.cursor) == Some(&b')') {
                self.cursor += 1;
                break;
            }
            let requires_test = self.parse_expression()?;
            saw_argument = true;
            any_requires_test &= requires_test;
            all_requires_test |= requires_test;
            if self.bytes.get(self.cursor) == Some(&b',') {
                self.cursor += 1;
                if self.bytes.get(self.cursor) == Some(&b')') {
                    self.cursor += 1;
                    break;
                }
                continue;
            }
            if self.bytes.get(self.cursor) == Some(&b')') {
                self.cursor += 1;
                break;
            }
            return None;
        }

        match name {
            b"any" => Some(saw_argument && any_requires_test),
            b"all" => Some(saw_argument && all_requires_test),
            b"not" => Some(false),
            _ => Some(false),
        }
    }

    fn skip_cfg_value(&mut self) -> Option<()> {
        let mut nesting = 0usize;
        while let Some(byte) = self.bytes.get(self.cursor).copied() {
            if let Some(end) = rust_literal_end(self.bytes, self.cursor) {
                self.cursor = end;
                continue;
            }
            match byte {
                b'(' | b'[' | b'{' => {
                    nesting = nesting.checked_add(1)?;
                    self.cursor += 1;
                },
                b')' | b',' if nesting == 0 => return Some(()),
                b')' | b']' | b'}' => {
                    nesting = nesting.checked_sub(1)?;
                    self.cursor += 1;
                },
                _ => self.cursor += 1,
            }
        }
        Some(())
    }
}

fn skip_rust_trivia(source: &str, mut cursor: usize) -> usize {
    let bytes = source.as_bytes();
    loop {
        while bytes.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'/') {
            cursor = bytes[cursor..]
                .iter()
                .position(|byte| *byte == b'\n')
                .map_or(bytes.len(), |offset| cursor + offset + 1);
            continue;
        }
        if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'*') {
            let next = rust_block_comment_end(bytes, cursor);
            if next == cursor {
                return cursor;
            }
            cursor = next;
            continue;
        }
        return cursor;
    }
}

fn rust_block_comment_end(bytes: &[u8], mut cursor: usize) -> usize {
    let mut depth = 0usize;
    while cursor < bytes.len() {
        if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'*') {
            depth = depth.saturating_add(1);
            cursor += 2;
        } else if bytes.get(cursor) == Some(&b'*') && bytes.get(cursor + 1) == Some(&b'/') {
            depth = depth.saturating_sub(1);
            cursor += 2;
            if depth == 0 {
                return cursor;
            }
        } else {
            cursor += 1;
        }
    }
    cursor
}

/// Return the production portion of a focused codec source file.
///
/// The generated-Prost builders and differential oracles intentionally live
/// below test-only `cfg` items. Keep those fixtures out of the production
/// ratchet while preserving every production item, including the small
/// test-only allocation probes that a few codecs place near their imports.
pub(crate) fn production_codec_source(source: &str) -> Cow<'_, str> {
    let mut production = String::with_capacity(source.len());
    let mut cursor = 0;
    let mut removed_test_item = false;

    while cursor < source.len() {
        if let Some(item_start) = cfg_test_attribute_item_start(source, cursor) {
            removed_test_item = true;
            cursor = skip_cfg_test_item(source, item_start);
            continue;
        }
        let line_end = source[cursor..]
            .find('\n')
            .map_or(source.len(), |offset| cursor + offset + 1);
        let line = &source[cursor..line_end];
        production.push_str(line);
        cursor = line_end;
    }

    if removed_test_item {
        Cow::Owned(production)
    } else {
        Cow::Borrowed(source)
    }
}

/// Skip one Rust item immediately following a standalone test-only `cfg`
/// attribute. This intentionally handles the item forms used by the codecs:
/// functions, modules, constants, and `thread_local!` blocks. Strings and
/// comments are tokenized so fixture braces do not terminate the item early.
fn skip_cfg_test_item(source: &str, mut cursor: usize) -> usize {
    let bytes = source.as_bytes();
    while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
        cursor += 1;
    }

    let mut braces = 0usize;
    let mut parentheses = 0usize;
    let mut brackets = 0usize;
    let mut saw_brace = false;
    let mut mode = LexMode::Normal;
    while cursor < bytes.len() {
        match mode {
            LexMode::Normal => match bytes[cursor] {
                b'/' if bytes.get(cursor + 1) == Some(&b'/') => {
                    mode = LexMode::LineComment;
                    cursor += 2;
                },
                b'/' if bytes.get(cursor + 1) == Some(&b'*') => {
                    mode = LexMode::BlockComment(1);
                    cursor += 2;
                },
                b'b' if bytes.get(cursor + 1) == Some(&b'"') => {
                    mode = LexMode::String;
                    cursor += 2;
                },
                b'r' | b'b' if raw_string_hashes(bytes, cursor).is_some() => {
                    let hashes = raw_string_hashes(bytes, cursor).expect("checked above");
                    mode = LexMode::RawString(hashes);
                    cursor += if bytes[cursor] == b'b' {
                        3 + hashes
                    } else {
                        2 + hashes
                    };
                },
                b'"' => {
                    mode = LexMode::String;
                    cursor += 1;
                },
                b'\'' => {
                    if let Some(end) = char_literal_end(bytes, cursor) {
                        cursor = end + 1;
                    } else {
                        // A lifetime or label is not a character literal.
                        cursor += 1;
                    }
                },
                b'{' => {
                    saw_brace = true;
                    braces += 1;
                    cursor += 1;
                },
                b'}' if saw_brace => {
                    braces = braces.saturating_sub(1);
                    cursor += 1;
                    if braces == 0 {
                        while bytes.get(cursor).is_some_and(u8::is_ascii_whitespace) {
                            cursor += 1;
                        }
                        if bytes.get(cursor) == Some(&b';') {
                            cursor += 1;
                        }
                        return cursor;
                    }
                },
                b'(' => {
                    parentheses += 1;
                    cursor += 1;
                },
                b')' => {
                    parentheses = parentheses.saturating_sub(1);
                    cursor += 1;
                },
                b'[' => {
                    brackets += 1;
                    cursor += 1;
                },
                b']' => {
                    brackets = brackets.saturating_sub(1);
                    cursor += 1;
                },
                b';' if !saw_brace && parentheses == 0 && brackets == 0 => {
                    return cursor + 1;
                },
                _ => cursor += 1,
            },
            LexMode::LineComment => {
                cursor += 1;
                if bytes.get(cursor.wrapping_sub(1)) == Some(&b'\n') {
                    mode = LexMode::Normal;
                }
            },
            LexMode::BlockComment(mut depth) => {
                if bytes.get(cursor) == Some(&b'/') && bytes.get(cursor + 1) == Some(&b'*') {
                    depth += 1;
                    cursor += 2;
                    mode = LexMode::BlockComment(depth);
                } else if bytes.get(cursor) == Some(&b'*') && bytes.get(cursor + 1) == Some(&b'/') {
                    depth -= 1;
                    cursor += 2;
                    mode = if depth == 0 {
                        LexMode::Normal
                    } else {
                        LexMode::BlockComment(depth)
                    };
                } else {
                    cursor += 1;
                }
            },
            LexMode::String => {
                if bytes[cursor] == b'\\' {
                    cursor = cursor.saturating_add(2);
                } else {
                    let terminator = bytes[cursor] == b'"';
                    cursor += 1;
                    if terminator {
                        mode = LexMode::Normal;
                    }
                }
            },
            LexMode::RawString(hashes) => {
                if bytes[cursor] == b'"'
                    && bytes
                        .get(cursor + 1..cursor + 1 + hashes)
                        .is_some_and(|tail| tail.iter().all(|byte| *byte == b'#'))
                {
                    cursor += 1 + hashes;
                    mode = LexMode::Normal;
                } else {
                    cursor += 1;
                }
            },
        }
    }
    cursor
}

#[derive(Clone, Copy)]
enum LexMode {
    Normal,
    LineComment,
    BlockComment(usize),
    String,
    RawString(usize),
}

fn raw_string_hashes(bytes: &[u8], cursor: usize) -> Option<usize> {
    let prefix = match bytes.get(cursor) {
        Some(b'r') => cursor + 1,
        Some(b'b') if bytes.get(cursor + 1) == Some(&b'r') => cursor + 2,
        _ => return None,
    };
    let mut quote = prefix;
    while bytes.get(quote) == Some(&b'#') {
        quote += 1;
    }
    (bytes.get(quote) == Some(&b'"')).then_some(quote - prefix)
}

/// Return the closing quote for a Rust character literal that starts at
/// `cursor`, if the token is a character literal rather than a lifetime or
/// label.
///
/// The source slicer is intentionally a tiny lexer rather than a Rust parser,
/// but it still needs to distinguish `'{'` and `b'{'` from `'label:`.  A
/// lifetime starts with an identifier and has no immediate closing quote;
/// escaped and punctuation character literals can be scanned until their
/// unescaped quote.  Newlines terminate the scan because they are not valid
/// in Rust character literals.
fn char_literal_end(bytes: &[u8], cursor: usize) -> Option<usize> {
    let first = *bytes.get(cursor + 1)?;
    if first.is_ascii_alphanumeric() || first == b'_' {
        return (bytes.get(cursor + 2) == Some(&b'\'')).then_some(cursor + 2);
    }

    let mut escaped = false;
    let mut probe = cursor + 1;
    while let Some(byte) = bytes.get(probe).copied() {
        if byte == b'\n' || byte == b'\r' {
            return None;
        }
        if escaped {
            escaped = false;
        } else if byte == b'\\' {
            escaped = true;
        } else if byte == b'\'' {
            return Some(probe);
        }
        probe += 1;
    }
    None
}

#[cfg(test)]
mod tests {
    use super::{has_forbidden_codec_marker, production_codec_source};

    const FOCUSED_CODECS: &[(&str, &str)] = &[
        ("archive", include_str!("archive_codec.rs")),
        ("text_storage", include_str!("text_storage_codec.rs")),
        ("hyperlink", include_str!("hyperlink_codec.rs")),
        ("comment_storage", include_str!("comment_storage_codec.rs")),
        (
            "comment_storage_lifecycle",
            include_str!("comment_storage_codec/lifecycle.rs"),
        ),
        (
            "group_node_category",
            include_str!("group_node_category_codec.rs"),
        ),
        (
            "keynote_document",
            include_str!("keynote_document_codec.rs"),
        ),
        (
            "keynote_chart_caption",
            include_str!("keynote_chart_caption_codec.rs"),
        ),
        (
            "keynote_movie_caption",
            include_str!("keynote_movie_caption_codec.rs"),
        ),
        ("movie_playback", include_str!("movie_playback_codec.rs")),
        (
            "keynote_movie_geometry",
            include_str!("keynote_movie_geometry_codec.rs"),
        ),
        (
            "keynote_chart_legend",
            include_str!("keynote_chart_legend_codec.rs"),
        ),
        (
            "chart_arrangement",
            include_str!("chart_arrangement_codec.rs"),
        ),
        ("chart_metadata", include_str!("chart_metadata_codec.rs")),
        ("chart_data", include_str!("chart_data_codec.rs")),
        (
            "chart_data_rewrite",
            include_str!("chart_data_codec/rewrite.rs"),
        ),
        (
            "chart_grid_creation",
            include_str!("chart_grid_creation_codec.rs"),
        ),
        (
            "keynote_chart_title",
            include_str!("keynote_chart_title_codec.rs"),
        ),
        (
            "keynote_chart_axis_title",
            include_str!("keynote_chart_axis_title_codec.rs"),
        ),
        (
            "keynote_chart_axis_value_settings",
            include_str!("keynote_chart_axis_value_settings_codec.rs"),
        ),
        (
            "keynote_placeholder_text",
            include_str!("keynote_placeholder_text_codec.rs"),
        ),
        (
            "keynote_speaker_notes",
            include_str!("keynote_speaker_notes_codec.rs"),
        ),
        (
            "keynote_slide_number",
            include_str!("keynote_slide_number_codec.rs"),
        ),
        (
            "keynote_soundtrack_settings",
            include_str!("keynote_soundtrack_settings_codec.rs"),
        ),
        ("keynote_media", include_str!("keynote_media_codec.rs")),
        (
            "keynote_media_lifecycle",
            include_str!("keynote_media_lifecycle_codec.rs"),
        ),
        (
            "keynote_media_lifecycle_node_cache",
            include_str!("keynote_media_lifecycle_codec/node_cache.rs"),
        ),
        (
            "keynote_slide_transition",
            include_str!("keynote_slide_transition_codec.rs"),
        ),
        (
            "keynote_slide_background",
            include_str!("keynote_slide_background_codec.rs"),
        ),
        ("keynote_show", include_str!("keynote_show_codec.rs")),
        ("numbers_names", include_str!("numbers_names_codec.rs")),
        (
            "numbers_sheet_order",
            include_str!("numbers_sheet_order_codec.rs"),
        ),
        (
            "numbers_table_header_settings",
            include_str!("numbers_table_header_settings_codec.rs"),
        ),
        (
            "numbers_table_title",
            include_str!("numbers_table_title_codec.rs"),
        ),
        (
            "numbers_table_cell_storage",
            include_str!("numbers_table_cell_storage_codec.rs"),
        ),
        (
            "numbers_table_cell_dependency",
            include_str!("numbers_table_cell_dependency_codec.rs"),
        ),
        (
            "numbers_table_cell_dependency_inbound",
            include_str!("numbers_table_cell_dependency_codec/inbound.rs"),
        ),
        (
            "numbers_table_cell_dependency_removal",
            include_str!("numbers_table_cell_dependency_codec/removal.rs"),
        ),
        (
            "package_metadata",
            include_str!("package_metadata_codec.rs"),
        ),
        (
            "package_metadata_component_removal",
            include_str!("package_metadata_codec/component_removal.rs"),
        ),
        ("numbers_formula", include_str!("numbers_formula_codec.rs")),
        ("table_info", include_str!("table_info_codec.rs")),
        ("pages_section", include_str!("pages_section_codec.rs")),
        (
            "pages_header_footer",
            include_str!("pages_header_footer_codec.rs"),
        ),
        (
            "pages_section_background",
            include_str!("pages_section_background_codec.rs"),
        ),
        ("pages_body", include_str!("pages_body_codec.rs")),
        (
            "pages_hidden_state",
            include_str!("pages_hidden_state_codec.rs"),
        ),
        ("pages_media", include_str!("pages_media_codec.rs")),
        ("drawable_parent", include_str!("drawable_parent_codec.rs")),
        (
            "pages_movie_caption",
            include_str!("pages_movie_caption_codec.rs"),
        ),
        ("pages_footnote", include_str!("pages_footnote_codec.rs")),
        (
            "pages_footnote_marker",
            include_str!("pages_footnote_marker_codec.rs"),
        ),
        (
            "pages_page_layout",
            include_str!("pages_page_layout_codec.rs"),
        ),
        (
            "pages_document_settings",
            include_str!("pages_document_settings_codec.rs"),
        ),
    ];

    #[test]
    fn production_ratchet_ignores_test_only_prost_oracle() {
        let source = r#"
pub fn decode_projection(source: &[u8]) {
    let _ = source;
}

#[cfg(test)]
mod oracle {
    use prost::Message as _;

    fn canonical_fixture(bytes: &[u8]) {
        let _ = fixture::Archive::decode(bytes);
        let _ = fixture::Archive::default().encode_to_vec();
        let _ = fixture::ArchiveOwnedView::decode(bytes);
    }
}
"#;

        assert!(!has_forbidden_codec_marker(source));
        assert!(production_codec_source(source).contains("decode_projection"));
        assert!(!production_codec_source(source).contains("canonical_fixture"));
    }

    #[test]
    fn production_ratchet_ignores_cfg_test_items_before_production() {
        let source = r###"
#[cfg(test)]
fn early_oracle() {
    let _fixture = br##"{ prost::Message::decode encode_to_vec }"##;
}

pub fn decode_projection(source: &[u8]) {
    let _ = source;
}

#[cfg(test)]
mod trailing_oracle {}
"###;

        let production = production_codec_source(source);
        assert!(!has_forbidden_codec_marker(source));
        assert!(production.contains("decode_projection"));
        assert!(!production.contains("early_oracle"));
        assert!(!production.contains("trailing_oracle"));
    }

    #[test]
    fn production_ratchet_ignores_cfg_test_items_with_trailing_comments() {
        let source = r###"
#[cfg(test)] // The oracle is never part of a production build.
mod line_commented_oracle {
    use prost::Message as _;
}

#[cfg(test)] /* A block comment may continue below the attribute.
*/
fn block_commented_oracle() {
    let _ = fixture::Archive::decode(&[]);
}

pub fn decode_projection(source: &[u8]) {
    let _ = source;
}
"###;

        let production = production_codec_source(source);
        assert!(!has_forbidden_codec_marker(source));
        assert!(production.contains("decode_projection"));
        assert!(!production.contains("line_commented_oracle"));
        assert!(!production.contains("block_commented_oracle"));
    }

    #[test]
    fn production_ratchet_masks_equivalent_test_only_cfg_forms() {
        let source = r###"
#[ cfg ( any ( test ) )]
fn spaced_oracle() {
    let _ = prost::Message::decode(&[]);
}

#[cfg(
    all(
        test,
        feature = "differential"
    )
)]
fn multiline_oracle() {
    let _ = prost::Message::decode(&[]);
}

#[cfg(any(test, feature = "production"))]
fn feature_oracle() {
    let _ = prost::Message::decode(&[]);
}
"###;

        let production = production_codec_source(source);
        assert!(!production.contains("spaced_oracle"));
        assert!(!production.contains("multiline_oracle"));
        assert!(production.contains("feature_oracle"));
        assert!(has_forbidden_codec_marker(source));
    }

    #[test]
    fn production_ratchet_keeps_items_after_interspersed_cfg_test() {
        let source = r###"
#[cfg(test)]
mod early_oracle {
    use prost::Message as _;
}

use prost::Message as _;

pub fn decode_projection(source: &[u8]) {
    let _ = fixture::Archive::decode(source);
}
"###;

        let production = production_codec_source(source);
        assert!(production.contains("decode_projection"));
        assert!(has_forbidden_codec_marker(source));
    }

    #[test]
    fn production_slicer_handles_character_literals_in_test_items() {
        let source = r###"
#[cfg(test)]
fn fixture_with_wire_literals() {
    let open = b'{';
    let escaped = '\'';
    let newline = '\\n';
    let _ = (open, escaped, newline);
}

pub fn decode_projection(source: &[u8]) {
    let _ = source;
}

#[cfg(test)]
mod oracle {
    fn nested() {
        let close = '}';
        let _ = close;
    }
}
"###;

        let production = production_codec_source(source);
        assert!(production.contains("decode_projection"));
        assert!(!production.contains("fixture_with_wire_literals"));
        assert!(!production.contains("nested"));
    }

    #[test]
    fn production_ratchet_rejects_prost_decode_and_encode() {
        let source = r#"
use prost::Message as _;

pub fn decode_projection(source: &[u8]) {
    let _ = fixture::Archive::decode(source);
    let _ = fixture::Archive::default().encode_to_vec();
}

#[cfg(test)]
mod tests {}
"#;

        assert!(has_forbidden_codec_marker(source));
    }

    #[test]
    fn production_ratchet_rejects_aliased_prost_import() {
        let source = r#"
use prost as p;
use p::Message as _;

pub fn decode_projection(source: &[u8]) {
    let _ = fixture::Archive::decode(source);
}
"#;

        assert!(has_forbidden_codec_marker(source));
    }

    #[test]
    fn production_ratchet_rejects_generated_buffa_encoder() {
        let source = r#"
pub fn decode_projection(source: &[u8]) {
    let _ = crate::buffa_archive_header_generated::LitchiIwaArchiveHeaderProjection::ArchiveInfo::default().try_encode_to_vec();
    let _ = source;
}

#[cfg(test)]
mod tests {}
"#;

        assert!(has_forbidden_codec_marker(source));
    }

    #[test]
    fn production_ratchet_rejects_generated_buffa_ownership_helpers() {
        let helpers = [
            (
                "generated owned-view wrapper",
                "crate::buffa_archive_header_generated::LitchiIwaArchiveHeaderProjection::ArchiveInfoOwnedView::decode(bytes)",
            ),
            (
                "raw owned view",
                "buffa::OwnedView::<ArchiveView>::decode(bytes)",
            ),
            ("owned-view constructor", "view.from_owned(message)"),
            (
                "owned-message conversion",
                "view.to_owned_from_source(None)",
            ),
            (
                "owned-view handle",
                "crate::buffa_archive_header_generated::LitchiIwaArchiveHeaderProjection::ArchiveInfo::decode_view_handle(bytes)",
            ),
            (
                "owned-view handle with options",
                "crate::buffa_archive_header_generated::LitchiIwaArchiveHeaderProjection::ArchiveInfo::decode_view_handle_with_options(bytes, options)",
            ),
            ("generated byte encoder", "view.encode_to_bytes()"),
            (
                "generated fallible byte encoder",
                "view.try_encode_to_bytes()",
            ),
        ];

        for (name, helper) in helpers {
            let source = format!(
                r#"
pub fn decode_projection(bytes: &[u8]) {{
    let _ = {helper};
}}
"#
            );
            assert!(
                has_forbidden_codec_marker(&source),
                "generated Buffa helper {name} escaped the production ratchet"
            );
        }
    }

    #[test]
    fn production_ratchet_rejects_generated_buffa_eager_decoders() {
        let helpers = [
            (
                "message slice decoder",
                "crate::buffa_archive_header_generated::LitchiIwaArchiveHeaderProjection::ArchiveInfo::decode_from_slice(bytes)",
            ),
            (
                "options slice decoder",
                "options.decode_from_slice::<Archive>(bytes)",
            ),
            ("message merge decoder", "view.merge_from_slice(bytes)"),
            ("reader decoder", "options.decode_reader::<Archive>(reader)"),
            (
                "length-delimited reader decoder",
                "options.decode_length_delimited_reader::<Archive>(reader)",
            ),
            (
                "eager view with options",
                "options.decode_view_with_options::<ArchiveView>(bytes)",
            ),
            (
                "eager message with options",
                "options.decode_with_options::<Archive>(bytes)",
            ),
            (
                "eager message decoder",
                "options.decode::<Archive>(&mut input)",
            ),
            (
                "eager length-delimited message decoder",
                "options.decode_length_delimited::<Archive>(&mut input)",
            ),
            (
                "eager message merge",
                "options.merge::<Archive>(&mut message, &mut input)",
            ),
            (
                "inferred eager message decoder",
                "options.buffa().decode(&mut input)",
            ),
            (
                "inferred eager message merge",
                "options.buffa().merge(&mut message, &mut input)",
            ),
        ];

        for (name, helper) in helpers {
            let source = format!(
                r#"
pub fn decode_projection(bytes: &[u8]) {{
    let _ = {helper};
}}
"#
            );
            assert!(
                has_forbidden_codec_marker(&source),
                "generated Buffa eager helper {name} escaped the production ratchet"
            );
        }
    }

    #[test]
    fn production_ratchet_rejects_lower_level_prost_decode_helpers() {
        let source = r#"
pub fn decode_projection(source: &[u8], message: &mut Message) {
    let _ = message.merge_length_delimited(source);
    let _ = Message::merge_field(message, 1, wire, source, context);
    let _ = Message::encode_raw(message, output);
    let _ = Message::encoded_len(message);
}
"#;

        assert!(has_forbidden_codec_marker(source));
    }

    #[test]
    fn production_ratchet_ignores_markers_in_comments_and_literals() {
        let source = r####"
// prost::Message::decode and OwnedView are documentation examples.
pub fn decode_projection(source: &[u8]) {
    let _ = "prost::Message::decode encode_to_vec OwnedView";
    let _ = br##"crate::buffa_generated::OwnedView"##;
    let _ = source;
}
"####;

        assert!(!has_forbidden_codec_marker(source));
    }

    #[test]
    fn production_ratchet_does_not_join_markers_across_comments() {
        let source = r#"
pub fn decode_projection(source: &[u8]) {
    let _ = prost/* a comment */::Message;
    let _ = source;
}
"#;

        assert!(!has_forbidden_codec_marker(source));
    }

    #[test]
    fn production_ratchet_allows_private_buffa_encoder() {
        let source = r#"
pub fn encode_projection(value: &Value, output: &mut Vec<u8>) {
    let _ = value.try_encoded_len();
    let _ = projection::Archive::decode_view(source);
    value.try_encode_bounded(4096, output);
}

#[cfg(test)]
mod oracle {
    use prost::Message as _;

    fn canonical_fixture() {
        let _ = fixture::Archive::default().encode_to_vec();
    }
}
"#;

        assert!(!has_forbidden_codec_marker(source));
    }

    #[test]
    fn focused_codecs_keep_prost_oracles_test_only() {
        for (name, source) in FOCUSED_CODECS {
            assert!(
                !has_forbidden_codec_marker(source),
                "focused codec {name} has a production Prost decode/encode marker"
            );
        }
    }

    #[test]
    fn build_script_tracks_every_codec_source() {
        let build_script = include_str!("../build.rs");
        for (name, _) in FOCUSED_CODECS {
            let source_path = match *name {
                "chart_data_rewrite" => "chart_data_codec/rewrite.rs".to_owned(),
                "chart_grid_creation" => "chart_grid_creation_codec.rs".to_owned(),
                "comment_storage_lifecycle" => "comment_storage_codec/lifecycle.rs".to_owned(),
                "keynote_media_lifecycle_node_cache" => {
                    "keynote_media_lifecycle_codec/node_cache.rs".to_owned()
                },
                "numbers_table_cell_dependency_inbound" => {
                    "numbers_table_cell_dependency_codec/inbound.rs".to_owned()
                },
                "numbers_table_cell_dependency_removal" => {
                    "numbers_table_cell_dependency_codec/removal.rs".to_owned()
                },
                "package_metadata_component_removal" => {
                    "package_metadata_codec/component_removal.rs".to_owned()
                },
                _ => format!("{name}_codec.rs"),
            };
            let rerun_marker = format!("cargo:rerun-if-changed=src/{source_path}");
            assert!(
                build_script.contains(&rerun_marker),
                "build.rs is missing rerun-if-changed for {source_path}"
            );
        }
    }

    #[test]
    fn build_script_pins_numbers_table_schema_provenance() {
        let build_script = include_str!("../build.rs");
        for marker in [
            "fn proto_message_block<'source>(source: &'source str, name: &str)",
            "fn enforce_numbers_table_data_list_provenance(",
            "const TABLE_DATA_LIST_NATIVE_IDS: [u32; 2] = [6005, 6201];",
            "const TABLE_DATA_LIST_SEGMENT_NATIVE_ID: u32 = 6011;",
            "TST_TABLE_DATA_LIST_DIGEST",
            "TST_TABLE_DATA_LIST_SEGMENT_DIGEST",
            "PROJECTION_TABLE_DATA_LIST_SEGMENT_DIGEST",
            "repeated .TST.TableDataList.ListEntry entries = 3;",
            "required .TSP.Range key_range = 2;",
            "key_range_location: u32,",
            "key_range_length: u32,",
            "3 => visitor.visit_list_entry(decode_table_data_list_entry_in(",
            "enforce_table_cell_storage_projection_budget(",
        ] {
            assert!(
                build_script.contains(marker),
                "build.rs lost Numbers table provenance marker: {marker}"
            );
        }
    }

    #[test]
    fn build_script_pins_pages_native_footnote_routes() {
        let build_script = include_str!("../build.rs");
        for marker in [
            "fn enforce_pages_native_message_provenance(",
            "const ROUTE_DECLARATIONS: [(&str, &str, &str, &str); 21]",
            "const PRODUCTION_ROUTE_MARKERS: [(&str, &str, usize); 40]",
            "let root = root_references_with_limits(components, package.state.source.limits())",
            r"let payload = unique_message_payload(\n            &reference.messages,\n            FOOTNOTE_REFERENCE_MESSAGE_TYPE,",
            r"let marker_payload = unique_message_payload(\n            &marker.messages,\n            TEXTUAL_ATTACHMENT_MESSAGE_TYPE,",
            r"let body =\n        find_object(components, body_identifier.get())",
            ".object_mut(native_footnote.reference_identifier.get())",
            ".object_mut(native_footnote.storage_identifier.get())",
            "let marker_identifier = footnote_marker_identifier(",
        ] {
            assert!(
                build_script.contains(marker),
                "build.rs lost Pages native footnote provenance marker: {marker}"
            );
        }
    }

    #[test]
    fn build_script_pins_pages_hidden_state_projection_boundary() {
        let build_script = include_str!("../build.rs");
        let projection = include_str!("buffa-projections/TSTPagesHiddenStateProjection.proto");
        let codec = include_str!("pages_hidden_state_codec.rs");
        let library = include_str!("lib.rs");

        assert_eq!(projection.len(), 3_043);
        assert_eq!(
            projection
                .matches("optional bool needs_to_update_filter_set_for_import = 6;")
                .count(),
            1
        );
        assert!(
            !projection
                .lines()
                .any(|line| line.trim().starts_with("repeated "))
        );
        assert!(!has_forbidden_codec_marker(codec));
        assert!(library.contains("mod buffa_pages_hidden_state_generated {"));

        for marker in [
            "fn enforce_pages_hidden_state_projection_provenance(",
            "const EXPECTED_PROJECTION_BYTES: usize = 3_043;",
            "a95c09076d1d35556287828365a21049f37f839b1289f7f2f591457cf649b713",
            "const EXPECTED_NATIVE_MESSAGE_DIGESTS: [(&str, &str); 13]",
            "const PRIVATE_MODULE_DECLARATION: &str = \"mod buffa_pages_hidden_state_generated {\";",
            "const STRICT_WRAPPER_CALLS: [(&str, &str); 8]",
            "options.buffa().decode_lazy_view(",
            "pub const COLUMN_ROW_UID_MAP_MESSAGE_TYPE: u32 = 6_267;",
            "enforce_pages_hidden_state_projection_budget(",
            "427_026",
            "72c30b51bacf13de1e5948b6d2f7bf986e7a90131e93bc6bd63030d1d558d92a",
        ] {
            assert!(
                build_script.contains(marker),
                "build.rs lost Pages hidden-state provenance marker: {marker}"
            );
        }
    }

    #[test]
    fn keynote_soundtrack_projection_keeps_canonical_scalar_provenance() {
        let build_script = include_str!("../build.rs");
        let canonical = include_str!("protos/KNArchives.proto");
        let projection = include_str!("buffa-projections/KNSoundtrackSettingsArchive.proto");
        let codec = include_str!("keynote_soundtrack_settings_codec.rs");
        let library = include_str!("lib.rs");

        assert_eq!(
            canonical
                .matches("optional .TSP.Reference soundtrack = 17;")
                .count(),
            1
        );
        for field in [
            "optional double volume = 1;",
            "optional .KN.Soundtrack.SoundtrackMode mode = 2 [default = kKNSoundtrackModePlayOnce];",
            "repeated .TSP.DataReference movie_media = 3;",
        ] {
            assert_eq!(
                canonical.matches(field).count(),
                1,
                "canonical field drifted: {field}"
            );
        }
        assert!(projection.contains("message SoundtrackArchive"));
        assert!(projection.contains("optional double volume = 1;"));
        assert!(projection.contains("optional int32 mode = 2 [default = 0];"));
        assert!(!projection.contains("repeated "));
        assert!(codec.contains("projection::SoundtrackArchiveLazyView"));
        assert!(library.contains("mod buffa_keynote_soundtrack_settings_generated"));

        for marker in [
            "fn enforce_keynote_soundtrack_settings_projection_provenance(",
            "projection.contains(\"repeated \")",
            "enforce_archive_header_projection_budget(&buffa_archive_header_out_directory)?;",
            "enforce_data_reference_projection_budget(&buffa_data_reference_out_directory)?;",
            "enforce_keynote_soundtrack_settings_projection_budget(",
            "enforce_table_cell_exact_budget(",
            "27_753",
            "ae5fcc212efd42eca31ff2bafba83032a599cd5f5846009c712996cd8c3ab7e5",
        ] {
            assert!(
                build_script.contains(marker),
                "build.rs lost Keynote soundtrack provenance marker: {marker}"
            );
        }
        assert!(
            !build_script
                .contains("458206e0b57d8ec5ae4c3fc706bf793ccd385ab867b7e92ac30d66ab1858b4d3")
        );
    }
}
