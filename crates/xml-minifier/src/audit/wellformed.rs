//! Lexical well-formedness rules that quick-xml's tokenizer does not check.
//!
//! quick-xml 0.41 splits a document into tokens and matches end-tag names; it
//! does not check that names are XML names, that references name a declared
//! entity or a legal character, that `<` is absent from attribute values, that
//! `]]>` is absent from character data, where the XML declaration stands or
//! what it says, what a processing-instruction target is, whether a comment
//! contains `--`, or which characters occur at all. These functions check each
//! rule of XML 1.0 (Fifth Edition) and Namespaces in XML 1.0 (Third Edition)
//! on the bytes of one token, so they depend on nothing outside it.

/// Bit of [`NAME_CLASS`]: the byte may start an NCName.
const NAME_START: u8 = 1;
/// Bit of [`NAME_CLASS`]: the byte may continue an NCName.
const NAME_CHAR: u8 = 2;
/// Bit of [`NAME_CLASS`]: the byte is the colon.
const NAME_COLON: u8 = 4;

/// Name classes of every byte. The colon is in neither name class: it is legal
/// in an XML Name, but a namespace-well-formed name has at most one, between
/// two NCNames. Bytes of multi-byte UTF-8 sequences have no class, so a name
/// holding one takes the character-by-character path.
static NAME_CLASS: [u8; 256] = name_classes();

const fn name_classes() -> [u8; 256] {
    let mut classes = [0; 256];
    let mut index = 0;
    while index < 128 {
        let byte = index as u8;
        classes[index] = if byte.is_ascii_alphabetic() || byte == b'_' {
            NAME_START | NAME_CHAR
        } else if byte.is_ascii_digit() || byte == b'-' || byte == b'.' {
            NAME_CHAR
        } else if byte == b':' {
            NAME_COLON
        } else {
            0
        };
        index += 1;
    }
    classes
}

/// `NameStartChar` of XML 1.0 (Fifth Edition), production 4, without `:`.
const fn is_name_start_char(character: char) -> bool {
    matches!(
        character,
        'A'..='Z'
            | '_'
            | 'a'..='z'
            | '\u{C0}'..='\u{D6}'
            | '\u{D8}'..='\u{F6}'
            | '\u{F8}'..='\u{2FF}'
            | '\u{370}'..='\u{37D}'
            | '\u{37F}'..='\u{1FFF}'
            | '\u{200C}'..='\u{200D}'
            | '\u{2070}'..='\u{218F}'
            | '\u{2C00}'..='\u{2FEF}'
            | '\u{3001}'..='\u{D7FF}'
            | '\u{F900}'..='\u{FDCF}'
            | '\u{FDF0}'..='\u{FFFD}'
            | '\u{10000}'..='\u{EFFFF}'
    )
}

/// `NameChar` of XML 1.0 (Fifth Edition), production 4a, without `:`.
const fn is_name_char(character: char) -> bool {
    is_name_start_char(character)
        || matches!(
            character,
            '-' | '.' | '0'..='9' | '\u{B7}' | '\u{300}'..='\u{36F}' | '\u{203F}'..='\u{2040}'
        )
}

/// Whether `name` is an NCName: an XML Name that contains no colon.
pub(super) fn is_ncname(name: &[u8]) -> bool {
    let Some(&first) = name.first() else {
        return false;
    };
    if first < 0x80 && NAME_CLASS[usize::from(first)] & NAME_START == 0 {
        return false;
    }
    for (index, &byte) in name.iter().enumerate() {
        if byte >= 0x80 {
            return is_unicode_ncname(name);
        }
        if index > 0 && NAME_CLASS[usize::from(byte)] & NAME_CHAR == 0 {
            return false;
        }
    }
    true
}

fn is_unicode_ncname(name: &[u8]) -> bool {
    let Ok(text) = core::str::from_utf8(name) else {
        return false;
    };
    let mut characters = text.chars();
    characters.next().is_some_and(is_name_start_char) && characters.all(is_name_char)
}

/// Whether `name` is an XML Name, in which colons are ordinary name
/// characters (XML 1.0 production 5). Used only to word a diagnostic.
fn is_xml_name(name: &[u8]) -> bool {
    let Ok(text) = core::str::from_utf8(name) else {
        return false;
    };
    let mut characters = text.chars();
    characters
        .next()
        .is_some_and(|character| character == ':' || is_name_start_char(character))
        && characters.all(|character| character == ':' || is_name_char(character))
}

/// Scans the name that starts at `bytes[start]` and ends before the first
/// XML whitespace, or before the first `=` when `equals_ends` is set, and
/// checks that it is a QName of Namespaces in XML 1.0: an NCName, or two
/// NCNames joined by one colon. Returns where the name ends and the verdict:
/// the colon's index within the name, or the diagnostic for a name that is
/// not an XML name, or that is one but is not namespace-well-formed.
///
/// This is one pass over the bytes a name scan reads anyway: an ASCII name
/// is decided on the way, and only a name with another byte, or a defect, is
/// decided again character by character.
#[inline]
pub(super) fn scan_qname(
    bytes: &[u8],
    start: usize,
    equals_ends: bool,
) -> (usize, Result<Option<usize>, &'static str>) {
    let class = |cursor: usize| {
        bytes
            .get(cursor)
            .map_or(0, |byte| NAME_CLASS[usize::from(*byte)])
    };
    // The common case: one or two ASCII NCNames joined by a colon.
    let mut cursor = start;
    let mut colon = None;
    // Whether the last NCName has at least its start character.
    let mut complete;
    loop {
        if class(cursor) & NAME_START == 0 {
            complete = false;
            break;
        }
        cursor += 1;
        while class(cursor) & NAME_CHAR != 0 {
            cursor += 1;
        }
        complete = true;
        if colon.is_some() || class(cursor) != NAME_COLON {
            break;
        }
        colon = Some(cursor - start);
        cursor += 1;
    }
    let ends = bytes
        .get(cursor)
        .is_none_or(|&byte| is_space(byte) || (equals_ends && byte == b'='));
    if complete && ends {
        return (cursor, Ok(colon));
    }
    // Anything else: find where the name ends and decide it again.
    let end = bytes[start..]
        .iter()
        .position(|&byte| is_space(byte) || (equals_ends && byte == b'='))
        .map_or(bytes.len(), |length| start + length);
    (end, qname_colon_by_character(&bytes[start..end]))
}

/// [`scan_qname`]'s verdict on a name that is not a well-formed ASCII name:
/// decides it character by character and words the diagnostic.
#[cold]
fn qname_colon_by_character(name: &[u8]) -> Result<Option<usize>, &'static str> {
    let colon = name.iter().position(|byte| *byte == b':');
    let valid = match colon {
        None => is_ncname(name),
        Some(colon) => is_ncname(&name[..colon]) && is_ncname(&name[colon + 1..]),
    };
    if valid {
        Ok(colon)
    } else if colon.is_some() && is_xml_name(name) {
        Err("name is not namespace-well-formed: a qualified name has one colon between two names")
    } else {
        Err("invalid XML name")
    }
}

/// `Char` of XML 1.0 (Fifth Edition), production 2.
pub(super) const fn is_xml_char(value: u32) -> bool {
    matches!(
        value,
        0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
    )
}

/// Checks the text of a reference, between its `&` and its `;`.
///
/// A reference is well-formed when it names one of the five predefined
/// entities or is a character reference to a legal XML character. Any other
/// entity is undeclared, because the audit refuses every DTD and therefore
/// every entity declaration (XML 1.0 well-formedness constraint "Entity
/// Declared").
///
/// # Errors
///
/// Returns the diagnostic for a malformed reference, a reference to an
/// undeclared entity, or a character reference to a character XML excludes.
pub(super) fn check_reference(content: &[u8]) -> Result<(), &'static str> {
    match content {
        b"lt" | b"gt" | b"amp" | b"apos" | b"quot" => Ok(()),
        [b'#', b'x', digits @ ..] => character_reference(digits, 16),
        [b'#', digits @ ..] => character_reference(digits, 10),
        name if is_ncname(name) => Err("reference to an undeclared entity"),
        _ => Err("malformed reference"),
    }
}

/// The character a well-formed reference stands for.
///
/// Callers pass only text [`check_reference`] accepted.
fn referenced_character(content: &[u8]) -> Option<char> {
    let value = match content {
        b"lt" => return Some('<'),
        b"gt" => return Some('>'),
        b"amp" => return Some('&'),
        b"apos" => return Some('\''),
        b"quot" => return Some('"'),
        [b'#', b'x', digits @ ..] => reference_value(digits, 16)?,
        [b'#', digits @ ..] => reference_value(digits, 10)?,
        _ => return None,
    };
    char::from_u32(value)
}

fn character_reference(digits: &[u8], radix: u32) -> Result<(), &'static str> {
    let value = reference_value(digits, radix).ok_or("malformed character reference")?;
    if is_xml_char(value) {
        Ok(())
    } else {
        Err("character reference to a character XML does not allow")
    }
}

/// The value of a character reference's digits, saturated above the
/// Unicode range; `None` when a digit is missing or invalid.
fn reference_value(digits: &[u8], radix: u32) -> Option<u32> {
    if digits.is_empty() {
        return None;
    }
    digits.iter().try_fold(0_u32, |value, digit| {
        let digit = char::from(*digit).to_digit(radix)?;
        Some(value.saturating_mul(radix).saturating_add(digit))
    })
}

/// Checks the content of an attribute value, between its quotes: no `<`, and
/// every `&` begins a well-formed reference (XML 1.0 production 10).
///
/// # Errors
///
/// Returns the offset within `value` and the diagnostic of the first defect.
pub(super) fn check_attribute_value(value: &[u8]) -> Result<(), (usize, &'static str)> {
    let mut cursor = 0;
    while let Some(found) = value[cursor..]
        .iter()
        .position(|byte| matches!(byte, b'<' | b'&'))
    {
        let at = cursor + found;
        if value[at] == b'<' {
            return Err((at, "'<' is not allowed in an attribute value"));
        }
        let Some(length) = value[at + 1..]
            .iter()
            .position(|byte| matches!(byte, b';' | b'&' | b'<'))
            .filter(|length| value[at + 1 + length] == b';')
        else {
            return Err((at, "unterminated reference in an attribute value"));
        };
        check_reference(&value[at + 1..at + 1 + length]).map_err(|detail| (at, detail))?;
        cursor = at + length + 2;
    }
    Ok(())
}

/// Appends the normalized form of an attribute value that
/// [`check_attribute_value`] accepted: end-of-line handling, then CDATA
/// attribute-value normalization (XML 1.0 sections 2.11 and 3.3.3).
///
/// A literal tab, line feed or carriage return, and a CR-LF pair, become one
/// space; each reference becomes the character it stands for, which is not
/// normalized further.
pub(super) fn push_normalized_value(value: &[u8], out: &mut Vec<u8>) {
    let mut cursor = 0;
    while cursor < value.len() {
        match value[cursor] {
            b'\r' => {
                out.push(b' ');
                cursor += if value.get(cursor + 1) == Some(&b'\n') {
                    2
                } else {
                    1
                };
            },
            b'\t' | b'\n' => {
                out.push(b' ');
                cursor += 1;
            },
            b'&' => {
                let end = value[cursor..]
                    .iter()
                    .position(|byte| *byte == b';')
                    .map_or(value.len(), |length| cursor + length);
                if let Some(character) = referenced_character(&value[cursor + 1..end]) {
                    let mut encoded = [0; 4];
                    out.extend_from_slice(character.encode_utf8(&mut encoded).as_bytes());
                }
                cursor = end + 1;
            },
            byte => {
                out.push(byte);
                cursor += 1;
            },
        }
    }
}

/// Offset of the first `]]>` in `bytes` at or after `from`, or `bytes.len()`.
///
/// `]]>` may close a CDATA section and may stand in comments, processing
/// instructions and attribute values; only in character data is it refused
/// (XML 1.0 production 14). A scan asks for the next one only when a text
/// token reaches the last one found, so the search moves forward over the
/// input once, skipping to each `]`, which markup rarely contains.
pub(super) fn next_cdata_close(bytes: &[u8], from: usize) -> usize {
    let mut cursor = from;
    while let Some(found) = bytes.get(cursor..).and_then(find_bracket) {
        let at = cursor + found;
        if bytes.get(at..at + 3) == Some(b"]]>".as_slice()) {
            return at;
        }
        cursor = at + 1;
    }
    bytes.len()
}

fn find_bracket(bytes: &[u8]) -> Option<usize> {
    // Text before the bracket need not be UTF-8 here; search it as bytes in
    // word-sized steps.
    const WORD: usize = size_of::<usize>();
    const ONES: usize = usize::MAX / 255;
    let pattern = ONES * usize::from(b']');
    let mut index = 0;
    while index + WORD <= bytes.len() {
        let mut word = [0; WORD];
        word.copy_from_slice(&bytes[index..index + WORD]);
        let difference = usize::from_ne_bytes(word) ^ pattern;
        if difference.wrapping_sub(ONES) & !difference & (ONES << 7) != 0 {
            break;
        }
        index += WORD;
    }
    bytes[index..]
        .iter()
        .position(|byte| *byte == b']')
        .map(|found| index + found)
}

/// Checks a comment token `<!--…-->`: its text contains no `--` and does not
/// end with `-` (XML 1.0 production 15).
///
/// # Errors
///
/// Returns the offset within the token and the diagnostic.
pub(super) fn check_comment(raw: &[u8]) -> Result<(), (usize, &'static str)> {
    let text = raw.get(4..raw.len().saturating_sub(3)).unwrap_or_default();
    if let Some(at) = text.windows(2).position(|window| window == b"--") {
        return Err((4 + at, "'--' is not allowed in a comment"));
    }
    if text.last() == Some(&b'-') {
        return Err((3 + text.len(), "a comment must not end with '-'"));
    }
    Ok(())
}

/// Checks a processing-instruction token `<?target …?>`: its target is an
/// NCName (XML 1.0 production 17, Namespaces in XML 1.0 section 7) other than
/// `xml` in any case, and whitespace separates it from any content.
///
/// # Errors
///
/// Returns the offset within the token and the diagnostic.
pub(super) fn check_processing_instruction(raw: &[u8]) -> Result<(), (usize, &'static str)> {
    let inner = raw.get(2..raw.len().saturating_sub(2)).unwrap_or_default();
    let length = inner
        .iter()
        .position(|byte| is_space(*byte))
        .unwrap_or(inner.len());
    let target = &inner[..length];
    if !is_ncname(target) {
        return Err((2, "invalid processing-instruction target"));
    }
    if target.eq_ignore_ascii_case(b"xml") {
        return Err((2, "processing-instruction target 'xml' is reserved"));
    }
    Ok(())
}

/// Checks the grammar of an XML declaration token `<?xml …?>` (XML 1.0
/// productions 23 to 26, 32, 80 and 81): a version `1.` followed by digits,
/// then optionally an encoding, then optionally `standalone` with `yes` or
/// `no`, in that order and nothing else.
///
/// The audit reads its input as UTF-8, so an encoding declaration must name
/// UTF-8 (compared without regard to ASCII case). OPC rule M1.17 (ECMA-376
/// Part 2, section 6.2.5 a) in the fifth edition) forbids a declaration that
/// names any encoding other than UTF-8 or UTF-16, whatever bytes follow it, so
/// `US-ASCII` or `ISO-8859-1` is refused even over bytes that are all ASCII. A
/// declaration of UTF-16 over the UTF-8 this audit has read presents the
/// entity in an encoding other than the one it declares, which XML 1.0
/// section 4.3.3 makes a fatal error.
///
/// # Errors
///
/// Returns the offset within the token and the diagnostic.
pub(super) fn check_declaration_grammar(raw: &[u8]) -> Result<(), (usize, &'static str)> {
    // `raw` is `<?xml` followed by whitespace or `?>`, as quick-xml decides.
    let inner_end = raw.len().saturating_sub(2);
    let mut cursor = 5;
    let mut pseudo = || next_pseudo_attribute(raw, &mut cursor, inner_end);

    let Some((name, value, at)) = pseudo()? else {
        return Err((2, "XML declaration must declare a version"));
    };
    if name != b"version" {
        return Err((at, "XML declaration must declare its version first"));
    }
    if !matches!(value, [b'1', b'.', digits @ ..] if !digits.is_empty() && digits.iter().all(u8::is_ascii_digit))
    {
        return Err((at, "XML declaration version must be 1.x"));
    }

    let mut next = pseudo()?;
    if let Some((b"encoding", value, at)) = next {
        let valid_name = value.first().is_some_and(u8::is_ascii_alphabetic)
            && value
                .iter()
                .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'));
        if !valid_name {
            return Err((at, "invalid encoding name in the XML declaration"));
        }
        if !value.eq_ignore_ascii_case(b"UTF-8") {
            return Err((at, "XML declaration names an encoding other than UTF-8"));
        }
        next = pseudo()?;
    }
    if let Some((b"standalone", value, at)) = next {
        if value != b"yes" && value != b"no" {
            return Err((at, "XML declaration standalone must be 'yes' or 'no'"));
        }
        next = pseudo()?;
    }
    match next {
        None => Ok(()),
        Some((_, _, at)) => Err((at, "unexpected attribute in the XML declaration")),
    }
}

type PseudoAttribute<'a> = (&'a [u8], &'a [u8], usize);

/// The next `name = "value"` of an XML declaration, starting at `cursor`, with
/// the offset of its name; `None` at the declaration's end.
fn next_pseudo_attribute<'a>(
    raw: &'a [u8],
    cursor: &mut usize,
    end: usize,
) -> Result<Option<PseudoAttribute<'a>>, (usize, &'static str)> {
    let separated = *cursor;
    while *cursor < end && is_space(raw[*cursor]) {
        *cursor += 1;
    }
    if *cursor == end {
        return Ok(None);
    }
    if *cursor == separated {
        return Err((
            *cursor,
            "XML declaration attributes must be separated by whitespace",
        ));
    }
    let name_start = *cursor;
    while *cursor < end && !is_space(raw[*cursor]) && raw[*cursor] != b'=' {
        *cursor += 1;
    }
    let name = &raw[name_start..*cursor];
    while *cursor < end && is_space(raw[*cursor]) {
        *cursor += 1;
    }
    if *cursor == end || raw[*cursor] != b'=' {
        return Err((name_start, "XML declaration attribute must have a value"));
    }
    *cursor += 1;
    while *cursor < end && is_space(raw[*cursor]) {
        *cursor += 1;
    }
    if *cursor == end || !matches!(raw[*cursor], b'"' | b'\'') {
        return Err((name_start, "XML declaration attribute value must be quoted"));
    }
    let quote = raw[*cursor];
    let value_start = *cursor + 1;
    let Some(length) = raw[value_start..end].iter().position(|byte| *byte == quote) else {
        return Err((name_start, "unterminated XML declaration attribute value"));
    };
    *cursor = value_start + length + 1;
    Ok(Some((
        name,
        &raw[value_start..value_start + length],
        name_start,
    )))
}

/// What one pass over a whole input finds.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) struct Characters {
    /// The first character XML does not allow: a C0 control other than tab,
    /// line feed and carriage return, or U+FFFE or U+FFFF. Surrogate code
    /// points are not valid UTF-8, so UTF-8 validation refuses them first.
    pub(super) illegal: Option<usize>,
    /// Where the first `]]>` starts, or the input's length. Only one before
    /// `illegal` is looked for, since a scan never reaches past `illegal`.
    pub(super) cdata_close: usize,
}

/// Finds the first character XML does not allow and the first `]]>` in one
/// pass over `bytes`.
///
/// Blocks free of every candidate byte, the common case, are skipped with a
/// branch-free pass over a fixed-size block, which the compiler vectorizes;
/// the function is kept out of line so that it is compiled on its own.
#[inline(never)]
pub(super) fn scan_characters(bytes: &[u8]) -> Characters {
    let mut illegal = None;
    let mut cdata_close = None;
    let (blocks, tail) = bytes.as_chunks::<64>();
    for (number, block) in blocks.iter().enumerate() {
        let (suspect, bracket) = block_flags(block);
        let base = number * 64;
        if bracket && cdata_close.is_none() && closes_cdata(bytes, base) {
            cdata_close = find_cdata_close(bytes, base.saturating_sub(1), base + 64);
        }
        if suspect && let Some(found) = precise_illegal_character(bytes, base, base + 64) {
            illegal = Some(found);
            break;
        }
    }
    if illegal.is_none() {
        let base = bytes.len() - tail.len();
        if cdata_close.is_none() {
            // A `]]>` whose `]>` starts in the tail starts at most one byte
            // before it.
            cdata_close = find_cdata_close(bytes, base.saturating_sub(1), bytes.len());
        }
        illegal = precise_illegal_character(bytes, base, bytes.len());
    }
    Characters {
        illegal,
        cdata_close: cdata_close.unwrap_or(bytes.len()),
    }
}

/// Whether a block holds a C0 control other than tab, line feed and carriage
/// return, or a byte 0xEF, the lead byte of U+FFFE and U+FFFF; and whether it
/// holds a `]`, which markup rarely contains.
#[inline]
fn block_flags(block: &[u8; 64]) -> (bool, bool) {
    let mut suspect = 0_u8;
    let mut bracket = 0_u8;
    for &byte in block {
        let allowed = u8::from(byte == b'\t') | u8::from(byte == b'\n') | u8::from(byte == b'\r');
        suspect |= (u8::from(byte < 0x20) & !allowed) | u8::from(byte == 0xEF);
        bracket |= u8::from(byte == b']');
    }
    (suspect != 0, bracket != 0)
}

/// Whether a `]>`, the end of every `]]>`, starts in the 64-byte block at
/// `base`, its `>` possibly the first byte after the block. A bracket in a
/// reference or a formula rarely comes before a tag close.
///
/// The block holds the pair's `]`, so the caller knows it holds a bracket;
/// and a `]]>` whose `]>` starts here starts at most one byte before it.
fn closes_cdata(bytes: &[u8], base: usize) -> bool {
    let Some(window) = bytes
        .get(base..base + 65)
        .and_then(|window| <&[u8; 65]>::try_from(window).ok())
    else {
        // The last block of the input has no byte after it.
        return bytes
            .get(base..)
            .is_some_and(|rest| rest.windows(2).any(|pair| pair == b"]>"));
    };
    let (Some(bracket), Some(close)) = (window.first_chunk::<64>(), window.last_chunk::<64>())
    else {
        return false;
    };
    let mut found = 0_u8;
    for (&bracket, &close) in bracket.iter().zip(close) {
        found |= u8::from(bracket == b']') & u8::from(close == b'>');
    }
    found != 0
}

/// The first character XML does not allow that starts in `bytes[start..end]`.
#[cold]
fn precise_illegal_character(bytes: &[u8], start: usize, end: usize) -> Option<usize> {
    (start..end).find(|&index| {
        let byte = bytes[index];
        (byte < 0x20 && !matches!(byte, b'\t' | b'\n' | b'\r'))
            || (byte == 0xEF
                && bytes.get(index + 1) == Some(&0xBF)
                && matches!(bytes.get(index + 2), Some(0xBE | 0xBF)))
    })
}

/// The first `]]>` that starts in `bytes[start..end]`.
#[cold]
fn find_cdata_close(bytes: &[u8], start: usize, end: usize) -> Option<usize> {
    (start..end).find(|&index| bytes.get(index..index + 3) == Some(b"]]>".as_slice()))
}

/// The code point of the illegal character [`scan_characters`] found
/// at `at`, for its diagnostic.
pub(super) fn illegal_character_detail(bytes: &[u8], at: usize) -> String {
    let code = match bytes.get(at..at + 3) {
        Some([0xEF, 0xBF, 0xBE]) => 0xFFFE,
        Some([0xEF, 0xBF, 0xBF]) => 0xFFFF,
        _ => bytes.get(at).map_or(0, |byte| u32::from(*byte)),
    };
    format!("character U+{code:04X} is not allowed in XML")
}

const fn is_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\n' | b'\r')
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn names_follow_the_fifth_edition_classes() {
        for name in [
            b"a".as_slice(),
            b"_a",
            b"a-b.c9",
            "\u{e9}t\u{e9}".as_bytes(),
            "a\u{b7}".as_bytes(),
            "\u{10000}".as_bytes(),
        ] {
            assert!(is_ncname(name), "{name:?}");
        }
        for name in [
            b"".as_slice(),
            b"1a",
            b"-a",
            b".a",
            b"a:b",
            b"a b",
            b"a\"",
            "\u{d7}".as_bytes(),
            "\u{300}a".as_bytes(),
            "\u{b7}".as_bytes(),
        ] {
            assert!(!is_ncname(name), "{name:?}");
        }
        let qname = |name: &[u8]| {
            let (end, verdict) = scan_qname(name, 0, false);
            assert_eq!(end, name.len(), "{name:?}");
            verdict
        };
        assert_eq!(qname(b"w:t"), Ok(Some(1)));
        assert_eq!(qname(b"t"), Ok(None));
        assert_eq!(qname("a:\u{e9}".as_bytes()), Ok(Some(1)));
        for name in [b":a".as_slice(), b"a:", b"a:b:c", b"a:1b", b"", b"a\"b"] {
            assert!(qname(name).is_err(), "{name:?}");
        }
        assert_eq!(scan_qname(b"a:b=\"1\"", 0, true), (3, Ok(Some(1))));
        assert_eq!(scan_qname(b" x a", 3, false), (4, Ok(None)));
    }

    #[test]
    fn references_name_a_predefined_entity_or_a_legal_character() {
        for content in [
            b"lt".as_slice(),
            b"gt",
            b"amp",
            b"apos",
            b"quot",
            b"#9",
            b"#0065",
            b"#x20",
            b"#x10FFFF",
            b"#xfffd",
        ] {
            assert_eq!(check_reference(content), Ok(()), "{content:?}");
        }
        for content in [
            b"".as_slice(),
            b"a b",
            b"nbsp",
            b"AMP",
            b"#",
            b"#x",
            b"#X41",
            b"#xZZ",
            b"#0",
            b"#x1",
            b"#xD800",
            b"#xFFFE",
            b"#x110000",
            b"#99999999999999999999",
            b"#-1",
        ] {
            assert!(check_reference(content).is_err(), "{content:?}");
        }
    }

    #[test]
    fn attribute_values_are_normalized_as_xml_prescribes() {
        let mut out = Vec::new();
        push_normalized_value(b"a\tb\r\nc\rd\ne&amp;&#9;&#x41;", &mut out);
        assert_eq!(out, b"a b c d e&\tA");
    }

    /// Change 0750's differential campaign found a `]]>` split as `]]` at the
    /// end of one block and `>` at the start of the next, missed by the pass.
    /// Every placement of `]]>` near a block boundary and near the end of the
    /// input is found, at its own offset.
    #[test]
    fn a_cdata_close_is_found_wherever_it_straddles_a_block() {
        for length in [64_usize, 65, 66, 127, 128, 129, 130, 200] {
            for at in 0..length.saturating_sub(2) {
                let mut bytes = vec![b'a'; length];
                bytes[at..at + 3].copy_from_slice(b"]]>");
                assert_eq!(
                    scan_characters(&bytes).cdata_close,
                    at,
                    "length {length}, at {at}"
                );
                // A lone `]` or `]>` elsewhere does not hide it.
                if at >= 4 {
                    bytes[at - 4] = b']';
                    bytes[at - 3] = b'>';
                    assert_eq!(scan_characters(&bytes).cdata_close, at);
                }
            }
        }
    }

    #[test]
    fn one_pass_finds_the_first_illegal_character_and_the_first_cdata_close() {
        let mut bytes = vec![b'a'; 200];
        assert_eq!(
            scan_characters(&bytes),
            Characters {
                illegal: None,
                cdata_close: 200
            }
        );
        // Sequences that straddle a 64-byte block boundary.
        bytes[63] = 0xEF;
        bytes[64] = 0xBF;
        bytes[65] = 0xBE;
        bytes[127] = b']';
        bytes[128] = b']';
        bytes[129] = b'>';
        assert_eq!(
            scan_characters(&bytes),
            Characters {
                illegal: Some(63),
                cdata_close: 200
            }
        );
        assert_eq!(
            illegal_character_detail(&bytes, 63),
            "character U+FFFE is not allowed in XML"
        );
        bytes[63] = b'a';
        assert_eq!(
            scan_characters(&bytes),
            Characters {
                illegal: None,
                cdata_close: 127
            }
        );
        bytes[150] = 0x0B;
        assert_eq!(scan_characters(&bytes).illegal, Some(150));
        bytes[198] = b']';
        assert_eq!(next_cdata_close(&bytes, 128), 200);
        let legal = "\t\n\r \u{7f}\u{85}\u{fffd}\u{feff}\u{f000}\u{10ffff}";
        assert_eq!(scan_characters(legal.as_bytes()).illegal, None);
    }
}
