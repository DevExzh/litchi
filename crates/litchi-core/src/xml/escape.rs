// The five escaped characters are all ASCII, so they are single identical
// bytes in UTF-8 and can never appear inside a multi-byte sequence. A plain
// left-to-right byte scan is therefore exactly equivalent to the previous
// Aho-Corasick automaton: every pattern is one byte, no two patterns can
// match at the same position, and replacement text is never rescanned.

/// Escape XML special characters.
///
/// # Examples
///
/// ```
/// use litchi_core::xml::escape_xml;
/// assert_eq!(escape_xml("a & b"), "a &amp; b");
/// assert_eq!(escape_xml("<tag>\"hello\"</tag>"), "&lt;tag&gt;&quot;hello&quot;&lt;/tag&gt;");
/// ```
#[inline]
#[allow(
    clippy::module_name_repetitions,
    reason = "public API name is stable and used by dependent crates; renaming would be a breaking change"
)]
pub fn escape_xml(s: &str) -> String {
    let bytes = s.as_bytes();
    let Some(mut special) = bytes
        .iter()
        .position(|byte| matches!(byte, b'&' | b'<' | b'>' | b'"' | b'\''))
    else {
        return s.to_string();
    };
    let mut out = String::with_capacity(s.len() + 16);
    let mut cursor = 0;
    loop {
        // `special` and `cursor` are char boundaries because every matched
        // byte is ASCII.
        out.push_str(&s[cursor..special]);
        out.push_str(match bytes[special] {
            b'&' => "&amp;",
            b'<' => "&lt;",
            b'>' => "&gt;",
            b'"' => "&quot;",
            _ => "&apos;",
        });
        cursor = special + 1;
        let Some(relative) = bytes[cursor..]
            .iter()
            .position(|byte| matches!(byte, b'&' | b'<' | b'>' | b'"' | b'\''))
        else {
            break;
        };
        special = cursor + relative;
    }
    out.push_str(&s[cursor..]);
    out
}

// The five entities share only the leading `&` and no entity is a prefix of
// another, so at any position at most one entity can match. Matching the
// single applicable entity at each `&` and never rescanning replacement text
// is therefore exactly equivalent to leftmost-longest automaton semantics.
const ENTITIES: [(&str, &str); 5] = [
    ("&amp;", "&"),
    ("&lt;", "<"),
    ("&gt;", ">"),
    ("&quot;", "\""),
    ("&apos;", "'"),
];

/// Unescape XML special characters.
///
/// Replaces the five standard XML entities with their corresponding characters.
/// Unknown or malformed entities are left unchanged.
///
/// # Examples
///
/// ```
/// use litchi_core::xml::unescape_xml;
/// assert_eq!(unescape_xml("&lt;a &amp; b&gt;"), "<a & b>");
/// assert_eq!(unescape_xml("&quot;hello&apos;"), "\"hello'");
/// assert_eq!(unescape_xml("&amp;lt;"), "&lt;"); // &amp; is matched first
/// assert_eq!(unescape_xml("a & b"), "a & b"); // unchanged
/// assert_eq!(unescape_xml("&invalid;"), "&invalid;"); // unknown entity
/// assert_eq!(unescape_xml("&amp"), "&amp"); // incomplete, no semicolon
/// ```
#[inline]
pub fn unescape_xml(s: &str) -> String {
    let bytes = s.as_bytes();
    let Some(first) = bytes.iter().position(|byte| *byte == b'&') else {
        return s.to_string();
    };
    let mut out = String::with_capacity(s.len());
    out.push_str(&s[..first]);
    let mut cursor = first;
    while cursor < bytes.len() {
        let rest = &s[cursor..];
        if let Some((entity, replacement)) =
            ENTITIES.iter().find(|(entity, _)| rest.starts_with(entity))
        {
            out.push_str(replacement);
            cursor += entity.len();
        } else {
            out.push('&');
            cursor += 1;
        }
        // `cursor` stays on a char boundary: it advances past either an
        // ASCII `&` or a complete ASCII entity.
        match bytes[cursor..].iter().position(|byte| *byte == b'&') {
            Some(relative) => {
                out.push_str(&s[cursor..cursor + relative]);
                cursor += relative;
            },
            None => {
                out.push_str(&s[cursor..]);
                break;
            },
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::{escape_xml, unescape_xml};

    #[test]
    fn escape_matches_every_special_byte_without_rescanning() {
        assert_eq!(escape_xml(""), "");
        assert_eq!(escape_xml("plain"), "plain");
        assert_eq!(escape_xml("&<>\"'"), "&amp;&lt;&gt;&quot;&apos;");
        assert_eq!(escape_xml("&&"), "&amp;&amp;");
        assert_eq!(
            escape_xml("caf\u{00E9} & \u{1F600}"),
            "caf\u{00E9} &amp; \u{1F600}"
        );
        // The ampersand in a literal "&lt;" is escaped, not preserved.
        assert_eq!(escape_xml("&lt;"), "&amp;lt;");
    }

    #[test]
    fn unescape_matches_leftmost_longest_without_rescanning() {
        assert_eq!(unescape_xml(""), "");
        assert_eq!(unescape_xml("plain"), "plain");
        // "&amp;lt;" decodes to the literal text "&lt;", never to "<".
        assert_eq!(unescape_xml("&amp;lt;"), "&lt;");
        assert_eq!(unescape_xml("&amp;amp;"), "&amp;");
        assert_eq!(unescape_xml("&amp;&amp;"), "&&");
        assert_eq!(unescape_xml("&"), "&");
        assert_eq!(unescape_xml("&&amp;"), "&&");
        assert_eq!(unescape_xml("&amp;&invalid;&lt;"), "&&invalid;<");
        assert_eq!(
            unescape_xml("caf\u{00E9}&lt;\u{1F600}"),
            "caf\u{00E9}<\u{1F600}"
        );
    }

    #[test]
    fn escape_and_unescape_round_trip() {
        for value in ["a & b", "<tag>\"x\"</tag>", "'", "&lt;already&gt;", "plain"] {
            assert_eq!(unescape_xml(&escape_xml(value)), value);
        }
    }
}
