use aho_corasick::{AhoCorasick, MatchKind};
use once_cell::sync::Lazy;

// Static initialization: automaton is built only once, thread-safe
static XML_ESCAPER: Lazy<AhoCorasick> = Lazy::new(|| {
    AhoCorasick::builder()
        .build(["&", "<", ">", "\"", "'"])
        .expect("Failed to build XML escaper")
});

// Use LeftmostLongest to ensure longer entities are matched first (e.g., &amp; instead of &lt;)
static XML_UNESCAPER: Lazy<AhoCorasick> = Lazy::new(|| {
    AhoCorasick::builder()
        .match_kind(MatchKind::LeftmostLongest)
        .build(["&amp;", "&lt;", "&gt;", "&quot;", "&apos;"])
        .expect("Failed to build XML unescaper")
});

/// Escape XML special characters.
///
/// # Examples
///
/// ```
/// use litchi::common::xml::escape_xml;
/// assert_eq!(escape_xml("a & b"), "a &amp; b");
/// assert_eq!(escape_xml("<tag>\"hello\"</tag>"), "&lt;tag&gt;&quot;hello&quot;&lt;/tag&gt;");
/// ```
#[inline]
pub fn escape_xml(s: &str) -> String {
    XML_ESCAPER.replace_all(s, &["&amp;", "&lt;", "&gt;", "&quot;", "&apos;"])
}

/// Unescape XML special characters.
///
/// Replaces the five standard XML entities with their corresponding characters.
/// Unknown or malformed entities are left unchanged.
///
/// # Examples
///
/// ```
/// use litchi::common::xml::unescape_xml;
/// assert_eq!(unescape_xml("&lt;a &amp; b&gt;"), "<a & b>");
/// assert_eq!(unescape_xml("&quot;hello&apos;"), "\"hello'");
/// assert_eq!(unescape_xml("&amp;lt;"), "&lt;"); // &amp; is matched first
/// assert_eq!(unescape_xml("a & b"), "a & b"); // unchanged
/// assert_eq!(unescape_xml("&invalid;"), "&invalid;"); // unknown entity
/// assert_eq!(unescape_xml("&amp"), "&amp"); // incomplete, no semicolon
/// ```
#[inline]
pub fn unescape_xml(s: &str) -> String {
    XML_UNESCAPER.replace_all(s, &["&", "<", ">", "\"", "'"])
}
