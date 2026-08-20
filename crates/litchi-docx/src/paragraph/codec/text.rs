#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::option_option,
    reason = "nested options distinguish omitted, present-empty, and present-valued XML"
)]
#![expect(
    clippy::ref_option,
    reason = "the public API shape is retained for compatibility"
)]
//! Streaming text extraction for paragraph and run content.

use crate::binding_tracker::BindingTracker;
use crate::error::{Error, Result};
use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{QName, ResolveResult};
#[cfg(test)]
use quick_xml::reader::NsReader;
use quick_xml::reader::Reader;

use super::super::model::Paragraph;
use super::xml::is_fragment_word_name;

/// Maximum nesting depth accepted when extracting paragraph text.
const MAX_TEXT_SCAN_DEPTH: usize = 128;
/// Maximum number of elements scanned while extracting paragraph text.
const MAX_TEXT_SCAN_NODES: usize = 1_000_000;

pub(crate) fn extract_word_text(xml_bytes: &[u8]) -> Result<String> {
    // Plain reader + hand-rolled binding maintenance (change 0229, the
    // litchi-odt 0227 analog): the tracker replicates the push/pop
    // `NsReader` performs inside `read_resolved_event` (the
    // `BindingTracker` byte-exactness contract). Both the old slice-backed
    // `NsReader` and this plain `Reader` borrow their events; the removed work
    // is namespace maintenance, not an event-buffer copy.
    // `NsReader::from_reader` is `Reader::from_reader` with default
    // configuration, so the tokenization and error stream are unchanged.
    let mut reader = Reader::from_reader(xml_bytes);
    let mut tracker = BindingTracker::new();
    let mut pending_pop = false;
    let mut result = String::with_capacity(xml_bytes.len() / 8);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut text_depth = None;

    loop {
        // The deferred pop of the previous `End`/`Empty` scope runs before
        // the read, exactly where `NsReader::read_event_impl` applies it.
        if pending_pop {
            tracker.pop();
            pending_pop = false;
        }
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        // The push for a `Start`/`Empty` runs before the event is
        // classified, so a namespace error preempts the event exactly where
        // `read_resolved_event` returned `Err`. A push error is a real
        // `NamespaceError`, whose `Display` is what
        // `quick_xml::Error::Namespace` forwards to, so the `Error::Xml`
        // message is byte-identical to the historical failure.
        //
        // `resolve_event` maps `Start`/`Empty`/`End` to
        // `resolve(name, use_default = true)` and every other event to
        // `Unbound`; the `End` name resolves in its own scope because the
        // pop is deferred to the next read. This path consumes the `End`
        // resolution (the `text_depth` closing match below), unlike the
        // litchi-odt text path.
        let namespace = match &event {
            Event::Start(element) => {
                tracker
                    .push(element)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                tracker.resolve_element(element.name()).0
            },
            Event::Empty(element) => {
                tracker
                    .push(element)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                // The scope an `Empty` element opens closes immediately:
                // defer its pop to the top of the next iteration.
                pending_pop = true;
                tracker.resolve_element(element.name()).0
            },
            Event::End(element) => {
                pending_pop = true;
                tracker.resolve_element(element.name()).0
            },
            _ => ResolveResult::Unbound,
        };

        if fragment_prefix.is_none()
            && let Event::Start(element) = &event
            && !matches!(namespace, ResolveResult::Bound(_))
        {
            fragment_prefix = Some(
                element
                    .name()
                    .prefix()
                    .map(|prefix| prefix.into_inner().to_vec()),
            );
        }

        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML element counter overflow".to_string())
                })?;
                if nodes > MAX_TEXT_SCAN_NODES {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                    )));
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML nesting is too deep".to_string())
                })?;
                if depth > MAX_TEXT_SCAN_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML nesting exceeds the {MAX_TEXT_SCAN_DEPTH} depth limit"
                    )));
                }
                if text_depth.is_none()
                    && is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = Some(depth);
                } else if let Some(character) =
                    word_special_character(&namespace, element.name(), &fragment_prefix)
                {
                    result.push(character);
                }
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML element counter overflow".to_string())
                })?;
                if nodes > MAX_TEXT_SCAN_NODES {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                    )));
                }
                if let Some(character) =
                    word_special_character(&namespace, element.name(), &fragment_prefix)
                {
                    result.push(character);
                }
            },
            Event::Text(text) if text_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                let unescaped = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                result.push_str(&unescaped);
            },
            Event::CData(text) if text_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                result.push_str(&decoded);
            },
            Event::GeneralRef(reference) if text_depth.is_some() => {
                result.push_str(&decode_xml_reference(&reference)?);
            },
            Event::End(element) => {
                if text_depth == Some(depth)
                    && is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid Word XML nesting".to_string()))?;
            },
            Event::Eof if depth != 0 || text_depth.is_some() => {
                return Err(Error::InvalidFormat("unterminated Word XML".to_string()));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    result.shrink_to_fit();
    Ok(result)
}

fn word_special_character(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Option<char> {
    if is_fragment_word_name(namespace, name, b"tab", fragment_prefix) {
        Some('\t')
    } else if is_fragment_word_name(namespace, name, b"br", fragment_prefix)
        || is_fragment_word_name(namespace, name, b"cr", fragment_prefix)
    {
        Some('\n')
    } else if is_fragment_word_name(namespace, name, b"noBreakHyphen", fragment_prefix) {
        Some('\u{2011}')
    } else if is_fragment_word_name(namespace, name, b"softHyphen", fragment_prefix) {
        Some('\u{00ad}')
    } else {
        None
    }
}

impl Paragraph {
    /// Get the text content of this paragraph.
    ///
    /// Concatenates all text from all runs in the paragraph.
    ///
    /// # Performance
    ///
    /// Uses streaming XML parsing with pre-allocated buffer to extract text efficiently.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn text(&self) -> Result<String> {
        extract_word_text(self.xml_bytes())
    }
}

/// Change-0229 differential oracle: the pre-0229 `NsReader` implementation
/// of [`extract_word_text`], retained test-only so the tracker-driven path
/// is pinned byte-for-byte against it (the litchi-odt 0227 pattern, where
/// the oracle stayed in production because other APIs used it; here nothing
/// else needs it).
#[cfg(test)]
fn extract_word_text_nsreader_oracle(xml_bytes: &[u8]) -> Result<String> {
    let mut reader = NsReader::from_reader(xml_bytes);
    let mut result = String::with_capacity(xml_bytes.len() / 8);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut text_depth = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;

        if fragment_prefix.is_none()
            && let Event::Start(element) = &event
            && !matches!(namespace, ResolveResult::Bound(_))
        {
            fragment_prefix = Some(
                element
                    .name()
                    .prefix()
                    .map(|prefix| prefix.into_inner().to_vec()),
            );
        }

        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML element counter overflow".to_string())
                })?;
                if nodes > MAX_TEXT_SCAN_NODES {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                    )));
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML nesting is too deep".to_string())
                })?;
                if depth > MAX_TEXT_SCAN_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML nesting exceeds the {MAX_TEXT_SCAN_DEPTH} depth limit"
                    )));
                }
                if text_depth.is_none()
                    && is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = Some(depth);
                } else if let Some(character) =
                    word_special_character(&namespace, element.name(), &fragment_prefix)
                {
                    result.push(character);
                }
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML element counter overflow".to_string())
                })?;
                if nodes > MAX_TEXT_SCAN_NODES {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                    )));
                }
                if let Some(character) =
                    word_special_character(&namespace, element.name(), &fragment_prefix)
                {
                    result.push(character);
                }
            },
            Event::Text(text) if text_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                let unescaped = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                result.push_str(&unescaped);
            },
            Event::CData(text) if text_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                result.push_str(&decoded);
            },
            Event::GeneralRef(reference) if text_depth.is_some() => {
                result.push_str(&decode_xml_reference(&reference)?);
            },
            Event::End(element) => {
                if text_depth == Some(depth)
                    && is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid Word XML nesting".to_string()))?;
            },
            Event::Eof if depth != 0 || text_depth.is_some() => {
                return Err(Error::InvalidFormat("unterminated Word XML".to_string()));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    result.shrink_to_fit();
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::{Path, PathBuf};

    /// Transitional (loose) and strict WordprocessingML namespace URIs.
    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const W_STRICT: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

    /// Differential parity: the tracker-driven path and the `NsReader`
    /// oracle must agree on the extracted text or fail with a
    /// byte-identical error string.
    fn assert_extract_parity(xml: &[u8]) {
        let tracker = extract_word_text(xml);
        let oracle = extract_word_text_nsreader_oracle(xml);
        match (tracker, oracle) {
            (Ok(tracker), Ok(oracle)) => {
                assert_eq!(tracker, oracle, "tracker/oracle extracted text diverges")
            },
            (Err(tracker), Err(oracle)) => assert_eq!(
                tracker.to_string(),
                oracle.to_string(),
                "tracker/oracle error strings diverge"
            ),
            (tracker, oracle) => {
                panic!("tracker/oracle outcome mismatch: {tracker:?} vs {oracle:?}")
            },
        }
    }

    #[test]
    fn tracker_path_matches_oracle_on_core_extraction() {
        let fixtures: Vec<(String, &str)> = vec![
            // Plain run text.
            (
                format!(r#"<w:p xmlns:w="{W}"><w:r><w:t>Hello</w:t></w:r></w:p>"#),
                "Hello",
            ),
            // All five special characters, as Empty and Start/End elements.
            (
                format!(
                    r#"<w:p xmlns:w="{W}"><w:r><w:tab/><w:br/><w:cr></w:cr><w:noBreakHyphen/><w:softHyphen/></w:r></w:p>"#
                ),
                "\t\n\n\u{2011}\u{00ad}",
            ),
            // Strict-namespace matching: strict `w:t` is extracted.
            (
                format!(r#"<w:p xmlns:w="{W_STRICT}"><w:r><w:t>S</w:t></w:r></w:p>"#),
                "S",
            ),
            // Bare fragment fallback: `w` never declared, the first unbound
            // Start fixes the fragment prefix and `w:t` matches through it.
            (r#"<w:p><w:r><w:t>Hi</w:t></w:r></w:p>"#.to_string(), "Hi"),
            // Unprefixed bare fragment: the first unbound Start has no
            // prefix, so unprefixed `t` matches the fallback.
            (r#"<p><r><t>Hi</t></r></p>"#.to_string(), "Hi"),
            // Prefix shadowing at depth: the rebound scope's `w:t` resolves
            // to the foreign URI and is skipped; the outer binding resumes.
            (
                format!(
                    r#"<w:p xmlns:w="{W}"><w:r><w:t>A</w:t></w:r><w:r xmlns:w="urn:foreign"><w:t>B</w:t></w:r><w:r><w:t>C</w:t></w:r></w:p>"#
                ),
                "AC",
            ),
            // Special characters under the shadowed prefix are skipped too.
            (
                format!(
                    r#"<w:p xmlns:w="{W}"><w:r><w:t>A</w:t><w:tab/></w:r><w:r xmlns:w="urn:foreign"><w:tab/></w:r><w:r><w:t>C</w:t></w:r></w:p>"#
                ),
                "A\tC",
            ),
        ];
        for (xml, expected) in &fixtures {
            assert_extract_parity(xml.as_bytes());
            assert_eq!(
                extract_word_text(xml.as_bytes()).unwrap(),
                *expected,
                "unexpected extracted text for {xml}"
            );
        }
    }

    #[test]
    fn tracker_path_matches_oracle_on_subtle_namespace_fallbacks() {
        // These cases exercise interactions between the fragment-prefix
        // fallback and binding scopes whose pinned outcome is subtle (an
        // emptied default binding captures the fallback; an emptied prefix
        // binding matches a previously captured prefix); parity with the
        // oracle is the assertion, no hardcoded text.
        let fixtures: Vec<String> = vec![
            // Default-namespace redefinition and unset.
            format!(r#"<p xmlns="{W}"><t>A</t><d xmlns=""><t>B</t></d><t>C</t></p>"#),
            // Emptied prefix binding (`xmlns:w=""`) after the prefix
            // resolved properly: the emptied scope's names resolve
            // `Unknown("w")`, which the fragment fallback may still match.
            format!(
                r#"<w:p xmlns:w="{W}"><w:r><w:t>A</w:t></w:r><w:r xmlns:w=""><w:t>B</w:t></w:r></w:p>"#
            ),
            // An unbound root with a properly bound block inside: the bound
            // `w:t` wins through the namespace match, not the fallback.
            format!(r#"<p xmlns:w="{W}"><w:r><w:t>in</w:t></w:r><r><t>out</t></r></p>"#),
            // An `Empty` element carrying a declaration: its push and
            // deferred pop bracket the event itself.
            format!(
                r#"<w:p xmlns:w="{W}"><w:marker xmlns:w="urn:foreign"/><w:r><w:t>After</w:t></w:r></w:p>"#
            ),
            // An attribute *value* containing the substring `xmlns` must not
            // disturb the binding scan outcome.
            format!(r#"<w:p xmlns:w="{W}"><w:r w:id="xmlns:shadow"><w:t>A</w:t></w:r></w:p>"#),
            // CDATA and a general entity reference inside and outside `w:t`.
            format!(
                r#"<w:p xmlns:w="{W}"><w:r><w:t>x<![CDATA[<raw>]]>&amp;y</w:t></w:r><w:instrText><![CDATA[skip]]></w:instrText></w:p>"#
            ),
            // Mixed loose root and strict inner binding.
            format!(
                r#"<w:p xmlns:w="{W}"><w:r><w:t>A</w:t></w:r><w:s xmlns:w="{W_STRICT}"><w:t>B</w:t></w:s></w:p>"#
            ),
        ];
        for xml in &fixtures {
            assert_extract_parity(xml.as_bytes());
        }
    }

    #[test]
    fn tracker_path_matches_oracle_on_namespace_errors() {
        let xml_ns = "http://www.w3.org/XML/1998/namespace";
        let xmlns_ns = "http://www.w3.org/2000/xmlns/";
        let error_fixtures: Vec<String> = vec![
            // Declaring the `xmlns` prefix itself.
            format!(r#"<w:p xmlns:w="{W}" xmlns:xmlns="urn:example:x"><w:t>x</w:t></w:p>"#),
            // Binding `xml` to a foreign URI.
            format!(r#"<w:p xmlns:w="{W}" xmlns:xml="urn:example:x"><w:t>x</w:t></w:p>"#),
            // Binding another prefix to the reserved xml URI.
            format!(r#"<w:p xmlns:w="{W}" xmlns:q="{xml_ns}"><w:t>x</w:t></w:p>"#),
            // Binding a prefix to the reserved xmlns URI.
            format!(r#"<w:p xmlns:w="{W}" xmlns:q="{xmlns_ns}"><w:t>x</w:t></w:p>"#),
            // The same failure mid-stream after extractable text: the push
            // error preempts the event exactly where the `NsReader` read
            // error did.
            format!(
                r#"<w:p xmlns:w="{W}"><w:r><w:t>A</w:t></w:r><w:r xmlns:xml="urn:example:x"><w:t>B</w:t></w:r></w:p>"#
            ),
            // A namespace error on an `Empty` element.
            format!(
                r#"<w:p xmlns:w="{W}"><w:r><w:t>A</w:t></w:r><w:tab xmlns:xmlns="urn:x"/></w:p>"#
            ),
        ];
        for xml in &error_fixtures {
            assert_extract_parity(xml.as_bytes());
            assert!(
                extract_word_text(xml.as_bytes()).is_err(),
                "expected an error for {xml}"
            );
        }
        // Rebinding `xml` to its reserved URI is a no-op, not an error.
        let benign = format!(r#"<w:p xmlns:w="{W}" xmlns:xml="{xml_ns}"><w:t>x</w:t></w:p>"#);
        assert_extract_parity(benign.as_bytes());
        assert_eq!(extract_word_text(benign.as_bytes()).unwrap(), "x");

        // Declaration-limit parity (`xmlns:w` accounts for one declaration
        // on the tag): 256 declarations pass, 257 fail identically.
        let declarations = |count: usize| {
            (0..count)
                .map(|index| format!(r#"xmlns:d{index}="urn:example:{index}""#))
                .collect::<Vec<_>>()
                .join(" ")
        };
        let within_limit = format!(
            r#"<w:p xmlns:w="{W}" {}><w:t>x</w:t></w:p>"#,
            declarations(255)
        );
        assert_extract_parity(within_limit.as_bytes());
        assert_eq!(extract_word_text(within_limit.as_bytes()).unwrap(), "x");
        let over_limit = format!(
            r#"<w:p xmlns:w="{W}" {}><w:t>x</w:t></w:p>"#,
            declarations(256)
        );
        assert_extract_parity(over_limit.as_bytes());
        assert!(extract_word_text(over_limit.as_bytes()).is_err());
    }

    #[test]
    fn tracker_path_matches_oracle_on_malformed_and_limited_xml() {
        // Unterminated document: the Eof structural check fires identically.
        let unterminated = format!(r#"<w:p xmlns:w="{W}"><w:r><w:t>unfinished"#);
        assert_extract_parity(unterminated.as_bytes());
        // Tokenizer error mid-tag.
        let broken = format!(r#"<w:p xmlns:w="{W}"><w:t a=">#</w:p>"#);
        assert_extract_parity(broken.as_bytes());
        // Depth limit: 129 nested elements exceed MAX_TEXT_SCAN_DEPTH = 128.
        let deep = format!(
            r#"<w:p xmlns:w="{W}">{}x{}</w:p>"#,
            "<w:a>".repeat(MAX_TEXT_SCAN_DEPTH + 1),
            "</w:a>".repeat(MAX_TEXT_SCAN_DEPTH + 1),
        );
        assert_extract_parity(deep.as_bytes());
        let error = extract_word_text(deep.as_bytes()).unwrap_err();
        assert_eq!(
            error.to_string(),
            extract_word_text_nsreader_oracle(deep.as_bytes())
                .unwrap_err()
                .to_string()
        );
        assert!(error.to_string().contains("depth limit"));
    }

    #[test]
    fn tracker_path_matches_oracle_on_docx_corpus() {
        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
        let mut files = Vec::new();
        collect_docx_corpus(&root, &mut files);
        files.sort();
        assert!(!files.is_empty(), "no .docx corpus fixtures discovered");
        let mut compared = 0usize;
        for path in &files {
            let Some(document_xml) = docx_document_xml(path) else {
                continue;
            };
            assert_extract_parity(&document_xml);
            compared += 1;
        }
        assert!(
            compared > 0,
            "no .docx corpus fixtures yielded word/document.xml"
        );
        eprintln!("corpus parity compared over {compared} document parts");
    }

    fn collect_docx_corpus(directory: &Path, files: &mut Vec<PathBuf>) {
        let Ok(entries) = std::fs::read_dir(directory) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                collect_docx_corpus(&path, files);
            } else if path
                .extension()
                .is_some_and(|extension| extension == "docx")
            {
                files.push(path);
            }
        }
    }

    fn docx_document_xml(path: &Path) -> Option<Vec<u8>> {
        let bytes = std::fs::read(path).ok()?;
        let reader = soapberry_zip::office::ArchiveReader::new(&bytes).ok()?;
        reader.read("word/document.xml").ok()
    }
}
