//! The admission proof against the processor it stands in for.
//!
//! The witness matrix pins each verdict 0008 and 0011 name. The generated
//! differential checks the property every caller relies on: whenever the
//! proof admits a part, processing that part succeeds, and the projection
//! carries exactly the part's elements, attributes and text, less the
//! document element's `Ignorable` and whatever the verdict says is dropped.

use std::borrow::Cow;

use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::Reader;

use super::{
    Admission, Held, Holder, Refusal, admit, admit_held, admit_with, binds_namespace_by_reference,
    mentions_namespace,
};
use crate::mce::{
    Capabilities, Limits, NAMESPACE, OffsetLimits, active_offsets, process_markup_compatibility,
};
use crate::xml::attributes::BytesStartExt as _;

const MC: &str = NAMESPACE;
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const X: &str = "urn:x";

/// The namespace with its last character written as a character reference,
/// the 0008 W7 spelling.
const MC_ESCAPED: &str = "http://schemas.openxmlformats.org/markup-compatibility/200&#54;";

/// A profile like the Word main document's: the OOXML baseline and `w14`.
fn word_profile() -> Capabilities {
    let mut capabilities = Capabilities::default();
    capabilities.understand_namespace(W14);
    capabilities
}

fn verdict(xml: &str) -> Result<Admission, Refusal> {
    admit(xml.as_bytes(), &word_profile(), &Limits::default())
}

fn document(root_attributes: &str, body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}" xmlns:w14="{W14}" xmlns:r="{R}" xmlns:x="{X}" {root_attributes}><w:body>{body}</w:body></w:document>"#
    )
}

const PARAGRAPH: &str = "<w:p><w:r><w:t>A</w:t></w:r></w:p>";

/// One node of a part as a namespace-aware reader sees it, names decoded.
#[derive(Debug, PartialEq, Eq)]
enum Node {
    Start(String, String, Vec<(String, String, String)>),
    End,
    Text(String),
    Comment(String),
}

/// The nodes of a part's document element, with declarations resolved on
/// their decoded values. Attributes in the markup-compatibility namespace
/// are left out, and so are elements (with their subtrees) and attributes
/// in any namespace of `hidden`.
fn nodes(xml: &[u8], hidden: &[&str]) -> Vec<Node> {
    let mut collector = Collector {
        hidden,
        scopes: Vec::new(),
        skipping: None,
        output: Vec::new(),
    };
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().unwrap();
        let visible = !collector.scopes.is_empty() && collector.skipping.is_none();
        match event {
            Event::Start(element) => collector.open(&element, decoder, false),
            Event::Empty(element) => collector.open(&element, decoder, true),
            Event::End(_) => collector.close(),
            Event::Text(text) if visible => collector.text(&text.decode().unwrap()),
            Event::CData(text) if visible => collector.text(&text.decode().unwrap()),
            Event::GeneralRef(reference) if visible => {
                let text = match reference.resolve_char_ref().unwrap() {
                    Some(character) => character.to_string(),
                    None => {
                        let name = reference.decode().unwrap();
                        quick_xml::escape::unescape(&format!("&{name};"))
                            .unwrap()
                            .into_owned()
                    },
                };
                collector.text(&text);
            },
            Event::Comment(comment) if visible => {
                let comment = comment.decode().unwrap().into_owned();
                collector.output.push(Node::Comment(comment));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    collector.output
}

struct Collector<'a> {
    hidden: &'a [&'a str],
    scopes: Vec<Vec<(String, String)>>,
    /// The depth of the hidden element being skipped.
    skipping: Option<usize>,
    output: Vec<Node>,
}

impl Collector<'_> {
    fn resolve(&self, prefix: &str, element: bool) -> String {
        if prefix == "xml" {
            return "http://www.w3.org/XML/1998/namespace".into();
        }
        if prefix.is_empty() && !element {
            return String::new();
        }
        self.scopes
            .iter()
            .rev()
            .flat_map(|scope| scope.iter().rev())
            .find(|(candidate, _)| candidate == prefix)
            .map_or_else(String::new, |(_, uri)| uri.clone())
    }

    fn open(
        &mut self,
        element: &quick_xml::events::BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
        empty: bool,
    ) {
        let mut declarations = Vec::new();
        let mut attributes = Vec::new();
        for attribute in element.checked_attributes() {
            let attribute = attribute.unwrap();
            let key = std::str::from_utf8(attribute.key.into_inner())
                .unwrap()
                .to_owned();
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .unwrap()
                .into_owned();
            if key == "xmlns" {
                declarations.push((String::new(), value));
            } else if let Some(prefix) = key.strip_prefix("xmlns:") {
                declarations.push((prefix.to_owned(), value));
            } else {
                attributes.push((key, value));
            }
        }
        self.scopes.push(declarations);
        let name = std::str::from_utf8(element.name().into_inner())
            .unwrap()
            .to_owned();
        let (prefix, local) = name.split_once(':').unwrap_or(("", &name));
        let namespace = self.resolve(prefix, true);
        if self.skipping.is_none() && self.hidden.contains(&namespace.as_str()) {
            self.skipping = Some(self.scopes.len());
        }
        if self.skipping.is_none() {
            let mut resolved: Vec<_> = attributes
                .into_iter()
                .map(|(key, value)| {
                    let (prefix, local) = key.split_once(':').unwrap_or(("", &key));
                    (self.resolve(prefix, false), local.to_owned(), value)
                })
                .filter(|(namespace, _, _)| {
                    namespace != MC && !self.hidden.contains(&namespace.as_str())
                })
                .collect();
            resolved.sort();
            self.output
                .push(Node::Start(namespace, local.to_owned(), resolved));
        }
        if empty {
            self.close();
        }
    }

    fn close(&mut self) {
        if self.skipping.is_none() {
            self.output.push(Node::End);
        }
        if self.skipping == Some(self.scopes.len()) {
            self.skipping = None;
        }
        self.scopes.pop();
    }

    fn text(&mut self, text: &str) {
        if let Some(Node::Text(previous)) = self.output.last_mut() {
            previous.push_str(text);
        } else {
            self.output.push(Node::Text(text.to_owned()));
        }
    }
}

/// Assert what an admission promises about processing `xml` with
/// `capabilities`; `dropped` names the namespaces the verdict drops.
#[track_caller]
fn assert_projection(
    xml: &[u8],
    capabilities: &Capabilities,
    admission: Admission,
    dropped: &[&str],
) {
    let output = process_markup_compatibility(xml, capabilities, &Limits::default())
        .unwrap_or_else(|error| panic!("an admitted part must process: {error}"));
    match (admission, &output.xml) {
        (Admission::Unbound, Cow::Borrowed(_)) => return,
        (Admission::Unbound, Cow::Owned(_)) => panic!("an unbound part must be borrowed"),
        (_, Cow::Borrowed(_)) => panic!("a bound part is rewritten"),
        (_, Cow::Owned(_)) => {},
    }
    let report = &output.report;
    assert_eq!(report.alternate_content_count, 0);
    assert_eq!(report.unwrapped_elements, 0);
    assert_eq!(report.preserved_elements + report.preserved_attributes, 0);
    // Every directive the part carries is one dropped attribute.
    let directive = std::str::from_utf8(xml).map_or(0, |text| text.matches(":Ignorable=").count());
    if admission.is_lexical() {
        assert_eq!(
            report.ignored_elements, 0,
            "a lexical verdict drops no element"
        );
        assert_eq!(
            report.ignored_attributes, directive,
            "a lexical verdict drops only the directives"
        );
        assert_eq!(
            nodes(xml, &[]),
            nodes(&output.xml, &[]),
            "a lexical verdict changes nothing but spelling"
        );
    } else {
        assert!(report.ignored_elements + report.ignored_attributes > directive);
        assert_eq!(nodes(xml, dropped), nodes(&output.xml, &[]));
    }
}

#[test]
fn mentions_namespace_decides_on_decoded_declarations() {
    let tail = MC.strip_prefix('h').unwrap();
    let all_references: String = MC
        .chars()
        .map(|character| format!("&#{};", u32::from(character)))
        .collect();
    for (xml, expected) in [
        // The literal namespace anywhere: the historical fast-path test.
        (format!(r#"<r xmlns:mc="{MC}"/>"#), true),
        (format!(r#"<r xmlns="{MC}"/>"#), true),
        (format!(r#"<r>{MC}</r>"#), true),
        (format!(r#"<r v="{MC}"/>"#), true),
        (format!(r#"<r xmlns:mc=" {MC}"/>"#), true),
        // Declarations that spell it with references.
        (format!(r#"<r xmlns:mc="{MC_ESCAPED}"/>"#), true),
        (format!(r#"<r xmlns:mc='{MC_ESCAPED}'/>"#), true),
        (format!("<r xmlns:mc \n=\t \"{MC_ESCAPED}\"/>"), true),
        (format!(r#"<r xmlns="{MC_ESCAPED}"/>"#), true),
        (format!(r#"<r xmlns:mc="&#x68;{tail}"/>"#), true),
        (format!(r#"<r xmlns:mc="&#000000000104;{tail}"/>"#), true),
        (format!(r#"<r xmlns:mc="{all_references}"/>"#), true),
        // Reads like a declaration: a slower path, never a missed binding.
        (format!(r#"<r><!-- xmlns:mc="{MC_ESCAPED}" --></r>"#), true),
        (format!(r#"<r a:xmlns="{MC_ESCAPED}"/>"#), true),
        // References that decode to something else, or to nothing.
        (format!(r#"<r xmlnsX="{MC_ESCAPED}"/>"#), false),
        (format!(r#"<r xmlns:mc="{MC_ESCAPED}/"/>"#), false),
        (format!(r#"<r xmlns:mc="&#32;{tail}"/>"#), false),
        (format!(r#"<r xmlns:mc="&#x68;{tail}&bogus;"/>"#), false),
        (format!(r#"<r xmlns:mc="&#x68;{tail}"#), false),
        (format!(r#"<r v="{MC_ESCAPED}"/>"#), false),
        (r#"<r xmlns:mc="urn:a&amp;b"/>"#.to_owned(), false),
        (r#"<r xmlns:mc="urn:other"/>"#.to_owned(), false),
        (String::new(), false),
    ] {
        assert_eq!(mentions_namespace(xml.as_bytes()), expected, "{xml}");
        // Only a binding spelled with references hides from the literal test.
        let literal = xml.contains(MC);
        assert_eq!(
            binds_namespace_by_reference(xml.as_bytes()),
            expected && !literal,
            "{xml}"
        );
    }
}

#[test]
fn processing_applies_an_escaped_namespace_and_skips_a_mention() {
    // 0008 W7: the escaped spelling binds the namespace, so the branch is
    // selected for every reader instead of both branches being read.
    let escaped = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="{MC_ESCAPED}"><w:body><mc:AlternateContent><mc:Choice Requires="w"><w:p><w:r><w:t>CHOICE</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>FALLBACK</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:body></w:document>"#
    );
    let output =
        process_markup_compatibility(escaped.as_bytes(), &word_profile(), &Limits::default())
            .unwrap();
    let text = std::str::from_utf8(&output.xml).unwrap();
    assert_eq!(output.report.alternate_content_count, 1);
    assert!(
        text.contains("CHOICE") && !text.contains("FALLBACK"),
        "{text}"
    );
    assert!(!text.contains("AlternateContent"), "{text}");
    assert_eq!(verdict(&escaped), Err(Refusal::Element));

    // Active offsets agree: the fallback paragraph is not active.
    let fallback = u32::try_from(escaped.find("<w:p><w:r><w:t>FALLBACK").unwrap()).unwrap();
    let choice = u32::try_from(escaped.find("<w:p><w:r><w:t>CHOICE").unwrap()).unwrap();
    let active = active_offsets(
        escaped.as_bytes(),
        &[choice, fallback],
        &word_profile(),
        &OffsetLimits::default(),
    )
    .unwrap();
    assert_eq!(active, vec![choice]);

    // A part that only mentions the namespace binds nothing: processing
    // respells it and the proof admits it as declared-only.
    let mention = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>{MC}</w:t></w:r></w:p></w:body></w:document>"#
    );
    assert_eq!(verdict(&mention), Ok(Admission::DeclaredOnly));
    assert_projection(
        mention.as_bytes(),
        &word_profile(),
        Admission::DeclaredOnly,
        &[],
    );
}

#[test]
fn witness_matrix() {
    let mc = format!(r#"xmlns:mc="{MC}""#);
    let escaped = format!(r#"xmlns:mc="{MC_ESCAPED}""#);
    let cases: Vec<(&str, String, Result<Admission, Refusal>)> = vec![
        (
            "no namespace",
            document("", PARAGRAPH),
            Ok(Admission::Unbound),
        ),
        (
            "declared only (Word 2007 `ve`)",
            document(&format!(r#"xmlns:ve="{MC}""#), PARAGRAPH),
            Ok(Admission::DeclaredOnly),
        ),
        (
            "declared only below the root",
            document(
                "",
                &format!(r#"<w:p xmlns:mc="{MC}"><w:r><w:t>A</w:t></w:r></w:p>"#),
            ),
            Ok(Admission::DeclaredOnly),
        ),
        (
            "declared only, escaped (W7 without a construct)",
            document(&escaped, PARAGRAPH),
            Ok(Admission::DeclaredOnly),
        ),
        (
            "root Ignorable, understood and unused prefixes",
            document(
                &format!(r#"{mc} xmlns:w15="urn:w15" mc:Ignorable="w14 w15""#),
                r#"<w:p w14:paraId="1A2B3C4D"><w:r><w:t>A</w:t></w:r></w:p>"#,
            ),
            Ok(Admission::Ignorable),
        ),
        (
            "root Ignorable through an escaped binding",
            document(&format!(r#"{escaped} mc:Ignorable="w14""#), PARAGRAPH),
            Ok(Admission::Ignorable),
        ),
        (
            "W6: ignorable attribute and leaf element",
            document(
                &format!(r#"{mc} mc:Ignorable="x""#),
                r#"<w:p x:id="1"><w:r><x:flag/><w:t>A</w:t></w:r></w:p>"#,
            ),
            Ok(Admission::IgnorableOpaque),
        ),
        (
            "wp14-style leaves with plain attributes and opaque text",
            document(
                &format!(r#"{mc} mc:Ignorable="x""#),
                r#"<w:p><w:r><x:size relativeFrom="page" xml:space="preserve"><x:pct>0</x:pct></x:size><w:t>A</w:t></w:r></w:p>"#,
            ),
            Ok(Admission::IgnorableOpaque),
        ),
        (
            "used ignorable prefix holding a run (W4)",
            document(
                &format!(r#"{mc} mc:Ignorable="x""#),
                r#"<w:p><x:ext><w:r><w:t>HIDDEN</w:t></w:r></x:ext><w:r><w:t>B</w:t></w:r></w:p>"#,
            ),
            Err(Refusal::HiddenContent),
        ),
        (
            "used ignorable prefix holding a relationship reference",
            document(
                &format!(r#"{mc} mc:Ignorable="x""#),
                r#"<w:p><x:ext r:id="rId1"/><w:r><w:t>B</w:t></w:r></w:p>"#,
            ),
            Err(Refusal::HiddenContent),
        ),
        (
            "document element in an ignorable namespace",
            format!(r#"<x:root xmlns:x="{X}" {mc} mc:Ignorable="x"/>"#),
            Err(Refusal::HiddenContent),
        ),
        (
            "AlternateContent in a run (W1)",
            document(
                &mc,
                r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:t>BOX</w:t></mc:Choice><mc:Fallback><w:t>BOX</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>"#,
            ),
            Err(Refusal::Element),
        ),
        (
            "AlternateContent in the body (W2)",
            document(
                &mc,
                r#"<w:p/><mc:AlternateContent><mc:Choice Requires="w14"><w:p/></mc:Choice><mc:Fallback/></mc:AlternateContent><w:p/>"#,
            ),
            Err(Refusal::Element),
        ),
        (
            "AlternateContent through an escaped binding (W7)",
            document(
                &escaped,
                r#"<mc:AlternateContent><mc:Choice Requires="w14"><w:p/></mc:Choice></mc:AlternateContent>"#,
            ),
            Err(Refusal::Element),
        ),
        (
            "ProcessContent (W3)",
            document(
                &format!(r#"{mc} mc:Ignorable="x" mc:ProcessContent="x:wrap""#),
                r#"<w:p><x:wrap><w:r><w:t>W</w:t></w:r></x:wrap></w:p>"#,
            ),
            Err(Refusal::Directive),
        ),
        (
            "MustUnderstand of an understood namespace",
            document(&format!(r#"{mc} mc:MustUnderstand="w14""#), PARAGRAPH),
            Err(Refusal::Directive),
        ),
        (
            "MustUnderstand of an unknown namespace (W5)",
            document(&format!(r#"{mc} mc:MustUnderstand="x""#), PARAGRAPH),
            Err(Refusal::Directive),
        ),
        (
            "PreserveElements",
            document(
                &format!(r#"{mc} mc:Ignorable="x" mc:PreserveElements="x:keep""#),
                PARAGRAPH,
            ),
            Err(Refusal::Directive),
        ),
        (
            "PreserveAttributes",
            document(
                &format!(r#"{mc} mc:Ignorable="x" mc:PreserveAttributes="x:*""#),
                PARAGRAPH,
            ),
            Err(Refusal::Directive),
        ),
        (
            "Ignorable below the document element, not understood",
            document(
                &mc,
                r#"<w:p mc:Ignorable="x"><w:r><w:t>A</w:t></w:r></w:p>"#,
            ),
            Err(Refusal::NestedIgnorable),
        ),
        (
            "Ignorable below the document element, understood (revision writer)",
            document(
                "",
                &format!(
                    r#"<w:p><w:ins w:id="1" w:author="A" xmlns:mc="{MC}" mc:Ignorable="w14" w14:dateUtc="1"><w:r><w:t>A</w:t></w:r></w:ins></w:p>"#
                ),
            ),
            Ok(Admission::Ignorable),
        ),
        (
            "Ignorable below the document element, listing the namespace",
            document(&mc, r#"<w:p mc:Ignorable="mc"/>"#),
            Err(Refusal::Unproven),
        ),
        (
            "unknown markup-compatibility attribute",
            document(&format!(r#"{mc} mc:Bogus="1""#), PARAGRAPH),
            Err(Refusal::Unproven),
        ),
        (
            "Ignorable listing an unbound prefix",
            document(&format!(r#"{mc} mc:Ignorable="u""#), PARAGRAPH),
            Err(Refusal::Unproven),
        ),
        (
            "Ignorable listing a prefix twice",
            document(&format!(r#"{mc} mc:Ignorable="x x""#), PARAGRAPH),
            Err(Refusal::Unproven),
        ),
        (
            "Ignorable listing the namespace itself",
            document(&format!(r#"{mc} mc:Ignorable="mc""#), PARAGRAPH),
            Err(Refusal::Unproven),
        ),
        (
            "Ignorable twice through two prefixes",
            document(
                &format!(r#"{mc} xmlns:m2="{MC}" mc:Ignorable="x" m2:Ignorable="x""#),
                PARAGRAPH,
            ),
            Err(Refusal::Unproven),
        ),
    ];
    for (label, xml, expected) in cases {
        let actual = verdict(&xml);
        assert_eq!(actual, expected, "{label}: {xml}");
        if let Ok(admission) = actual {
            assert_projection(xml.as_bytes(), &word_profile(), admission, &[X]);
        }
    }
}

#[test]
fn rebinding_a_prefix_follows_the_namespace_not_the_spelling() {
    // `x` is ignorable at the root, then rebound to WordprocessingML: the
    // nested `x:r` is a run, kept by processing, and nothing is dropped.
    let rebound = document(
        &format!(r#"xmlns:mc="{MC}" mc:Ignorable="x""#),
        &format!(r#"<w:p xmlns:x="{W}"><x:r><x:t>A</x:t></x:r></w:p>"#),
    );
    assert_eq!(verdict(&rebound), Ok(Admission::Ignorable));
    assert_projection(
        rebound.as_bytes(),
        &word_profile(),
        Admission::Ignorable,
        &[X],
    );

    // `y` names another namespace at the root and the ignorable one below.
    let aliased = document(
        &format!(r#"xmlns:mc="{MC}" xmlns:y="urn:other" mc:Ignorable="x""#),
        &format!(r#"<w:p><y:keep/><w:r xmlns:y="{X}"><y:drop/><w:t>A</w:t></w:r></w:p>"#),
    );
    assert_eq!(verdict(&aliased), Ok(Admission::IgnorableOpaque));
    assert_projection(
        aliased.as_bytes(),
        &word_profile(),
        Admission::IgnorableOpaque,
        &[X],
    );
}

#[test]
fn malformed_or_unbounded_parts_are_never_admitted() {
    let mc = format!(r#"xmlns:mc="{MC}""#);
    let limits = Limits::default();
    for (label, xml) in [
        ("unbound element prefix", document(&mc, "<w:p><u:x/></w:p>")),
        (
            "unbound attribute prefix",
            document(&mc, r#"<w:p u:x="1"/>"#),
        ),
        ("not a qualified name", document(&mc, r#"<w:p a:b:c="1"/>"#)),
        (
            "empty prefixed declaration",
            document(&mc, r#"<w:p xmlns:z=""/>"#),
        ),
        (
            "custom entity",
            document(&mc, "<w:p><w:r><w:t>&bogus;</w:t></w:r></w:p>"),
        ),
        (
            "invalid character reference",
            document(&mc, "<w:p><w:r><w:t>&#0;</w:t></w:r></w:p>"),
        ),
        (
            "custom entity in a value",
            document(&mc, r#"<w:p w:x="&bogus;"/>"#),
        ),
        (
            "duplicate expanded attribute",
            document(
                &format!(r#"{mc} xmlns:y="{X}""#),
                r#"<w:p x:a="1" y:a="2"/>"#,
            ),
        ),
        ("mismatched end tag", document(&mc, "<w:p></w:r>")),
        (
            "unterminated part",
            format!(r#"<w:document xmlns:w="{W}" {mc}><w:body>"#),
        ),
        (
            "two document elements",
            format!(r#"<w:document xmlns:w="{W}" {mc}/><w:document xmlns:w="{W}"/>"#),
        ),
        (
            "late declaration",
            format!(r#"<w:document xmlns:w="{W}" {mc}/><?xml version="1.0"?>"#),
        ),
        (
            "document type declaration",
            format!(r#"<!DOCTYPE w:document><w:document xmlns:w="{W}" {mc}/>"#),
        ),
        (
            "processing instruction",
            format!(r#"<?pi data?><w:document xmlns:w="{W}" {mc}/>"#),
        ),
        (
            "rebound xml prefix",
            document(&mc, r#"<w:p xmlns:xml="urn:not-xml"/>"#),
        ),
    ] {
        let processed = process_markup_compatibility(xml.as_bytes(), &word_profile(), &limits);
        let admitted = verdict(&xml);
        assert!(admitted.is_err(), "{label}: admitted {admitted:?}: {xml}");
        // Where processing itself fails, the proof can only refuse.
        if processed.is_err() {
            assert_eq!(admitted, Err(Refusal::Unproven), "{label}");
        }
    }

    // Names and values must be UTF-8, as the processor decodes them; text and
    // comments are copied as bytes by both.
    let mut value = document(&mc, PARAGRAPH).into_bytes();
    let at = value
        .windows(6)
        .position(|window| window == b"<w:p><")
        .unwrap()
        + 4;
    value.splice(at..at, b" w:x=\"\xff\"".iter().copied());
    assert!(process_markup_compatibility(&value, &word_profile(), &limits).is_err());
    assert_eq!(
        admit(&value, &word_profile(), &limits),
        Err(Refusal::Unproven)
    );
    let mut comment = document(&mc, PARAGRAPH).into_bytes();
    comment.extend_from_slice(b"<!--\xff-->");
    assert_eq!(
        admit(&comment, &word_profile(), &limits),
        Ok(Admission::DeclaredOnly)
    );
    assert!(process_markup_compatibility(&comment, &word_profile(), &limits).is_ok());

    let deep = document(
        &mc,
        &format!("{}{}", "<w:sdt>".repeat(300), "</w:sdt>".repeat(300)),
    );
    assert!(process_markup_compatibility(deep.as_bytes(), &word_profile(), &limits).is_err());
    assert_eq!(verdict(&deep), Err(Refusal::Unproven));

    let tight = Limits {
        max_attributes_per_element: 2,
        ..Limits::default()
    };
    let wide = document(&mc, r#"<w:p w:a="1" w:b="2" w:c="3"/>"#);
    assert!(process_markup_compatibility(wide.as_bytes(), &word_profile(), &tight).is_err());
    assert_eq!(
        admit(wide.as_bytes(), &word_profile(), &tight),
        Err(Refusal::Unproven)
    );

    let small = Limits {
        max_input_bytes: 64,
        ..Limits::default()
    };
    assert_eq!(
        admit(document("", PARAGRAPH).as_bytes(), &word_profile(), &small),
        Err(Refusal::Unproven)
    );

    let mut extended = word_profile();
    extended.preserve_extension_element(crate::mce::Name {
        namespace: X.into(),
        local_name: "ext".into(),
    });
    assert_eq!(
        admit(document(&mc, PARAGRAPH).as_bytes(), &extended, &limits),
        Err(Refusal::Unproven)
    );
}

/// Every combination of document-element markup and body below, with the
/// namespace bound literally and through references: whenever the proof
/// admits, processing succeeds and its projection is what the verdict says.
#[test]
fn every_admission_is_sound_against_the_processor() {
    let roots = [
        "",
        r#"xmlns:mc="MC""#,
        r#"xmlns:mc="MC" mc:Ignorable="w14""#,
        r#"xmlns:mc="MC" mc:Ignorable="x""#,
        r#"xmlns:mc="MC" mc:Ignorable="x w14""#,
        r#"xmlns:mc="MC" mc:Ignorable="""#,
        r#"xmlns:mc="MC" mc:Ignorable="u""#,
        r#"xmlns:mc="MC" mc:Ignorable="x x""#,
        r#"xmlns:mc="MC" mc:Ignorable="mc""#,
        r#"xmlns:mc="MC" mc:Ignorable="x" mc:ProcessContent="x:wrap""#,
        r#"xmlns:mc="MC" mc:MustUnderstand="x""#,
        r#"xmlns:mc="MC" mc:MustUnderstand="w14""#,
        r#"xmlns:mc="MC" mc:Bogus="1""#,
        r#"xmlns:mc="MC" mc:Ignorable="x" mc:PreserveElements="x:keep""#,
        r#"xmlns:mc="MC" xmlns:y="urn:x" mc:Ignorable="x" y:dup="1" x:dup="2""#,
        r#"xmlns="MC""#,
        r#"xmlns:mc="MC" x:attr="1" mc:Ignorable="x""#,
    ];
    let bodies = [
        PARAGRAPH,
        r#"<w:p x:id="1"><w:r><x:flag/><w:t>A</w:t></w:r></w:p>"#,
        r#"<w:p><x:ext><w:r><w:t>HIDDEN</w:t></w:r></x:ext><w:r><w:t>B</w:t></w:r></w:p>"#,
        r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:t>C</w:t></mc:Choice><mc:Fallback><w:t>F</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>"#,
        r#"<mc:AlternateContent><mc:Choice Requires="w14"><w:p/></mc:Choice><mc:Fallback/></mc:AlternateContent>"#,
        r#"<w:p><x:wrap><w:r><w:t>W</w:t></w:r></x:wrap></w:p>"#,
        r#"<w:p w14:paraId="1"><w:r><w:t xml:space="preserve"> &amp;&#65;&lt; </w:t></w:r></w:p>"#,
        "<w:p><w:r><w:t>&bogus;</w:t></w:r></w:p>",
        "<w:p><u:foo/></w:p>",
        r#"<w:p mc:Ignorable="x"/>"#,
        r#"<w:p><x:ext r:id="rId1"/></w:p>"#,
        r#"<w:p><x:ext xml:space="preserve">t<x:in a="1"/><![CDATA[c]]></x:ext></w:p>"#,
        r#"<w:p xmlns:mc2="MC"/>"#,
        r#"<w:p><x:keep/><!-- note --></w:p>"#,
        r#"<w:p w:x="&#x9;a&#10;&quot;'"/>"#,
        r#"<w:p xmlns:x="WORD"><x:r><x:t>R</x:t></x:r></w:p>"#,
        r#"<w:p><x:ext><x:ext><y2:n xmlns:y2="urn:z"/></x:ext></x:ext></w:p>"#,
        r#"<w:p mc:Ignorable="w14" w14:paraId="1"><w:r><w:t>N</w:t></w:r></w:p>"#,
        r#"<w:p><w:ins xmlns:m3="MC" m3:Ignorable="w14 xml" w14:d="1"><w:r><w:t>I</w:t></w:r></w:ins></w:p>"#,
        "",
    ];
    let mut admitted = 0usize;
    let mut refused = 0usize;
    for spelling in [MC, MC_ESCAPED] {
        for root in roots {
            for body in bodies {
                let xml = document(
                    &root.replace("MC", spelling),
                    &body.replace("MC", spelling).replace("WORD", W),
                );
                let processed = process_markup_compatibility(
                    xml.as_bytes(),
                    &word_profile(),
                    &Limits::default(),
                );
                // The held proof is the plain proof wherever it finds no
                // wrapper, and admits a wrapper only where the plain proof
                // refuses an element.
                match held(&xml) {
                    Ok(held) if held.wrappers == 0 => {
                        assert_eq!(Ok(held.outside), verdict(&xml), "{xml}");
                    },
                    Ok(_) => assert_eq!(verdict(&xml), Err(Refusal::Element), "{xml}"),
                    Err(_) => assert!(verdict(&xml).is_err(), "{xml}"),
                }
                match verdict(&xml) {
                    Ok(admission) => {
                        admitted += 1;
                        assert!(processed.is_ok(), "admitted but processing fails: {xml}");
                        assert_projection(xml.as_bytes(), &word_profile(), admission, &[X]);
                    },
                    Err(_) => refused += 1,
                }
            }
        }
    }
    assert!(
        admitted > 40 && refused > 100,
        "{admitted} admitted, {refused} refused"
    );
}

#[test]
fn baseline_profile_drops_word_extensions_as_opaque_leaves() {
    // The shape of Word 2013+ settings: `w14`/`w15` leaves the baseline
    // profile does not understand, beside the variables a reader looks for.
    let settings = format!(
        r#"<w:settings xmlns:w="{W}" xmlns:mc="{MC}" xmlns:w14="{W14}" xmlns:w15="urn:w15" mc:Ignorable="w14 w15"><w:zoom w:percent="100"/><w:docVars><w:docVar w:name="a" w:val="1"/></w:docVars><w14:docId w14:val="7F00AA10"/><w15:chartTrackingRefBased/></w:settings>"#
    );
    let baseline = Capabilities::default();
    assert_eq!(
        admit(settings.as_bytes(), &baseline, &Limits::default()),
        Ok(Admission::IgnorableOpaque)
    );
    assert_projection(
        settings.as_bytes(),
        &baseline,
        Admission::IgnorableOpaque,
        &[W14, "urn:w15"],
    );
    // The Word main-document profile understands `w14` and keeps it.
    let main_like = settings.replace("<w15:chartTrackingRefBased/>", "");
    assert_eq!(verdict(&main_like), Ok(Admission::Ignorable));
}

#[test]
fn observer_sees_every_event_and_can_stop_the_proof() {
    let xml = document(&format!(r#"xmlns:mc="{MC}" mc:Ignorable="w14""#), PARAGRAPH);
    let mut events = 0usize;
    let mut bytes = 0usize;
    let verdict = admit_with(
        xml.as_bytes(),
        &word_profile(),
        &Limits::default(),
        |length| {
            events += 1;
            bytes += length;
            Ok::<(), ()>(())
        },
    )
    .unwrap();
    assert_eq!(verdict, Ok(Admission::Ignorable));
    assert!(events >= 12, "{events}");
    assert!(bytes >= xml.len() - 1, "{bytes} of {}", xml.len());

    let mut seen = 0usize;
    let stopped = admit_with(xml.as_bytes(), &word_profile(), &Limits::default(), |_| {
        seen += 1;
        if seen == 3 { Err("stop") } else { Ok(()) }
    });
    assert_eq!(stopped, Err("stop"));
    assert_eq!(seen, 3);

    // An unbound part is decided before any event.
    let mut unbound_events = 0usize;
    let unbound = admit_with(
        document("", PARAGRAPH).as_bytes(),
        &word_profile(),
        &Limits::default(),
        |_| {
            unbound_events += 1;
            Ok::<(), ()>(())
        },
    )
    .unwrap();
    assert_eq!(unbound, Ok(Admission::Unbound));
    assert_eq!(unbound_events, 0);
}

/// The main document's holder: a WordprocessingML run.
const RUN: Holder<'static> = Holder {
    namespaces: &[W],
    local: "r",
};

fn held(xml: &str) -> Result<Held, Refusal> {
    admit_held(xml.as_bytes(), &word_profile(), &Limits::default(), RUN)
}

const RUN_WRAPPER: &str = r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:t>C</w:t></mc:Choice><mc:Fallback><w:t>F</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>"#;

#[test]
fn held_wrappers_are_opaque_only_inside_holders() {
    let mc = format!(r#"xmlns:mc="{MC}""#);
    let ignorable = format!(r#"{mc} mc:Ignorable="x""#);
    let cases: Vec<(&str, String, Result<usize, Refusal>)> = vec![
        ("a wrapper in a run", document(&mc, RUN_WRAPPER), Ok(1)),
        (
            "two runs with wrappers",
            document(&mc, &RUN_WRAPPER.repeat(2)),
            Ok(2),
        ),
        (
            "a wrapper nested in a branch counts once",
            document(
                &mc,
                r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w14"><mc:AlternateContent><mc:Choice Requires="w14"/></mc:AlternateContent></mc:Choice></mc:AlternateContent></w:r></w:p>"#,
            ),
            Ok(1),
        ),
        (
            "a wrapper deep inside a run's drawing",
            document(
                &mc,
                r#"<w:p><w:r><w:drawing><x:anchor><mc:AlternateContent><mc:Choice Requires="w14"><x:size/></mc:Choice></mc:AlternateContent></x:anchor></w:drawing></w:r></w:p>"#,
            ),
            Ok(1),
        ),
        (
            "an empty wrapper with its own declaration",
            document(
                &mc,
                r#"<w:p><w:r><mc:AlternateContent xmlns:q="urn:q"/><w:t>A</w:t></w:r></w:p>"#,
            ),
            Ok(1),
        ),
        (
            "anything inside a wrapper",
            document(
                &mc,
                r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w14" mc:Ignorable="zz" xmlns:zz="urn:zz"><zz:any mc:ProcessContent="zz:any"><w:p/></zz:any></mc:Choice><mc:Fallback><mc:Bogus/></mc:Fallback></mc:AlternateContent></w:r></w:p>"#,
            ),
            Ok(1),
        ),
        (
            "outside wrappers the plain verdict holds",
            document(&ignorable, r#"<w:p x:id="1"><w:r><w:t>A</w:t></w:r></w:p>"#),
            Ok(0),
        ),
        (
            "a wrapper in a paragraph, outside any run",
            document(
                &mc,
                r#"<w:p><mc:AlternateContent><mc:Choice Requires="w14"><w:r/></mc:Choice></mc:AlternateContent></w:p>"#,
            ),
            Err(Refusal::Element),
        ),
        (
            "a body-level wrapper",
            document(
                &mc,
                r#"<mc:AlternateContent><mc:Choice Requires="w14"><w:p/></mc:Choice><mc:Fallback/></mc:AlternateContent>"#,
            ),
            Err(Refusal::Element),
        ),
        (
            "a run of another namespace",
            document(
                &mc,
                r#"<w:p><x:r><mc:AlternateContent><mc:Choice Requires="w14"/></mc:AlternateContent></x:r></w:p>"#,
            ),
            Err(Refusal::Element),
        ),
        (
            "a branch without its wrapper",
            document(
                &mc,
                r#"<w:p><w:r><mc:Choice Requires="w14"><w:t>C</w:t></mc:Choice></w:r></w:p>"#,
            ),
            Err(Refusal::Element),
        ),
        (
            "a wrapper in a dropped subtree of a run",
            document(
                &ignorable,
                r#"<w:p><w:r><x:ext><mc:AlternateContent/></x:ext></w:r></w:p>"#,
            ),
            Err(Refusal::Element),
        ),
        (
            "a directive outside the wrappers",
            document(
                &mc,
                &format!(
                    r#"<w:p mc:ProcessContent="x:wrap">{}</w:p>"#,
                    &RUN_WRAPPER[5..RUN_WRAPPER.len() - 6]
                ),
            ),
            Err(Refusal::Directive),
        ),
        (
            "an unbalanced wrapper",
            document(
                &mc,
                r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w14"></w:r></w:p>"#,
            ),
            Err(Refusal::Unproven),
        ),
    ];
    for (name, xml, expected) in cases {
        let verdict = held(&xml);
        assert_eq!(verdict.map(|held| held.wrappers), expected, "{name}: {xml}");
        if let Ok(held) = verdict {
            if held.wrappers == 0 {
                assert_eq!(Ok(held.outside), self::verdict(&xml), "{name}");
            } else {
                assert_eq!(self::verdict(&xml), Err(Refusal::Element), "{name}");
            }
        }
    }
    // Inside a wrapper the processor's own checks still decide: the proof
    // holds, processing fails.
    let unbound = document(
        &mc,
        r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="nowhere"/></mc:AlternateContent></w:r></w:p>"#,
    );
    assert_eq!(held(&unbound).map(|held| held.wrappers), Ok(1));
    assert!(
        process_markup_compatibility(unbound.as_bytes(), &word_profile(), &Limits::default())
            .is_err()
    );
}

#[test]
fn processing_replaces_each_held_wrapper_in_place() {
    // Each pair is a body with wrappers and the body processing must
    // produce: the selected branch's content where the wrapper stood.
    let pairs = [
        (
            r#"<w:p><w:r><w:rPr><w:b/></w:rPr><mc:AlternateContent><mc:Choice Requires="w14"><w:t>C</w:t></mc:Choice><mc:Fallback><w:t>F</w:t></mc:Fallback></mc:AlternateContent><w:t>T</w:t></w:r></w:p><w:p/>"#,
            r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>C</w:t><w:t>T</w:t></w:r></w:p><w:p/>"#,
        ),
        (
            r#"<w:p><w:r><mc:AlternateContent xmlns:q="urn:q"><mc:Choice Requires="q"><w:t>C</w:t></mc:Choice><mc:Fallback><w:t>F</w:t></mc:Fallback></mc:AlternateContent></w:r><w:r><w:t>2</w:t></w:r></w:p>"#,
            r#"<w:p><w:r><w:t>F</w:t></w:r><w:r><w:t>2</w:t></w:r></w:p>"#,
        ),
        (
            r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="x"><w:t>C</w:t></mc:Choice></mc:AlternateContent><w:t>T</w:t></w:r></w:p>"#,
            r#"<w:p><w:r><w:t>T</w:t></w:r></w:p>"#,
        ),
    ];
    for root in [
        format!(r#"xmlns:mc="{MC}""#),
        format!(r#"xmlns:mc="{MC}" mc:Ignorable="w14""#),
        format!(r#"xmlns:mc="{MC}" mc:Ignorable="x""#),
    ] {
        for (body, selected) in pairs {
            let part = document(&root, body);
            let held = held(&part).unwrap_or_else(|refusal| panic!("{refusal:?}: {part}"));
            assert!(held.wrappers > 0);
            let projected =
                process_markup_compatibility(part.as_bytes(), &word_profile(), &Limits::default())
                    .unwrap();
            let selected = document(&root, selected);
            let expected = process_markup_compatibility(
                selected.as_bytes(),
                &word_profile(),
                &Limits::default(),
            )
            .unwrap();
            assert_eq!(
                nodes(&projected.xml, &[]),
                nodes(&expected.xml, &[]),
                "{part}"
            );
        }
    }
}
