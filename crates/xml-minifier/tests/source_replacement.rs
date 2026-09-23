//! `verify_source_replacement` returns exactly what auditing the original and
//! then the replacement with `verify_source` returns, and proves a local
//! edit by scanning only the replaced element.
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "each test states one fixed expected verdict"
)]

use xml_minifier::audit::{
    Error, Limits, ReplacementError, ReplacementProof, Resource, verify_source,
    verify_source_replacement,
};

/// The two audits the pair API replaces, in their order.
fn oracle(original: &[u8], replacement: &[u8], limits: Limits) -> Result<(), ReplacementError> {
    let _original = verify_source(original, limits).map_err(ReplacementError::Original)?;
    let _replacement = verify_source(replacement, limits).map_err(ReplacementError::Replacement)?;
    Ok(())
}

/// Run the pair API and require its verdict to be the oracle's.
fn checked(
    original: &[u8],
    replacement: &[u8],
    limits: Limits,
) -> Result<ReplacementProof, ReplacementError> {
    let actual = verify_source_replacement(original, replacement, limits);
    let expected = oracle(original, replacement, limits);
    match (&actual, &expected) {
        (Ok(_), Ok(())) => {},
        (Err(actual_error), Err(expected_error)) => assert_eq!(actual_error, expected_error),
        _ => panic!(
            "pair verdict {actual:?} differs from the oracle {expected:?}\noriginal: {}\nreplacement: {}",
            String::from_utf8_lossy(original),
            String::from_utf8_lossy(replacement)
        ),
    }
    actual
}

fn window(
    original: std::ops::Range<usize>,
    replacement: std::ops::Range<usize>,
) -> ReplacementProof {
    ReplacementProof::Window {
        original,
        replacement,
    }
}

fn span(document: &str, needle: &str) -> std::ops::Range<usize> {
    let start = document.find(needle).expect("fixture needle");
    start..start + needle.len()
}

const WORKSHEET_HEAD: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<worksheet xmlns=\"urn:sml\"><dimension ref=\"A1:C3\"/>\n  <sheetData>\n";

fn worksheet(a1: &str) -> String {
    let mut sheet = String::from(WORKSHEET_HEAD);
    for row in 1..=3 {
        sheet.push_str(&format!("    <row r=\"{row}\">"));
        for column in ["A", "B", "C"] {
            if row == 1 && column == "A" {
                sheet.push_str(a1);
            } else {
                sheet.push_str(&format!("<c r=\"{column}{row}\"><v>{row}</v></c>"));
            }
        }
        sheet.push_str("</row>\n");
    }
    sheet.push_str("  </sheetData>\n  <extLst><ext uri=\"{X}\"><x:y xmlns:x=\"urn:x\" a=\"1\"/></ext></extLst>\n</worksheet>");
    sheet
}

#[test]
fn a_one_cell_value_edit_is_proved_by_its_value_element() {
    let original = worksheet("<c r=\"A1\"><v>0</v></c>");
    let replacement = worksheet("<c r=\"A1\"><v>1</v></c>");
    let proof = checked(
        original.as_bytes(),
        replacement.as_bytes(),
        Limits::default(),
    )
    .expect("both sides are valid");
    let value = span(&original, "<v>0</v>");
    assert_eq!(proof, window(value.clone(), value));
}

#[test]
fn a_length_changing_cell_edit_is_proved_by_the_enclosing_element() {
    let original = worksheet("<c r=\"A1\" s=\"1\"><v>7</v></c>");
    let replacement = worksheet("<c r=\"A1\"><v>123456</v></c>");
    let proof = checked(
        original.as_bytes(),
        replacement.as_bytes(),
        Limits::default(),
    )
    .expect("both sides are valid");
    let cell = span(&original, "<c r=\"A1\" s=\"1\"><v>7</v></c>");
    let edited = span(&replacement, "<c r=\"A1\"><v>123456</v></c>");
    assert_eq!(proof, window(cell, edited));
}

#[test]
fn identical_payloads_have_the_original_verdict() {
    let original = worksheet("<c r=\"A1\"><v>0</v></c>");
    assert_eq!(
        checked(original.as_bytes(), original.as_bytes(), Limits::default()),
        Ok(ReplacementProof::Identical)
    );
    let broken = b"<root><!DOCTYPE x></root>";
    assert!(matches!(
        checked(broken, broken, Limits::default()),
        Err(ReplacementError::Original(Error::Doctype { .. }))
    ));
}

#[test]
fn an_original_failure_precedes_any_replacement() {
    let original = worksheet("<c r=\"A1\"><v xml:space=\"bogus\">0</v></c>");
    let replacement = worksheet("<c r=\"A1\"><v>1</v></c>");
    let error = checked(
        original.as_bytes(),
        replacement.as_bytes(),
        Limits::default(),
    )
    .expect_err("the original is refused");
    assert!(matches!(
        error,
        ReplacementError::Original(Error::Malformed { .. })
    ));
}

#[test]
fn a_defect_inside_the_window_is_refused_with_the_complete_audit_error() {
    let original = worksheet("<c r=\"A1\"><v>0</v></c>");
    for edited in [
        "<c r=\"A1\"><v xml:space=\"bogus\">1</v></c>",
        "<c r=\"A1\"><v>1</w></c>",
        "<c r=\"A1\"><v>1</v></c></c>",
        "<c r=\"A1\"><v><!DOCTYPE x>1</v></c>",
        "<c r=\"A1\"><v a=\"1\"b=\"2\">1</v></c>",
        "<c r=\"A1\"><v>1&unclosed</v></c>",
        "<c r=\"A1\"><v>1<</v></c>",
    ] {
        let replacement = worksheet(edited);
        let error = checked(
            original.as_bytes(),
            replacement.as_bytes(),
            Limits::default(),
        )
        .expect_err("the replacement is refused");
        assert!(
            matches!(error, ReplacementError::Replacement(_)),
            "{edited}: {error:?}"
        );
    }
    let mut invalid = worksheet("<c r=\"A1\"><v>0</v></c>").into_bytes();
    let at = invalid.windows(3).position(|w| w == b">0<").unwrap() + 1;
    invalid[at] = 0xFF;
    assert_eq!(
        checked(original.as_bytes(), &invalid, Limits::default()),
        Err(ReplacementError::Replacement(Error::Encoding {
            valid_up_to: at
        }))
    );
}

#[test]
fn tampering_outside_the_edit_is_scanned_and_refused() {
    // The intended edit is one value; a second, distant change carries a
    // defect. The window must cover both, so the defect is audited.
    let original = worksheet("<c r=\"A1\"><v>0</v></c>");
    let tampered = worksheet("<c r=\"A1\"><v>1</v></c>").replace(
        "<x:y xmlns:x=\"urn:x\" a=\"1\"/>",
        "<x:y xmlns:x=\"urn:x\" xml:space=\"x\"/>",
    );
    let expected = verify_source(tampered.as_bytes(), Limits::default())
        .expect_err("the substituted bytes are invalid");
    assert_eq!(
        checked(original.as_bytes(), tampered.as_bytes(), Limits::default()),
        Err(ReplacementError::Replacement(expected))
    );

    // A distant valid substitution is audited too: the proof no longer
    // claims the small window.
    let substituted = worksheet("<c r=\"A1\"><v>1</v></c>").replace(" a=\"1\"", " a=\"2\"");
    let proof = checked(
        original.as_bytes(),
        substituted.as_bytes(),
        Limits::default(),
    )
    .expect("both sides are valid");
    assert!(
        !matches!(&proof, ReplacementProof::Window { original: covered, .. } if covered.len() < 100),
        "{proof:?}"
    );
}

#[test]
fn edits_of_the_document_element_or_between_its_children_are_scanned_completely() {
    let original = worksheet("<c r=\"A1\"><v>0</v></c>");
    for replacement in [
        original.replace(
            "<worksheet xmlns=\"urn:sml\">",
            "<worksheet xmlns=\"urn:sml\" x=\"1\">",
        ),
        original.replace("</worksheet>", "</worksheet >"),
        original.replace("\n  <sheetData>", "\n\n  <sheetData>"),
        original.replacen("\r\n", "\n", 1),
        original
            .replace("<dimension ref=\"A1:C3\"/>", "<dimension ref=\"A1:D3\"/>")
            .replace("<v>0</v>", "<v>1</v>"),
    ] {
        assert_eq!(
            checked(
                original.as_bytes(),
                replacement.as_bytes(),
                Limits::default()
            ),
            Ok(ReplacementProof::Complete),
            "{replacement}"
        );
    }
}

#[test]
fn a_window_that_would_not_start_or_end_with_markup_is_widened() {
    // An element replaced by text: the bytes that took its place begin with
    // character data, which would merge with the text before it.
    let original = "<r><s> <p/></s><t/></r>";
    let replacement = "<r><s> x</s><t/></r>";
    let proof = checked(
        original.as_bytes(),
        replacement.as_bytes(),
        Limits::default(),
    )
    .expect("both sides are valid");
    assert_eq!(
        proof,
        window(
            span(original, "<s> <p/></s>"),
            span(replacement, "<s> x</s>")
        )
    );

    // An element followed by new text: the window must end with markup.
    let original = "<r><s><p/></s>tail<pad>padding that keeps the window small</pad></r>";
    let replacement = "<r><s><p/>x</s>tail<pad>padding that keeps the window small</pad></r>";
    let proof = checked(
        original.as_bytes(),
        replacement.as_bytes(),
        Limits::default(),
    )
    .expect("both sides are valid");
    assert_eq!(
        proof,
        window(
            span(original, "<s><p/></s>"),
            span(replacement, "<s><p/>x</s>")
        )
    );

    // Text replacing the only child of the document element leaves no
    // window at all.
    let original = "<r><p/></r>";
    let replacement = "<r>x</r>";
    assert_eq!(
        checked(
            original.as_bytes(),
            replacement.as_bytes(),
            Limits::default()
        ),
        Ok(ReplacementProof::Complete)
    );
}

#[test]
fn an_insertion_between_siblings_is_proved_with_a_neighbouring_sibling() {
    // The common prefix ends inside the `<` both payloads share, so the
    // following sibling is the innermost element covering the difference.
    let original = "<r><s><a/><b/></s><z/></r>";
    let replacement = "<r><s><a/><n>1</n><b/></s><z/></r>";
    let proof = checked(
        original.as_bytes(),
        replacement.as_bytes(),
        Limits::default(),
    )
    .expect("both sides are valid");
    assert_eq!(
        proof,
        window(span(original, "<b/>"), span(replacement, "<n>1</n><b/>"))
    );

    // After a text node the difference starts after the shared `>`, so the
    // preceding sibling covers it.
    let original = "<r><s><a/>t</s><z/></r>";
    let replacement = "<r><s><a/><n>1</n>t</s><z/></r>";
    let proof = checked(
        original.as_bytes(),
        replacement.as_bytes(),
        Limits::default(),
    )
    .expect("both sides are valid");
    assert_eq!(
        proof,
        window(span(original, "<a/>"), span(replacement, "<a/><n>1</n>"))
    );
}

#[test]
fn a_window_larger_than_half_the_replacement_is_scanned_completely() {
    let original = "<r><s><a/></s></r>";
    let replacement = "<r><s><a/><b/><c/><d/><e/></s></r>";
    assert_eq!(
        checked(
            original.as_bytes(),
            replacement.as_bytes(),
            Limits::default()
        ),
        Ok(ReplacementProof::Complete)
    );
}

#[test]
fn a_marked_original_keeps_physical_offsets() {
    let original = format!("\u{feff}{}", worksheet("<c r=\"A1\"><v>0</v></c>"));
    let replacement = format!("\u{feff}{}", worksheet("<c r=\"A1\"><v>1</v></c>"));
    let proof = checked(
        original.as_bytes(),
        replacement.as_bytes(),
        Limits::default(),
    )
    .expect("both sides are valid");
    let value = span(&original, "<v>0</v>");
    assert_eq!(proof, window(value.clone(), value));
}

fn narrow(resource: Resource, maximum: usize) -> Limits {
    Limits::default().narrow(resource, maximum)
}

/// Every aggregate budget is re-totalled for the whole replacement: an edit
/// that stays inside the window can still push the document over a budget,
/// and exactly-at-budget stays accepted.
#[test]
fn every_budget_is_re_totalled_for_the_whole_replacement() {
    let original = worksheet("<c r=\"A1\"><v>0</v></c>");
    let at = |edited: &str| worksheet(edited);

    let report = verify_source(original.as_bytes(), Limits::default()).unwrap();
    // Events: the edit adds a comment event inside the value element.
    let grown = at("<c r=\"A1\"><v>1<!--x--></v></c>");
    let limits = narrow(Resource::Events, report.events() + 1);
    assert!(matches!(
        checked(original.as_bytes(), grown.as_bytes(), limits),
        Ok(ReplacementProof::Window { .. })
    ));
    let limits = narrow(Resource::Events, report.events());
    assert!(matches!(
        checked(original.as_bytes(), grown.as_bytes(), limits),
        Err(ReplacementError::Replacement(Error::Limit {
            resource: Resource::Events,
            ..
        }))
    ));

    // Attributes.
    let grown = at("<c r=\"A1\"><v t=\"1\">1</v></c>");
    let limits = narrow(Resource::Attributes, report.attributes() + 1);
    assert!(matches!(
        checked(original.as_bytes(), grown.as_bytes(), limits),
        Ok(ReplacementProof::Window { .. })
    ));
    let limits = narrow(Resource::Attributes, report.attributes());
    assert!(matches!(
        checked(original.as_bytes(), grown.as_bytes(), limits),
        Err(ReplacementError::Replacement(Error::Limit {
            resource: Resource::Attributes,
            ..
        }))
    ));

    // Character data.
    let grown = at("<c r=\"A1\"><v>12</v></c>");
    let limits = narrow(Resource::TextBytes, report.text_bytes() + 1);
    assert!(matches!(
        checked(original.as_bytes(), grown.as_bytes(), limits),
        Ok(ReplacementProof::Window { .. })
    ));
    let limits = narrow(Resource::TextBytes, report.text_bytes());
    assert!(matches!(
        checked(original.as_bytes(), grown.as_bytes(), limits),
        Err(ReplacementError::Replacement(Error::Limit {
            resource: Resource::TextBytes,
            ..
        }))
    ));

    // Depth: the value element sits at depth five; nesting inside it is
    // checked against the absolute limit.
    let deeper = at("<c r=\"A1\"><v><w>1</w></v></c>");
    assert!(matches!(
        checked(
            original.as_bytes(),
            deeper.as_bytes(),
            narrow(Resource::Depth, 6)
        ),
        Ok(ReplacementProof::Window { .. })
    ));
    assert!(matches!(
        checked(
            original.as_bytes(),
            deeper.as_bytes(),
            narrow(Resource::Depth, 5)
        ),
        Err(ReplacementError::Replacement(Error::Limit {
            resource: Resource::Depth,
            ..
        }))
    ));

    // One token. The original's longest token is its 55-byte declaration.
    let limits = narrow(Resource::TokenBytes, 64);
    let long = at(&format!("<c r=\"A1\"><v>{}</v></c>", "9".repeat(64)));
    assert!(matches!(
        checked(original.as_bytes(), long.as_bytes(), limits),
        Ok(ReplacementProof::Window { .. })
    ));
    let longer = at(&format!("<c r=\"A1\"><v>{}</v></c>", "9".repeat(65)));
    assert!(matches!(
        checked(original.as_bytes(), longer.as_bytes(), limits),
        Err(ReplacementError::Replacement(Error::Limit {
            resource: Resource::TokenBytes,
            ..
        }))
    ));

    // Total length.
    let longer = at("<c r=\"A1\"><v>10</v></c>");
    assert!(matches!(
        checked(
            original.as_bytes(),
            longer.as_bytes(),
            narrow(Resource::Bytes, longer.len())
        ),
        Ok(ReplacementProof::Window { .. })
    ));
    assert!(matches!(
        checked(
            original.as_bytes(),
            longer.as_bytes(),
            narrow(Resource::Bytes, original.len())
        ),
        Err(ReplacementError::Replacement(Error::Limit {
            resource: Resource::Bytes,
            ..
        }))
    ));
}

// ------------------------------------------------------------ differential

/// A small deterministic generator (xorshift64*), so the differential test
/// needs no dependency and every failure is reproducible from its seed.
struct Rng(u64);

impl Rng {
    fn next(&mut self) -> u64 {
        let mut x = self.0;
        x ^= x >> 12;
        x ^= x << 25;
        x ^= x >> 27;
        self.0 = x;
        x.wrapping_mul(0x2545_F491_4F6C_DD1D)
    }

    fn below(&mut self, bound: usize) -> usize {
        (self.next() % bound as u64) as usize
    }

    fn chance(&mut self, percent: usize) -> bool {
        self.below(100) < percent
    }

    fn pick<'a>(&mut self, items: &[&'a str]) -> &'a str {
        items[self.below(items.len())]
    }
}

const NAMES: &[&str] = &["a", "b", "c", "row", "v", "x:y", "p", "q"];
/// `ATTRIBUTES[0]` and `[1]` have distinct keys, `[2..4]` share
/// `xml:space`, `[4]` is valid with a `>` in its value; the rest are
/// malformed or refused by the audit.
const ATTRIBUTES: &[&str] = &[
    " r=\"1\"",
    " s='2'",
    " xml:space=\"preserve\"",
    " xml:space=\"default\"",
    " g=\"a>b\"",
    " xml:space=\"bogus\"",
    " xmlns:x=\"urn:x\"",
    " t=\"a&amp;b\"",
    " u = \"3\"",
    " w=\"4\"z=\"5\"",
    " k",
    " m=\"",
    " e='<'",
];
/// The first nine are valid character data; the rest are not, or are
/// refused outside the document element.
const TEXTS: &[&str] = &[
    "0", "12", " ", "\n  ", "a&amp;b", "&#65;", "a>b", "&lt;&gt;", "\u{e9}", "&bogus;", "x]]>y",
    "t\tt", "&", "<",
];
const VALID_TEXTS: usize = 9;
/// The first seven are valid markup, several with a `>` or `<` inside the
/// token; the rest are refused or unbalanced.
const MISC: &[&str] = &[
    "<!--c-->",
    "<![CDATA[d]]>",
    "<![CDATA[ ]]>",
    "<?pi x?>",
    "<!--a>b<c-->",
    "<![CDATA[x>y<z]]>",
    "<?pi a>b?>",
    "<?xml version=\"1.0\"?>",
    "<!DOCTYPE z>",
    "</a>",
    "<a>",
    "<a/>",
    "<b></b>",
    "<v>1</v>",
    "<c r=\"A1\"><v>7</v></c>",
];
const VALID_MISC: usize = 7;

fn element(rng: &mut Rng, depth: usize, out: &mut String, spans: &mut Vec<(usize, usize)>) {
    let start = out.len();
    let name = rng.pick(NAMES);
    out.push('<');
    out.push_str(name);
    // Distinct valid keys, and now and then one attribute from the whole
    // pool, which may be malformed or repeat a key.
    if rng.chance(50) {
        out.push_str(ATTRIBUTES[0]);
    }
    if rng.chance(30) {
        out.push_str(ATTRIBUTES[1]);
    }
    if rng.chance(20) {
        out.push_str(rng.pick(&ATTRIBUTES[2..4]));
    }
    if rng.chance(10) {
        out.push_str(ATTRIBUTES[4]);
    }
    if rng.chance(3) {
        out.push_str(rng.pick(ATTRIBUTES));
    }
    if depth == 0 || rng.chance(25) {
        out.push_str("/>");
        spans.push((start, out.len()));
        return;
    }
    out.push('>');
    for _ in 0..rng.below(4) {
        match rng.below(10) {
            0..=4 => element(rng, depth - 1, out, spans),
            5..=7 => out.push_str(if rng.chance(97) {
                rng.pick(&TEXTS[..VALID_TEXTS])
            } else {
                rng.pick(TEXTS)
            }),
            _ => out.push_str(if rng.chance(97) {
                rng.pick(&MISC[..VALID_MISC])
            } else {
                rng.pick(MISC)
            }),
        }
    }
    out.push_str("</");
    out.push_str(if rng.chance(99) { name } else { "zz" });
    out.push('>');
    spans.push((start, out.len()));
}

/// A generated document and the byte span of every element in it.
fn document(rng: &mut Rng) -> (Vec<u8>, Vec<(usize, usize)>) {
    let mut out = String::new();
    let mut spans = Vec::new();
    if rng.chance(5) {
        out.push('\u{feff}');
    }
    if rng.chance(50) {
        out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        if rng.chance(50) {
            out.push('\n');
        }
    }
    out.push_str("<root>");
    for _ in 0..1 + rng.below(4) {
        element(rng, 4, &mut out, &mut spans);
        if rng.chance(30) {
            out.push_str("\n  ");
        }
    }
    out.push_str("</root>");
    if rng.chance(3) {
        out.push_str(rng.pick(&["<extra/>", " x", "<![CDATA[ ]]>", "\n", "<!--t-->"]));
    }
    let mut bytes = out.into_bytes();
    if rng.chance(1) && !bytes.is_empty() {
        let at = rng.below(bytes.len());
        bytes[at] = 0xC0;
    }
    (bytes, spans)
}

fn fragment(rng: &mut Rng) -> String {
    let mut fragment = String::new();
    match rng.below(5) {
        0 | 1 => element(rng, 2, &mut fragment, &mut Vec::new()),
        2 => fragment.push_str(rng.pick(TEXTS)),
        3 => fragment.push_str(rng.pick(MISC)),
        _ => {},
    }
    fragment
}

/// A local edit of `original`: most often a whole element replaced,
/// inserted, deleted or retouched, otherwise an arbitrary span replaced.
fn edit(rng: &mut Rng, original: &[u8], spans: &[(usize, usize)]) -> Vec<u8> {
    let length = original.len();
    let (start, end, insert) = match rng.below(10) {
        _ if spans.is_empty() => {
            let start = rng.below(length + 1);
            (start, (start + rng.below(6)).min(length), fragment(rng))
        },
        0..=3 => {
            let (start, end) = spans[rng.below(spans.len())];
            (start, end, fragment(rng))
        },
        4 | 5 => {
            let (start, end) = spans[rng.below(spans.len())];
            let at = if rng.chance(50) { start } else { end };
            (at, at, fragment(rng))
        },
        6 => {
            let (start, end) = spans[rng.below(spans.len())];
            (start, end, String::new())
        },
        7 => {
            // Retouch one byte inside an element: a value, a name, a quote.
            let (start, end) = spans[rng.below(spans.len())];
            let at = start + rng.below(end - start);
            let byte = rng.pick(&["0", "7", "a", " ", "\"", "/", "<", ">", "&"]);
            (at, at + 1, byte.to_string())
        },
        _ => {
            let start = rng.below(length + 1);
            (start, (start + rng.below(6)).min(length), fragment(rng))
        },
    };
    let mut replacement = original[..start].to_vec();
    replacement.extend_from_slice(insert.as_bytes());
    replacement.extend_from_slice(&original[end..]);
    replacement
}

/// Default limits half the time; otherwise every budget within two of what
/// the original or the replacement needs, so budget boundaries are crossed
/// in both directions.
fn limits(rng: &mut Rng, original: &[u8], replacement: &[u8]) -> Limits {
    if rng.chance(50) {
        return Limits::default();
    }
    let basis = if rng.chance(50) {
        original
    } else {
        replacement
    };
    let Ok(report) = verify_source(basis, Limits::default()) else {
        return Limits::default();
    };
    let around = |rng: &mut Rng, value: usize| value.saturating_sub(2) + rng.below(5);
    let (bytes, depth, events, attributes, text) = (
        around(rng, basis.len()),
        around(rng, report.max_depth()),
        around(rng, report.events()),
        around(rng, report.attributes()),
        around(rng, report.text_bytes()),
    );
    let mut limits = Limits::default();
    for (resource, value) in [
        (Resource::Bytes, bytes),
        (Resource::Depth, depth),
        (Resource::Events, events),
        (Resource::Attributes, attributes),
        (Resource::TextBytes, text),
        (Resource::TokenBytes, 24 + rng.below(48)),
    ] {
        if rng.chance(40) {
            limits = limits.narrow(resource, value);
        }
    }
    limits
}

/// A campaign size or seed from the environment, for longer runs than the
/// default; the default is what the suite runs.
fn setting(name: &str, default: u64) -> u64 {
    std::env::var(name)
        .ok()
        .and_then(|value| value.parse().ok())
        .unwrap_or(default)
}

#[test]
fn the_pair_verdict_equals_two_complete_audits_on_generated_edits() {
    let cases = setting("XML_MINIFIER_REPLACEMENT_CASES", 40_000);
    let mut rng = Rng(setting(
        "XML_MINIFIER_REPLACEMENT_SEED",
        0x0747_5eed_0747_5eed,
    ));
    let scale = |count: u64| count * cases / 40_000;
    let (mut windows, mut accepted, mut refused) = (0u64, 0u64, 0u64);
    for _ in 0..cases {
        let (original, spans) = document(&mut rng);
        let replacement = edit(&mut rng, &original, &spans);
        let limits = limits(&mut rng, &original, &replacement);
        match checked(&original, &replacement, limits) {
            Ok(ReplacementProof::Window { .. }) => {
                windows += 1;
                accepted += 1;
            },
            Ok(_) => accepted += 1,
            Err(_) => refused += 1,
        }
    }
    eprintln!("accepted {accepted} (windows {windows}), refused {refused}");
    // The generator must exercise every path, and the window path often.
    assert!(accepted > scale(8_000), "accepted {accepted}");
    assert!(refused > scale(8_000), "refused {refused}");
    assert!(windows > scale(4_000), "windows {windows}");
}
