// MTEF writer tests: byte-level checks plus write/parse round-trips
//
// The round-trips are the load-bearing part: every construct is written, handed
// to `MtefParser`, and the recovered AST is compared against what the reader can
// represent. Leaves come back as `MathNode::Text` holding the reader's LaTeX
// spelling, so leaf comparisons go through `flatten_text`.

/// Decorations, style runs, degradation and format limits
mod decorations;
/// Characters, operators, symbols and spaces
mod leaves;
/// Fractions, roots, scripts, fences, operators, matrices and piles
mod structures;

use super::{MtefWriteOptions, MtefWriter};
use crate::ast::{Formula, MathNode};
use crate::mtef::MtefParser;
use crate::mtef::constants::*;
use std::borrow::Cow;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Serialize nodes with the OLE equation header
fn write(nodes: &[MathNode<'_>]) -> Vec<u8> {
    MtefWriter::new()
        .write_nodes(nodes)
        .expect("serialization should succeed")
}

/// Serialize nodes without the OLE equation header
fn write_bare(nodes: &[MathNode<'_>]) -> Vec<u8> {
    MtefWriter::with_options(MtefWriteOptions::without_ole_header())
        .write_nodes(nodes)
        .expect("serialization should succeed")
}

/// Serialize nodes, parse them back, and hand the recovered AST to `check`
///
/// The recovered nodes borrow the parser's arena, so they are inspected inside
/// the closure rather than returned.
fn roundtrip<R>(nodes: &[MathNode<'_>], check: impl FnOnce(&[MathNode<'_>]) -> R) -> R {
    let bytes = write(nodes);
    let formula = Formula::new();
    let mut parser = MtefParser::new(formula.arena(), &bytes);
    assert!(parser.is_valid(), "writer produced an unparsable header");
    let recovered = parser.parse().expect("parsing should succeed");
    check(&recovered)
}

/// Serialize nodes, parse them back, and return the text they render to
fn roundtrip_text(nodes: &[MathNode<'_>]) -> String {
    roundtrip(nodes, flatten_text)
}

/// Concatenate every piece of text a subtree renders
///
/// The reader lowers each character to its own `Text` node, so a subtree that
/// started life as `Number("12")` comes back as two nodes.
fn flatten_text(nodes: &[MathNode<'_>]) -> String {
    let mut out = String::new();
    for node in nodes {
        match node {
            MathNode::Text(text) => out.push_str(text),
            MathNode::Number(number) => out.push_str(number),
            MathNode::Row(content) => out.push_str(&flatten_text(content)),
            MathNode::Run { content, .. } => out.push_str(&flatten_text(content)),
            other => out.push_str(&format!("<{}>", node_kind(other))),
        }
    }
    out
}

/// Short name of a node variant, for readable assertion failures
fn node_kind(node: &MathNode<'_>) -> &'static str {
    match node {
        MathNode::Frac { .. } => "frac",
        MathNode::Root { .. } => "root",
        MathNode::Power { .. } => "power",
        MathNode::Sub { .. } => "sub",
        MathNode::SubSup { .. } => "subsup",
        MathNode::Under { .. } => "under",
        MathNode::Over { .. } => "over",
        MathNode::UnderOver { .. } => "underover",
        MathNode::Fenced { .. } => "fenced",
        MathNode::LargeOp { .. } => "largeop",
        MathNode::Matrix { .. } => "matrix",
        MathNode::GroupChar { .. } => "groupchar",
        _ => "other",
    }
}

/// Assert that a round-trip produced exactly one node and return it
fn single<'n, 'a>(nodes: &'n [MathNode<'a>]) -> &'n MathNode<'a> {
    assert_eq!(nodes.len(), 1, "expected one node, got {nodes:?}");
    &nodes[0]
}

/// Text leaf helper
fn text(value: &'static str) -> MathNode<'static> {
    MathNode::Text(Cow::Borrowed(value))
}

/// Number leaf helper
fn number(value: &'static str) -> MathNode<'static> {
    MathNode::Number(Cow::Borrowed(value))
}

// ---------------------------------------------------------------------------
// Stream shape
// ---------------------------------------------------------------------------

#[test]
fn stream_starts_with_an_ole_header_describing_the_payload() {
    let bytes = write(&[text("x")]);

    assert_eq!(
        u16::from_le_bytes([bytes[0], bytes[1]]),
        OLE_HEADER_CB_HDR,
        "header length"
    );
    assert_eq!(
        u32::from_le_bytes([bytes[2], bytes[3], bytes[4], bytes[5]]),
        OLE_HEADER_VERSION,
        "header version"
    );
    assert_eq!(
        u32::from_le_bytes([bytes[8], bytes[9], bytes[10], bytes[11]]) as usize,
        bytes.len() - OLE_HEADER_LEN,
        "payload length"
    );
}

#[test]
fn mtef_header_announces_version_five() {
    let bytes = write(&[text("x")]);
    let header = &bytes[OLE_HEADER_LEN..];

    assert_eq!(header[0], MTEF_VERSION_5);
    assert_eq!(header[1], PLATFORM_WINDOWS);
    assert_eq!(header[2], PRODUCT_MATHTYPE);
    assert_eq!(header[5], 0, "empty application key");
    assert_eq!(header[6], EQUATION_DISPLAY);
}

#[test]
fn header_can_be_omitted_for_bare_streams() {
    let bytes = write_bare(&[text("x")]);

    assert_eq!(
        bytes[0], MTEF_VERSION_5,
        "stream starts with the MTEF header"
    );
    assert_eq!(
        bytes.len() + OLE_HEADER_LEN,
        write(&[text("x")]).len(),
        "the two forms differ by exactly the OLE header"
    );
}

#[test]
fn inline_equations_set_the_inline_flag() {
    let writer = MtefWriter::with_options(MtefWriteOptions::default().with_inline(true));
    let bytes = writer.write_nodes(&[text("x")]).expect("serialization");
    assert_eq!(bytes[OLE_HEADER_LEN + 6], EQUATION_INLINE);
}

#[test]
fn a_single_character_encodes_the_expected_records() {
    let bytes = write_bare(&[text("x")]);

    // MTEF header (7 bytes), FULL, LINE + options, CHAR record, END, END
    assert_eq!(
        &bytes[7..],
        &[FULL, LINE, 0, CHAR, 0, TYPEFACE_VARIABLE, b'x', 0, END, END]
    );
}

#[test]
fn writing_into_a_shared_buffer_appends() {
    let writer = MtefWriter::new();
    let mut buffer = vec![0xAA, 0xBB];
    writer
        .write_nodes_into(&[text("x")], &mut buffer)
        .expect("serialization");

    assert_eq!(&buffer[..2], &[0xAA, 0xBB], "existing bytes are preserved");
    assert_eq!(
        u32::from_le_bytes([buffer[10], buffer[11], buffer[12], buffer[13]]) as usize,
        buffer.len() - 2 - OLE_HEADER_LEN,
        "the header describes only the appended payload"
    );
}

#[test]
fn an_empty_equation_still_produces_a_valid_stream() {
    roundtrip(&[], |recovered| assert!(recovered.is_empty()));
}

#[test]
fn a_formula_can_be_serialized_directly() {
    let mut formula = Formula::new();
    formula.set_root(vec![text("x")]);

    let bytes = MtefWriter::new().write(&formula).expect("serialization");
    assert_eq!(bytes, write(&[text("x")]));
}

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

/// Decode a whitespace-separated hex fixture, ignoring comment lines
fn decode_hex_fixture(fixture: &str) -> Vec<u8> {
    fixture
        .lines()
        .filter(|line| !line.trim_start().starts_with('#'))
        .flat_map(str::split_whitespace)
        .map(|byte| u8::from_str_radix(byte, 16).expect("fixture holds hex bytes"))
        .collect()
}

/// Prefix an MTEF payload with the OLE equation header its container supplies
fn with_ole_header(payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(OLE_HEADER_LEN + payload.len());
    bytes.extend_from_slice(&OLE_HEADER_CB_HDR.to_le_bytes());
    bytes.extend_from_slice(&OLE_HEADER_VERSION.to_le_bytes());
    bytes.extend_from_slice(&OLE_CLIPBOARD_FORMAT.to_le_bytes());
    bytes.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    bytes.resize(OLE_HEADER_LEN, 0);
    bytes.extend_from_slice(payload);
    bytes
}

#[test]
fn the_mtef3_fixture_still_parses() {
    // Guards the MTEF 1-4 framing, where the record tag doubles as the attribute
    // byte, against the MTEF 5 framing the writer emits.
    let fixture = include_str!("../../../../../../test-data/ole/doc/mtef/equation3-minimal.hex");
    let bytes = with_ole_header(&decode_hex_fixture(fixture));

    let formula = Formula::new();
    let mut parser = MtefParser::new(formula.arena(), &bytes);
    assert!(parser.is_valid(), "fixture header is accepted");
    assert_eq!(parser.version_info().map(|info| info.0), Some(3));

    let nodes = parser.parse().expect("fixture parses");
    assert_eq!(flatten_text(&nodes), "=");
}
