//! Marker-bearing PPTX and DOCX corpora, and their marker-stripped controls.
//!
//! Change [0649] found that every prior PPTX record in this program measured
//! on generated corpora whose members never mention the markup-compatibility
//! (MCE) namespace.  `litchi_ooxml_common::mce` takes a cheap borrowing path
//! on such a part -- a namespace scan and nothing else -- and a whole-part
//! re-serializing path on a part that mentions it.  On the real 103-member
//! deck 0649 measured, the rewriting path produced **16.25x its input** and
//! cost **93.9%** of one opened-transaction shape-text edit; on the generated
//! corpus it never ran at all.  Change [0601] gave this harness producer-shaped
//! corpora, but only its XLSX family carries markers across the package: its
//! PPTX shape marks the slide parts alone and its DOCX shape marks
//! `word/document.xml` alone.
//!
//! This module adds the missing corpora.  Two families, each in two variants:
//!
//! * **marker** -- the package a production writer emitted, with the real
//!   producer's root namespace declarations (and, for DOCX, its `mc:Ignorable`
//!   token list and `w14` paragraph attributes) merged into every part kind
//!   that carries them in the fixture, plus the `mc:AlternateContent` wrapper
//!   PowerPoint emits on a slide;
//! * **control** -- the same archive with every occurrence of the MCE
//!   namespace URI replaced by an inert URI of **exactly the same length**.
//!   Member names, per-member uncompressed byte counts, element counts and
//!   attribute counts are identical to the marker variant by construction and
//!   proved at build time; the only thing that differs is which branch of
//!   `process_markup_compatibility` each part takes.  Change 0588 introduced
//!   this technique and 0649 used it on the real deck.
//!
//! The shape is not invented here.  It is **derived from real fixtures that
//! are ordinary tracked files in this repository** by
//! `docs/performance/results/change-0664/scripts/derive_marker_shape.py`,
//! which censuses every member of
//! `test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx`
//! (0649's deck),
//! `test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/layout-in-cell-2.docx`
//! (real Word output: 95.0% of 463,565 uncompressed bytes in marker-bearing
//! parts) and
//! `test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf89064.pptx` (the only
//! PPTX fixture in the corpus whose marker coverage reaches the notes parts,
//! because 0649's deck has no notes slides at all).  The script's `verify`
//! mode re-derives the declaration sets and fails if this file drifts from
//! them.
//!
//! Generation is a pure function of the family and variant: no clock, no PRNG,
//! no ambient state, no file read.  Nothing here changes an existing selector,
//! an existing corpus identity or the default matrix; every case that uses
//! these corpora is opt-in and absent from [`crate::Case::DEFAULT`].
//!
//! [0649]: ../../../docs/performance/0649-pptx-opened-transaction-real-deck-edit.md
//! [0601]: ../../../docs/performance/0601-perf-harness-real-producer-shape.md

use std::{error::Error, sync::Arc, time::Instant};

use litchi_core::OwnedSource;
use serde::Serialize;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

use crate::{
    Case, CaseResult, Corpus, CorpusManifest, HashingDiscardSink, SemanticShape, SinkSummary,
    iteration_count, record_elapsed, semantic_docx_bytes, semantic_pptx_text, sha256_hex,
    statistics,
};

/// Generator identity of the marker-bearing PPTX family.
pub(crate) const PPTX_MARKER_SHAPE_GENERATOR: &str = "litchi-pptx-marker-shape-v1";
/// Generator identity of the marker-bearing DOCX family.
pub(crate) const DOCX_MARKER_SHAPE_GENERATOR: &str = "litchi-docx-marker-shape-v1";

/// The markup-compatibility namespace.  Both library gates that matter here --
/// the MCE codec's presence scan and `source_stream_eligible` -- key on this
/// string occurring anywhere in a part's bytes.
const MCE_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// The control's replacement URI.  It must be **exactly as long** as
/// [`MCE_NAMESPACE`] so that the control's parts are byte-for-byte the same
/// length as the marker variant's, and it must be a URI no OOXML codec
/// recognizes.  `.invalid` is reserved by RFC 2606 and can never resolve.
const CONTROL_NAMESPACE: &str = "http://litchi.invalid/perf-baseline/marker-stripped/0664/xx";

/// The fixtures the shape below was derived from.  They are named here so the
/// derivation script can check that this file still agrees with them, and so a
/// reader can re-run the census that produced these constants.
const PPTX_DERIVATION_FIXTURE: &str =
    "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx";
const DOCX_DERIVATION_FIXTURE: &str =
    "test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/layout-in-cell-2.docx";
const PPTX_NOTES_DERIVATION_FIXTURE: &str =
    "test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf89064.pptx";

/// The derivation script that produced every declaration list below.
const DERIVATION_SCRIPT: &str =
    "docs/performance/results/change-0664/scripts/derive_marker_shape.py";

// ---------------------------------------------------------------------------
// The derived shape
// ---------------------------------------------------------------------------

/// Root namespace declarations every marker-bearing PPTX part carries in the
/// fixture, beyond the three (`a`, `p`, `r`) the production writer already
/// emits.  The fixture declares six bindings on the root of every slide,
/// layout, master, notes slide, notes master and `presentation.xml`; the MCE
/// codec re-declares **every in-scope binding on every emitted start tag**, so
/// this count is what sets the output amplification (0649 measured 6.00
/// bindings per emitted element and 16.28x output).
const PPTX_ADDED_DECLARATIONS: [(&str, &str); 3] = [
    (
        "p14",
        "http://schemas.microsoft.com/office/powerpoint/2010/main",
    ),
    (
        "p15",
        "http://schemas.microsoft.com/office/powerpoint/2012/main",
    ),
    ("mc", MCE_NAMESPACE),
];

/// Bindings the fixture declares on a marker-bearing PPTX part's root.
const PPTX_ROOT_DECLARATION_COUNT: usize = 6;

/// The `mc:AlternateContent` wrapper the fixture emits once per slide: a
/// `p14`-qualified transition duration with an unqualified fallback.  It is
/// byte-identical across the whole LibreOffice/Collabora PPTX family and is
/// what makes the slide parts carry a resolvable alternate-content branch and
/// not merely a namespace declaration.
const PPTX_ALTERNATE_CONTENT: &str = concat!(
    "<mc:AlternateContent>",
    "<mc:Choice Requires=\"p14\">",
    "<p:transition spd=\"slow\" p14:dur=\"2000\"></p:transition>",
    "</mc:Choice>",
    "<mc:Fallback>",
    "<p:transition spd=\"slow\"></p:transition>",
    "</mc:Fallback>",
    "</mc:AlternateContent>",
);

/// The 32 root declarations Word writes on `word/document.xml`,
/// `word/numbering.xml`, `word/footnotes.xml`, `word/endnotes.xml` and every
/// header/footer part in the derivation fixture.
const DOCX_FULL_DECLARATIONS: [(&str, &str); 32] = [
    (
        "wpc",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    ),
    (
        "cx",
        "http://schemas.microsoft.com/office/drawing/2014/chartex",
    ),
    (
        "cx1",
        "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
    ),
    (
        "cx2",
        "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
    ),
    (
        "cx3",
        "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
    ),
    (
        "cx4",
        "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
    ),
    (
        "cx5",
        "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
    ),
    (
        "cx6",
        "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
    ),
    (
        "cx7",
        "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
    ),
    (
        "cx8",
        "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
    ),
    ("mc", MCE_NAMESPACE),
    (
        "aink",
        "http://schemas.microsoft.com/office/drawing/2016/ink",
    ),
    (
        "am3d",
        "http://schemas.microsoft.com/office/drawing/2017/model3d",
    ),
    ("o", "urn:schemas-microsoft-com:office:office"),
    (
        "r",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    ),
    (
        "m",
        "http://schemas.openxmlformats.org/officeDocument/2006/math",
    ),
    ("v", "urn:schemas-microsoft-com:vml"),
    (
        "wp14",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    ),
    (
        "wp",
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    ),
    ("w10", "urn:schemas-microsoft-com:office:word"),
    (
        "w",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    ),
    (
        "w14",
        "http://schemas.microsoft.com/office/word/2010/wordml",
    ),
    (
        "w15",
        "http://schemas.microsoft.com/office/word/2012/wordml",
    ),
    (
        "w16cex",
        "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    ),
    (
        "w16cid",
        "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    ),
    (
        "w16",
        "http://schemas.microsoft.com/office/word/2018/wordml",
    ),
    (
        "w16sdtdh",
        "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    ),
    (
        "w16se",
        "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    ),
    (
        "wpg",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    ),
    (
        "wpi",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    ),
    (
        "wne",
        "http://schemas.microsoft.com/office/word/2006/wordml",
    ),
    (
        "wps",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    ),
];

/// The 15 root declarations Word writes on `word/settings.xml`.
const DOCX_SETTINGS_DECLARATIONS: [(&str, &str); 15] = [
    ("mc", MCE_NAMESPACE),
    ("o", "urn:schemas-microsoft-com:office:office"),
    (
        "r",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    ),
    (
        "m",
        "http://schemas.openxmlformats.org/officeDocument/2006/math",
    ),
    ("v", "urn:schemas-microsoft-com:vml"),
    ("w10", "urn:schemas-microsoft-com:office:word"),
    (
        "w",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    ),
    (
        "w14",
        "http://schemas.microsoft.com/office/word/2010/wordml",
    ),
    (
        "w15",
        "http://schemas.microsoft.com/office/word/2012/wordml",
    ),
    (
        "w16cex",
        "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    ),
    (
        "w16cid",
        "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    ),
    (
        "w16",
        "http://schemas.microsoft.com/office/word/2018/wordml",
    ),
    (
        "w16sdtdh",
        "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    ),
    (
        "w16se",
        "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    ),
    (
        "sl",
        "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    ),
];

/// The 10 root declarations Word writes on `word/styles.xml`,
/// `word/fontTable.xml` and `word/webSettings.xml`.
const DOCX_STYLES_DECLARATIONS: [(&str, &str); 10] = [
    ("mc", MCE_NAMESPACE),
    (
        "r",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    ),
    (
        "w",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    ),
    (
        "w14",
        "http://schemas.microsoft.com/office/word/2010/wordml",
    ),
    (
        "w15",
        "http://schemas.microsoft.com/office/word/2012/wordml",
    ),
    (
        "w16cex",
        "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    ),
    (
        "w16cid",
        "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    ),
    (
        "w16",
        "http://schemas.microsoft.com/office/word/2018/wordml",
    ),
    (
        "w16sdtdh",
        "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    ),
    (
        "w16se",
        "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    ),
];

/// `mc:Ignorable` as the fixture writes it on the 32-declaration parts.
const DOCX_FULL_IGNORABLE: &str = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14";
/// `mc:Ignorable` as the fixture writes it on the 15- and 10-declaration parts.
/// It drops `wp14`, which those roots do not declare.
const DOCX_SHORT_IGNORABLE: &str = "w14 w15 w16se w16cid w16 w16cex w16sdtdh";

/// Slides in the PPTX marker shape.  0649's deck has 13.
const PPTX_MARKER_SLIDES: usize = 13;
/// Text boxes per slide.  Chosen so the shape's **slide** byte total lands near
/// the real deck's, because that is what the cost is a function of: change
/// 0649 found that all three per-slide sites of one capture read every slide in
/// full and that the capture never reads a layout or a master at all, so the
/// deck's 291,551 layout bytes and 178,069 master bytes are not on the measured
/// path.  Thirteen slides of 48 text boxes give 247,741 slide bytes against the
/// fixture's 269,178 (92.0%), where the generated corpus every prior PPTX
/// record used has 40,788 (15.2%).  Matching the fixture's *package* total
/// instead would mean 130 text boxes a slide, which triples the edit for no
/// extra signal: this harness's writer emits eleven small layouts and one
/// master where the fixture has eighteen large layouts and eleven masters, and
/// no slide/layout ratio can reproduce both.
const PPTX_MARKER_TEXT_BOXES: usize = 48;

/// Paragraph count of the DOCX marker shape.  `SemanticShape::Medium` is the
/// shape the existing `docx_ordinary_save_*` selectors measure, so the marker
/// corpus is byte-comparable with the generated one it is read against.
const DOCX_MARKER_SHAPE: SemanticShape = SemanticShape::Medium;

// ---------------------------------------------------------------------------
// Family, variant and census
// ---------------------------------------------------------------------------

/// Which marker-bearing family a corpus belongs to.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Family {
    Docx,
    Pptx,
}

impl Family {
    pub(crate) const ALL: [Self; 2] = [Self::Docx, Self::Pptx];

    const fn generator(self) -> &'static str {
        match self {
            Self::Docx => DOCX_MARKER_SHAPE_GENERATOR,
            Self::Pptx => PPTX_MARKER_SHAPE_GENERATOR,
        }
    }

    const fn package_format(self) -> &'static str {
        match self {
            Self::Docx => "DOCX/OPC/ZIP",
            Self::Pptx => "PPTX/OPC/ZIP",
        }
    }

    const fn fixture(self) -> &'static str {
        match self {
            Self::Docx => DOCX_DERIVATION_FIXTURE,
            Self::Pptx => PPTX_DERIVATION_FIXTURE,
        }
    }
}

/// Which side of the pair a corpus is.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Variant {
    /// Carries the markup-compatibility namespace: the codec rewrites.
    Marker,
    /// The same bytes with the namespace URI replaced by an inert URI of the
    /// same length: the codec borrows.
    Control,
}

impl Variant {
    pub(crate) const ALL: [Self; 2] = [Self::Marker, Self::Control];

    const fn suffix(self) -> &'static str {
        match self {
            Self::Marker => "marker",
            Self::Control => "control",
        }
    }
}

/// The marker census of one archive member.
#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
pub(crate) struct MemberCensus {
    pub(crate) part: String,
    pub(crate) kind: &'static str,
    pub(crate) uncompressed_bytes: usize,
    pub(crate) sha256: String,
    /// Whether the part mentions the markup-compatibility namespace anywhere,
    /// which is exactly the condition the codec's presence scan tests.
    pub(crate) markup_compatibility_namespace: bool,
    pub(crate) mc_ignorable: Option<String>,
    pub(crate) alternate_content_occurrences: usize,
    /// Namespace bindings declared on the part's root element.  The codec
    /// re-declares every in-scope binding on every emitted start tag.
    pub(crate) root_declarations: usize,
    pub(crate) start_tags: usize,
    pub(crate) attributes: usize,
}

/// Everything a marker-shape corpus states about itself beyond the ordinary
/// [`CorpusManifest`].
#[derive(Clone, Debug, Serialize)]
pub(crate) struct MarkerEvidence {
    pub(crate) generator: &'static str,
    pub(crate) family: &'static str,
    pub(crate) variant: &'static str,
    pub(crate) shape: String,
    /// Where the shape came from, so a record does not have to assert it.
    pub(crate) derived_from_fixture: &'static str,
    /// 0649's deck has no notes slides at all, so the notes part kinds' root
    /// declaration set is derived from a second fixture rather than assumed.
    pub(crate) derived_notes_from_fixture: Option<&'static str>,
    pub(crate) derivation_script: &'static str,
    pub(crate) markup_compatibility_namespace: &'static str,
    pub(crate) control_namespace: &'static str,
    pub(crate) member_count: usize,
    pub(crate) uncompressed_bytes: usize,
    pub(crate) marked_member_count: usize,
    pub(crate) marked_bytes: usize,
    /// Share of uncompressed bytes in parts the codec would rewrite, in basis
    /// points, so the census stays an exact integer.
    pub(crate) marked_byte_share_basis_points: usize,
    pub(crate) marked_part_kinds: Vec<&'static str>,
    /// What the production writer's own package already carries before this
    /// module marks anything.  Change 0032 recorded that the generated XLSX
    /// worksheets are marker free; the DOCX writer is not, so the baseline is
    /// stated rather than assumed.
    pub(crate) skeleton_member_count: usize,
    pub(crate) skeleton_uncompressed_bytes: usize,
    pub(crate) skeleton_marked_member_count: usize,
    pub(crate) skeleton_marked_bytes: usize,
    pub(crate) alternate_content_blocks: usize,
    pub(crate) parts: Vec<MemberCensus>,
    /// The scenario oracles, derived from the built package rather than from
    /// the generator's intent: the eager full text, the source-backed full
    /// text, and the sequential sink projection with its own byte and object
    /// counts.  A marker corpus and its control must agree on all three.
    pub(crate) text_sha256: String,
    pub(crate) text_bytes: usize,
    pub(crate) text_objects: usize,
    pub(crate) source_text_sha256: String,
    pub(crate) source_text_bytes: usize,
    /// The sequential sink projection, when the documented sink entry point
    /// admits this corpus.  A producer-shaped package can be *refused* by a
    /// sink parser that a marker-free one is not, because the codec re-declares
    /// every in-scope binding on every emitted start tag and the parser counts
    /// those bindings against a fixed limit.  A refusal is an outcome: the
    /// message is frozen here rather than the scenario being dropped.
    pub(crate) sink_sha256: Option<String>,
    pub(crate) sink_bytes: Option<u64>,
    pub(crate) sink_objects: Option<u64>,
    pub(crate) sink_refusal: Option<String>,
}

/// A marker-shape corpus: an ordinary [`Corpus`] plus its evidence block.
#[derive(Debug)]
pub(crate) struct MarkerCorpus {
    pub(crate) corpus: Corpus,
    pub(crate) evidence: MarkerEvidence,
}

// ---------------------------------------------------------------------------
// XML rewriting helpers
// ---------------------------------------------------------------------------

fn part_kind(name: &str) -> &'static str {
    // Relationship parts live under the same prefixes as the parts they
    // describe and are not PresentationML or WordprocessingML, so they are
    // never marked.
    if !name.ends_with(".xml") || name.contains("/_rels/") {
        return "other";
    }
    if name.starts_with("ppt/slides/") {
        "pptx-slide"
    } else if name.starts_with("ppt/slideLayouts/") {
        "pptx-slide-layout"
    } else if name.starts_with("ppt/slideMasters/") {
        "pptx-slide-master"
    } else if name.starts_with("ppt/notesSlides/") {
        "pptx-notes-slide"
    } else if name.starts_with("ppt/notesMasters/") {
        "pptx-notes-master"
    } else if name == "ppt/presentation.xml" {
        "pptx-presentation"
    } else if name.starts_with("ppt/theme/") {
        "pptx-theme"
    } else if name == "word/document.xml" {
        "docx-document"
    } else if name == "word/settings.xml" {
        "docx-settings"
    } else if name == "word/styles.xml" {
        "docx-styles"
    } else if name == "word/numbering.xml" {
        "docx-numbering"
    } else if name == "word/fontTable.xml" {
        "docx-font-table"
    } else if name == "word/webSettings.xml" {
        "docx-web-settings"
    } else if name == "word/footnotes.xml" {
        "docx-footnotes"
    } else if name == "word/endnotes.xml" {
        "docx-endnotes"
    } else if name.starts_with("word/header") || name.starts_with("word/footer") {
        "docx-header-footer"
    } else {
        "other"
    }
}

/// Locate the root start tag of an XML part, skipping the XML declaration.
fn root_span(xml: &str) -> Result<(usize, usize), Box<dyn Error>> {
    let mut cursor = 0;
    if let Some(position) = xml.find("?>") {
        cursor = position + "?>".len();
    }
    let open = xml[cursor..]
        .find('<')
        .ok_or("marker shape rewrite found no root element")?
        + cursor;
    let close = xml[open..]
        .find('>')
        .ok_or("marker shape root start tag is unterminated")?
        + open;
    Ok((open, close))
}

/// Merge namespace declarations (and optionally `mc:Ignorable`) into a part's
/// root start tag, skipping any prefix the writer already declares so the
/// result never carries a duplicate attribute.
fn merge_root(
    xml: &str,
    declarations: &[(&str, &str)],
    ignorable: Option<&str>,
) -> Result<String, Box<dyn Error>> {
    let (open, close) = root_span(xml)?;
    let root = &xml[open..=close];
    if root.contains("mc:Ignorable") || root.contains(MCE_NAMESPACE) {
        return Err("marker shape root already carries markup-compatibility markup".into());
    }
    let mut added = String::new();
    for (prefix, uri) in declarations {
        let needle = format!("xmlns:{prefix}=\"");
        if root.contains(needle.as_str()) {
            continue;
        }
        added.push_str(&format!(" xmlns:{prefix}=\"{uri}\""));
    }
    if let Some(tokens) = ignorable {
        added.push_str(&format!(" mc:Ignorable=\"{tokens}\""));
    }
    if added.is_empty() {
        return Err("marker shape rewrite added no declaration to a root".into());
    }
    let mut rebuilt = String::with_capacity(xml.len() + added.len());
    rebuilt.push_str(&xml[..close]);
    rebuilt.push_str(&added);
    rebuilt.push_str(&xml[close..]);
    Ok(rebuilt)
}

/// Insert a block immediately before the last occurrence of a closing tag.
fn insert_before_last(xml: &str, close_tag: &str, block: &str) -> Result<String, Box<dyn Error>> {
    let position = xml
        .rfind(close_tag)
        .ok_or("marker shape rewrite found no closing tag to insert before")?;
    let mut rebuilt = String::with_capacity(xml.len() + block.len());
    rebuilt.push_str(&xml[..position]);
    rebuilt.push_str(block);
    rebuilt.push_str(&xml[position..]);
    Ok(rebuilt)
}

/// Give every `<w:p>` the revision and paragraph-identity attributes Word
/// writes.  They are `w14`-qualified, so a conforming consumer that honours
/// `mc:Ignorable` drops them and the projected text is unchanged.
fn mark_paragraphs(xml: &str) -> (String, usize) {
    let mut rebuilt = String::with_capacity(xml.len());
    let mut rest = xml;
    let mut edits = 0usize;
    while let Some(position) = rest.find("<w:p>") {
        rebuilt.push_str(&rest[..position]);
        rebuilt.push_str(&format!(
            r#"<w:p w14:paraId="{:08X}" w14:textId="{:08X}" w:rsidR="00AB12CD" w:rsidRDefault="00AB12CD">"#,
            edits + 1,
            0x0664_0000_u32.wrapping_add(edits as u32)
        ));
        rest = &rest[position + "<w:p>".len()..];
        edits += 1;
    }
    rebuilt.push_str(rest);
    (rebuilt, edits)
}

fn count_occurrences(haystack: &[u8], needle: &[u8]) -> usize {
    if needle.is_empty() || needle.len() > haystack.len() {
        return 0;
    }
    haystack
        .windows(needle.len())
        .filter(|window| *window == needle)
        .count()
}

fn contains(haystack: &[u8], needle: &[u8]) -> bool {
    !needle.is_empty() && needle.len() <= haystack.len() && count_occurrences(haystack, needle) > 0
}

/// Count start tags: every `<` followed by a name character.  Comments,
/// declarations, closing tags and processing instructions are excluded.
fn start_tags(bytes: &[u8]) -> usize {
    bytes
        .windows(2)
        .filter(|window| {
            window[0] == b'<' && (window[1].is_ascii_alphabetic() || window[1] == b'_')
        })
        .count()
}

fn extract_attribute(xml: &str, prefix: &str) -> Option<String> {
    let start = xml.find(prefix)? + prefix.len();
    let rest = xml.get(start..)?;
    let end = rest.find('"')?;
    rest.get(..end).map(str::to_owned)
}

fn census_of(part: &str, bytes: &[u8]) -> Result<MemberCensus, Box<dyn Error>> {
    let text = std::str::from_utf8(bytes).ok();
    let root_declarations = match text {
        Some(xml) if part.ends_with(".xml") || part.ends_with(".rels") => {
            let (open, close) = root_span(xml)?;
            xml[open..=close].matches("xmlns:").count()
        },
        _ => 0,
    };
    Ok(MemberCensus {
        part: part.to_owned(),
        kind: part_kind(part),
        uncompressed_bytes: bytes.len(),
        sha256: sha256_hex(bytes),
        markup_compatibility_namespace: contains(bytes, MCE_NAMESPACE.as_bytes()),
        mc_ignorable: text.and_then(|xml| extract_attribute(xml, "mc:Ignorable=\"")),
        alternate_content_occurrences: count_occurrences(bytes, b"<mc:AlternateContent"),
        root_declarations,
        start_tags: start_tags(bytes),
        attributes: count_occurrences(bytes, b"=\""),
    })
}

// ---------------------------------------------------------------------------
// Corpus construction
// ---------------------------------------------------------------------------

/// Census every member of an archive.
fn archive_census(archive: &[u8]) -> Result<Vec<MemberCensus>, Box<dyn Error>> {
    let reader = ArchiveReader::new(archive)?;
    let names = reader.file_names().map(str::to_owned).collect::<Vec<_>>();
    let mut parts = Vec::with_capacity(names.len());
    for name in &names {
        parts.push(census_of(name.as_str(), &reader.read(name.as_str())?)?);
    }
    Ok(parts)
}

/// Author the PPTX skeleton with the production writer, then mark it.
///
/// Returns the writer's own package and the marked one, so a corpus can state
/// what the production writer already emits.
fn pptx_marker_archive() -> Result<(Vec<u8>, Vec<u8>), Box<dyn Error>> {
    let mut package = litchi_pptx::Package::new()?;
    let presentation = package.presentation_mut()?;
    for slide_index in 0..PPTX_MARKER_SLIDES {
        let slide = presentation.add_slide()?;
        for shape_index in 0..PPTX_MARKER_TEXT_BOXES {
            slide.add_text_box(
                &semantic_pptx_text(slide_index, shape_index, false),
                36 + i64::try_from(shape_index % 4)? * 180,
                36 + i64::try_from(shape_index / 4)? * 90,
                144,
                54,
            );
        }
    }
    let skeleton = package.to_bytes()?;

    let reader = ArchiveReader::new(&skeleton)?;
    let names = reader.file_names().map(str::to_owned).collect::<Vec<_>>();
    let mut writer = StreamingArchiveWriter::new();
    let mut marked = 0usize;
    let mut wrapped = 0usize;
    let mut declarations = 0usize;
    for name in &names {
        let bytes = reader.read(name.as_str())?;
        let marks = matches!(
            part_kind(name.as_str()),
            "pptx-slide"
                | "pptx-slide-layout"
                | "pptx-slide-master"
                | "pptx-notes-slide"
                | "pptx-notes-master"
                | "pptx-presentation"
        );
        if !marks {
            writer.write_deflated(name, &bytes)?;
            continue;
        }
        let xml = std::str::from_utf8(&bytes)?;
        let mut patched = merge_root(xml, &PPTX_ADDED_DECLARATIONS, None)?;
        if part_kind(name.as_str()) == "pptx-slide" {
            patched = insert_before_last(&patched, "</p:sld>", PPTX_ALTERNATE_CONTENT)?;
            wrapped += 1;
        }
        let (open, close) = root_span(&patched)?;
        declarations = patched[open..=close].matches("xmlns:").count();
        writer.write_deflated(name, patched.as_bytes())?;
        marked += 1;
    }
    if wrapped != PPTX_MARKER_SLIDES {
        return Err("marker PPTX rewrite did not wrap every slide".into());
    }
    if declarations != PPTX_ROOT_DECLARATION_COUNT {
        return Err(
            "marker PPTX rewrite produced a root declaration count the fixture does not have"
                .into(),
        );
    }
    if marked < wrapped {
        return Err("marker PPTX rewrite marked fewer parts than it wrapped".into());
    }
    Ok((skeleton, writer.finish_to_bytes()?))
}

/// Author the DOCX skeleton with the production writer, then mark it.
fn docx_marker_archive() -> Result<(Vec<u8>, Vec<u8>), Box<dyn Error>> {
    let skeleton = semantic_docx_bytes(DOCX_MARKER_SHAPE)?;
    let reader = ArchiveReader::new(&skeleton)?;
    let names = reader.file_names().map(str::to_owned).collect::<Vec<_>>();
    let mut writer = StreamingArchiveWriter::new();
    let mut marked = 0usize;
    let mut paragraph_edits = 0usize;
    for name in &names {
        let bytes = reader.read(name.as_str())?;
        let kind = part_kind(name.as_str());
        let plan: Option<(&[(&str, &str)], &str)> = match kind {
            "docx-document" | "docx-numbering" | "docx-footnotes" | "docx-endnotes"
            | "docx-header-footer" => Some((&DOCX_FULL_DECLARATIONS, DOCX_FULL_IGNORABLE)),
            "docx-settings" => Some((&DOCX_SETTINGS_DECLARATIONS, DOCX_SHORT_IGNORABLE)),
            "docx-styles" | "docx-font-table" | "docx-web-settings" => {
                Some((&DOCX_STYLES_DECLARATIONS, DOCX_SHORT_IGNORABLE))
            },
            _ => None,
        };
        let Some((declarations, ignorable)) = plan else {
            writer.write_deflated(name, &bytes)?;
            continue;
        };
        let xml = std::str::from_utf8(&bytes)?;
        if xml.contains(MCE_NAMESPACE) {
            // The production DOCX writer already declares `xmlns:mc` on
            // `settings.xml`, `numbering.xml` and `fontTable.xml`, so those
            // parts are already on the codec's rewriting path. Leave them
            // exactly as the writer emitted them and count them as marked:
            // the corpus states what the package contains, not what the
            // generator intended.
            writer.write_deflated(name, &bytes)?;
            marked += 1;
            continue;
        }
        let mut patched = merge_root(xml, declarations, Some(ignorable))?;
        if kind == "docx-document" {
            let (rebuilt, edits) = mark_paragraphs(&patched);
            patched = rebuilt;
            paragraph_edits = edits;
        }
        writer.write_deflated(name, patched.as_bytes())?;
        marked += 1;
    }
    if paragraph_edits == 0 {
        return Err("marker DOCX rewrite marked no paragraph".into());
    }
    if marked < 2 {
        return Err("marker DOCX rewrite marked fewer than two parts".into());
    }
    Ok((skeleton, writer.finish_to_bytes()?))
}

/// Rewrite an archive member by member, replacing the markup-compatibility
/// namespace with the equal-length inert one.  Nothing else changes, so every
/// member keeps its name, its length, its element count and its attribute
/// count.
fn strip_markers(archive: &[u8]) -> Result<Vec<u8>, Box<dyn Error>> {
    if MCE_NAMESPACE.len() != CONTROL_NAMESPACE.len() {
        return Err("the marker-stripped control namespace is not the same length".into());
    }
    let reader = ArchiveReader::new(archive)?;
    let names = reader.file_names().map(str::to_owned).collect::<Vec<_>>();
    let mut writer = StreamingArchiveWriter::new();
    let mut replaced = 0usize;
    for name in &names {
        let bytes = reader.read(name.as_str())?;
        match std::str::from_utf8(&bytes) {
            Ok(text) if text.contains(MCE_NAMESPACE) => {
                let patched = text.replace(MCE_NAMESPACE, CONTROL_NAMESPACE);
                if patched.len() != text.len() {
                    return Err("the marker-stripped control changed a member's length".into());
                }
                writer.write_deflated(name, patched.as_bytes())?;
                replaced += 1;
            },
            _ => writer.write_deflated(name, &bytes)?,
        }
    }
    if replaced == 0 {
        return Err("the marker-stripped control found no marker to strip".into());
    }
    Ok(writer.finish_to_bytes()?)
}

/// Build one marker-shape corpus.
pub(crate) fn build(family: Family, variant: Variant) -> Result<MarkerCorpus, Box<dyn Error>> {
    let (skeleton, marker) = match family {
        Family::Docx => docx_marker_archive()?,
        Family::Pptx => pptx_marker_archive()?,
    };
    let skeleton_census = archive_census(&skeleton)?;
    let archive = match variant {
        Variant::Marker => marker,
        Variant::Control => strip_markers(&marker)?,
    };

    // Census every member, and derive the scenario oracle from the package the
    // library actually reads rather than from the generator's intent.
    let parts = archive_census(&archive)?;
    let member_count = parts.len();
    let uncompressed_bytes = parts.iter().map(|part| part.uncompressed_bytes).sum();
    let marked = parts
        .iter()
        .filter(|part| part.markup_compatibility_namespace)
        .collect::<Vec<_>>();
    let marked_bytes: usize = marked.iter().map(|part| part.uncompressed_bytes).sum();
    let mut marked_part_kinds = marked.iter().map(|part| part.kind).collect::<Vec<_>>();
    marked_part_kinds.sort_unstable();
    marked_part_kinds.dedup();
    let marked_member_count = marked.len();
    let alternate_content_blocks = parts
        .iter()
        .map(|part| part.alternate_content_occurrences)
        .sum();

    let oracles = read_back_oracles(family, &archive)?;
    let text = oracles.text.clone();
    let text_objects = oracles.objects;
    let shape = format!("{}-{}", family_shape(family), variant.suffix());

    let manifest = CorpusManifest {
        name: format!("{}-{}", family_corpus_name(family), variant.suffix()),
        generator: family.generator(),
        package_format: family.package_format(),
        shape: family_shape(family),
        payload_kind: match variant {
            Variant::Marker => "marker-bearing-producer-shape",
            Variant::Control => "marker-stripped-control",
        },
        compression: "deflate",
        entry_count: text_objects,
        archive_member_count: member_count,
        entry_bytes: text.len(),
        uncompressed_payload_bytes: uncompressed_bytes,
        archive_bytes: archive.len(),
        archive_sha256: sha256_hex(&archive),
        target_entry: shape.clone(),
        target_payload_bytes: text.len(),
        target_payload_sha256: sha256_hex(text.as_bytes()),
        rtf_variant: None,
        xlsx: None,
    };

    let evidence = MarkerEvidence {
        generator: family.generator(),
        family: match family {
            Family::Docx => "docx",
            Family::Pptx => "pptx",
        },
        variant: variant.suffix(),
        shape,
        derived_from_fixture: family.fixture(),
        derived_notes_from_fixture: match family {
            Family::Pptx => Some(PPTX_NOTES_DERIVATION_FIXTURE),
            Family::Docx => None,
        },
        derivation_script: DERIVATION_SCRIPT,
        markup_compatibility_namespace: MCE_NAMESPACE,
        control_namespace: CONTROL_NAMESPACE,
        member_count,
        uncompressed_bytes,
        marked_member_count,
        marked_bytes,
        marked_byte_share_basis_points: if uncompressed_bytes == 0 {
            0
        } else {
            marked_bytes
                .checked_mul(10_000)
                .ok_or("marked byte share overflows usize")?
                .checked_div(uncompressed_bytes)
                .ok_or("marked byte share division failed")?
        },
        marked_part_kinds,
        skeleton_member_count: skeleton_census.len(),
        skeleton_uncompressed_bytes: skeleton_census
            .iter()
            .map(|part| part.uncompressed_bytes)
            .sum(),
        skeleton_marked_member_count: skeleton_census
            .iter()
            .filter(|part| part.markup_compatibility_namespace)
            .count(),
        skeleton_marked_bytes: skeleton_census
            .iter()
            .filter(|part| part.markup_compatibility_namespace)
            .map(|part| part.uncompressed_bytes)
            .sum(),
        alternate_content_blocks,
        parts,
        text_sha256: sha256_hex(oracles.text.as_bytes()),
        text_bytes: oracles.text.len(),
        text_objects: oracles.objects,
        source_text_sha256: sha256_hex(oracles.source_text.as_bytes()),
        source_text_bytes: oracles.source_text.len(),
        sink_sha256: oracles.sink_sha256,
        sink_bytes: oracles.sink_bytes,
        sink_objects: oracles.sink_objects,
        sink_refusal: oracles.sink_refusal,
    };

    let target_payload = text.into_bytes();
    Ok(MarkerCorpus {
        corpus: Corpus {
            manifest,
            archive,
            target_name: evidence.shape.clone(),
            target_payload,
            xlsx: None,
        },
        evidence,
    })
}

const fn family_shape(family: Family) -> &'static str {
    match family {
        Family::Docx => "marker-medium",
        Family::Pptx => "marker-deck",
    }
}

const fn family_corpus_name(family: Family) -> &'static str {
    match family {
        Family::Docx => "docx-marker-medium",
        Family::Pptx => "pptx-marker-deck",
    }
}

/// The frozen oracles of one corpus, derived by reading the built package
/// back through the documented public entry points.
struct Oracles {
    text: String,
    objects: usize,
    source_text: String,
    sink_sha256: Option<String>,
    sink_bytes: Option<u64>,
    sink_objects: Option<u64>,
    sink_refusal: Option<String>,
}

/// Read the corpus back through the documented readers and freeze what every
/// selector over it must reproduce.  Nothing here is asserted by the
/// generator: the separator behaviour, the object count and the sink's byte
/// count all come from the library.
fn read_back_oracles(family: Family, archive: &[u8]) -> Result<Oracles, Box<dyn Error>> {
    match family {
        Family::Docx => {
            let package =
                litchi_docx::Package::from_reader(std::io::Cursor::new(archive.to_vec()))?;
            let document = package.document()?;
            let text = document.text()?;
            let objects = document.paragraphs()?.len();

            let ceiling = u64::try_from(archive.len().saturating_mul(64).max(1 << 20))?;
            let mut sink = HashingDiscardSink::without_authoring_window(ceiling);
            let options = litchi_core::TextOutputOptions::new("\n", "", ceiling, ceiling);
            let (sink_sha256, sink_bytes, sink_objects, sink_refusal) =
                match document.write_text_to(&mut sink, options) {
                    Ok(report) => {
                        let bytes = report.bytes_written();
                        let objects = report.objects_written();
                        let (_summary, digest) = sink.finish();
                        (Some(digest), Some(bytes), Some(objects), None)
                    },
                    Err(refusal) => (None, None, None, Some(refusal.to_string())),
                };

            let source = litchi_docx::source_backed::Package::from_read_at(Arc::new(
                OwnedSource::new(archive.to_vec()),
            ))?;
            let source_text = source.document()?.extract_text()?;
            Ok(Oracles {
                text,
                objects,
                source_text,
                sink_sha256,
                sink_bytes,
                sink_objects,
                sink_refusal,
            })
        },
        Family::Pptx => {
            let package = litchi_pptx::Package::from_bytes(archive)?;
            let presentation = package.presentation()?;
            let text = presentation.text()?;
            let objects = presentation.slide_count()?;
            let source = litchi_pptx::SourceBackedPresentation::from_read_at(Arc::new(
                OwnedSource::new(archive.to_vec()),
            ))?;
            let mut parts = Vec::with_capacity(source.slide_count());
            for slide in source.slides() {
                parts.push(slide.text()?);
            }
            Ok(Oracles {
                text,
                objects,
                source_text: parts.join("\n"),
                sink_sha256: None,
                sink_bytes: None,
                sink_objects: None,
                sink_refusal: None,
            })
        },
    }
}

/// Prove the control is the same archive with a different branch taken.
///
/// This is the invariant that makes the pair evidence rather than two
/// differently shaped corpora: same members, same per-member lengths, same
/// element counts, same attribute counts, same projected text.  Only the
/// namespace URI, and therefore the codec branch, differs.
pub(crate) fn prove_control_is_byte_comparable(
    marker: &MarkerCorpus,
    control: &MarkerCorpus,
) -> Result<(), Box<dyn Error>> {
    if marker.evidence.member_count != control.evidence.member_count
        || marker.evidence.uncompressed_bytes != control.evidence.uncompressed_bytes
    {
        return Err("the marker-stripped control has a different member or byte count".into());
    }
    for (left, right) in marker.evidence.parts.iter().zip(&control.evidence.parts) {
        if left.part != right.part
            || left.uncompressed_bytes != right.uncompressed_bytes
            || left.start_tags != right.start_tags
            || left.attributes != right.attributes
        {
            return Err(format!(
                "the marker-stripped control differs from {} in shape, not only in namespace",
                left.part
            )
            .into());
        }
    }
    if marker.evidence.text_sha256 != control.evidence.text_sha256 {
        return Err("the marker-stripped control projects different text".into());
    }
    if control.evidence.marked_member_count != 0 {
        return Err("the marker-stripped control still mentions the namespace".into());
    }
    if marker.evidence.marked_member_count == 0 {
        return Err("the marker corpus mentions no namespace".into());
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Scenario runners
// ---------------------------------------------------------------------------

/// Scenario a marker-shape selector measures.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Scenario {
    /// A fresh eager open and the complete projected text.
    EagerFullText,
    /// A fresh source-backed open and the complete projected text.
    SourceFullText,
}

impl Scenario {
    const fn as_str(self) -> &'static str {
        match self {
            Self::EagerFullText => "eager-open-and-full-text",
            Self::SourceFullText => "source-backed-open-and-full-text",
        }
    }
}

pub(crate) fn run_case(
    case: Case,
    scenario: Scenario,
    family: Family,
    corpus: &MarkerCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    match scenario {
        Scenario::EagerFullText => {
            run_full_text(case, family, corpus, warmup_iterations, samples, false)
        },
        Scenario::SourceFullText => {
            run_full_text(case, family, corpus, warmup_iterations, samples, true)
        },
    }
}

fn run_full_text(
    case: Case,
    family: Family,
    corpus: &MarkerCorpus,
    warmup_iterations: usize,
    samples: usize,
    source_backed: bool,
) -> Result<CaseResult, Box<dyn Error>> {
    let expected = if source_backed {
        &corpus.evidence.source_text_sha256
    } else {
        &corpus.evidence.text_sha256
    };
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let text = match (family, source_backed) {
            (Family::Docx, false) => {
                let package = litchi_docx::Package::from_reader(std::io::Cursor::new(
                    corpus.corpus.archive.clone(),
                ))?;
                package.document()?.text()?
            },
            (Family::Docx, true) => {
                let package = litchi_docx::source_backed::Package::from_read_at(Arc::new(
                    OwnedSource::new(corpus.corpus.archive.clone()),
                ))?;
                package.document()?.extract_text()?
            },
            (Family::Pptx, false) => {
                let package = litchi_pptx::Package::from_bytes(&corpus.corpus.archive)?;
                package.presentation()?.text()?
            },
            (Family::Pptx, true) => {
                // `SourceBackedPresentation` has no whole-presentation text
                // entry point, so the documented source-backed full text is
                // every slide's `text()` in order.  The oracle is derived the
                // same way, so the two cannot drift apart.
                let presentation = litchi_pptx::SourceBackedPresentation::from_read_at(Arc::new(
                    OwnedSource::new(corpus.corpus.archive.clone()),
                ))?;
                let mut parts = Vec::with_capacity(presentation.slide_count());
                for slide in presentation.slides() {
                    parts.push(slide.text()?);
                }
                parts.join("\n")
            },
        };
        let duration = started.elapsed();
        if sha256_hex(text.as_bytes()) != *expected {
            return Err("marker-shape full text differs from the corpus oracle".into());
        }
        std::hint::black_box(&text);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let mut result = plain_result(case, corpus, elapsed, None);
    result.output_sha256 = Some(expected.clone());
    Ok(result)
}

fn plain_result(
    case: Case,
    corpus: &MarkerCorpus,
    elapsed: Vec<u64>,
    sink: Option<SinkSummary>,
) -> CaseResult {
    CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink,
        source: None,
        execution: None,
        output_sha256: None,
        operation_metrics: None,
    }
}

/// Write the collected marker census for a retained packet.
pub(crate) fn write_evidence(
    path: &std::path::Path,
    evidence: &[MarkerEvidence],
) -> Result<(), Box<dyn Error>> {
    let document = serde_json::json!({
        "schema": "litchi.perf-baseline.marker-shape-evidence.v1",
        "scenarios": Scenario::EagerFullText.as_str(),
        "corpora": evidence,
    });
    let mut bytes = serde_json::to_vec_pretty(&document)?;
    bytes.push(b'\n');
    std::fs::write(path, bytes)
        .map_err(|source| format!("marker evidence {} is unwritable: {source}", path.display()))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_control_namespace_is_the_same_length_as_the_real_one() {
        assert_eq!(MCE_NAMESPACE.len(), CONTROL_NAMESPACE.len());
        assert!(!CONTROL_NAMESPACE.contains("openxmlformats"));
    }

    #[test]
    fn marker_generation_is_byte_identical_across_runs() {
        for family in Family::ALL {
            for variant in Variant::ALL {
                let first = build(family, variant).expect("first build");
                let second = build(family, variant).expect("second build");
                assert_eq!(
                    first.corpus.manifest.archive_sha256, second.corpus.manifest.archive_sha256,
                    "{:?}/{:?} is not deterministic",
                    family, variant
                );
            }
        }
    }

    #[test]
    fn every_marker_corpus_marks_the_part_kinds_the_fixture_marks() {
        let pptx = build(Family::Pptx, Variant::Marker).expect("PPTX marker corpus builds");
        for kind in [
            "pptx-slide",
            "pptx-slide-layout",
            "pptx-slide-master",
            "pptx-notes-master",
            "pptx-presentation",
        ] {
            assert!(
                pptx.evidence.marked_part_kinds.contains(&kind),
                "{kind} is not marked"
            );
        }
        // The fixture's themes take the borrowed path; so must the corpus's.
        assert!(!pptx.evidence.marked_part_kinds.contains(&"pptx-theme"));
        assert_eq!(pptx.evidence.alternate_content_blocks, PPTX_MARKER_SLIDES);
        for part in &pptx.evidence.parts {
            if part.markup_compatibility_namespace {
                assert_eq!(
                    part.root_declarations, PPTX_ROOT_DECLARATION_COUNT,
                    "{} declares {} bindings",
                    part.part, part.root_declarations
                );
            }
        }

        let docx = build(Family::Docx, Variant::Marker).expect("DOCX marker corpus builds");
        assert!(docx.evidence.marked_part_kinds.contains(&"docx-document"));
        assert!(docx.evidence.marked_member_count >= 2);
        let document = docx
            .evidence
            .parts
            .iter()
            .find(|part| part.part == "word/document.xml")
            .expect("the marker DOCX has a main document part");
        assert_eq!(document.mc_ignorable.as_deref(), Some(DOCX_FULL_IGNORABLE));
        assert!(document.root_declarations >= DOCX_FULL_DECLARATIONS.len());
    }

    #[test]
    fn the_control_is_the_same_archive_with_the_other_branch() {
        for family in Family::ALL {
            let marker = build(family, Variant::Marker).expect("marker corpus builds");
            let control = build(family, Variant::Control).expect("control corpus builds");
            prove_control_is_byte_comparable(&marker, &control)
                .expect("the control is byte-comparable");
            assert_ne!(
                marker.corpus.manifest.archive_sha256,
                control.corpus.manifest.archive_sha256
            );
        }
    }

    #[test]
    fn the_marker_corpora_carry_most_of_their_bytes_in_rewritten_parts() {
        // The derivation fixtures are 93.0% (PPTX) and 95.0% (DOCX) marked by
        // uncompressed bytes. The corpora are built to the same profile; the
        // floor here is what makes the pair a codec measurement rather than a
        // measurement of two unrelated packages.
        for family in Family::ALL {
            let corpus = build(family, Variant::Marker).expect("marker corpus builds");
            assert!(
                corpus.evidence.marked_byte_share_basis_points >= 8_000,
                "{:?} marks only {} basis points of its bytes",
                family,
                corpus.evidence.marked_byte_share_basis_points
            );
        }
    }
}
