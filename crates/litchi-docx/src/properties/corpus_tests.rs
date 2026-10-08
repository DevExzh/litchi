//! Corpus checks for the direct-property readers, the addressable run set
//! (`OfficeCLI`'s `r[N]`) and the source owner's main-document
//! relationships, over every `.docx` fixture in `test-data`.

use std::path::{Path, PathBuf};

use quick_xml::Reader;
use quick_xml::events::Event;

use super::{HalfPoints, VerticalAlign};
use crate::table::VMergeState;
use crate::{Inline, Package, Paragraph, Run, source_backed};
use litchi_core::VerticalPosition;

fn collect_fixtures(directory: &Path, found: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(directory) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect_fixtures(&path, found);
        } else if path
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("docx"))
        {
            found.push(path);
        }
    }
}

fn fixtures() -> Vec<PathBuf> {
    let mut found = Vec::new();
    collect_fixtures(Path::new("../../test-data"), &mut found);
    found.sort();
    found
}

/// One element of the oracle tree.
struct Node {
    prefix: Vec<u8>,
    local: Vec<u8>,
    start: usize,
    end: usize,
    parent: Option<usize>,
    children: Vec<usize>,
}

/// An independent reading of `OfficeCLI`'s `GetNavRunIndex`
/// (`WordHandler.Navigation.cs:970-1004`) over a paragraph: every
/// descendant run, minus comment-reference runs and runs below an inline
/// content control, a simple field, a text box inside the paragraph or a
/// markup-compatibility wrapper. Word elements are those written with the
/// paragraph's own prefix, as every fixture writes them.
fn oracle_runs(xml: &[u8]) -> Vec<&[u8]> {
    let mut reader = Reader::from_reader(xml);
    let mut nodes: Vec<Node> = Vec::new();
    let mut open: Vec<usize> = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap();
        let event = reader.read_event().unwrap();
        let end = usize::try_from(reader.buffer_position()).unwrap();
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let empty = end > 1 && xml[end - 2] == b'/';
                let name = element.name();
                let index = nodes.len();
                nodes.push(Node {
                    prefix: name
                        .prefix()
                        .map(|prefix| prefix.into_inner().to_vec())
                        .unwrap_or_default(),
                    local: name.local_name().as_ref().to_vec(),
                    start,
                    end,
                    parent: open.last().copied(),
                    children: Vec::new(),
                });
                if let Some(parent) = open.last() {
                    nodes[*parent].children.push(index);
                }
                if !empty {
                    open.push(index);
                }
            },
            Event::End(_) => {
                let index = open.pop().unwrap();
                nodes[index].end = end;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let word = nodes[0].prefix.clone();
    let is_word = |node: &Node, local: &[u8]| node.prefix == word && node.local == local;
    let mut runs = Vec::new();
    for (index, node) in nodes.iter().enumerate().skip(1) {
        if !is_word(node, b"r") {
            continue;
        }
        if node
            .children
            .iter()
            .any(|child| is_word(&nodes[*child], b"commentReference"))
        {
            continue;
        }
        let mut excluded = false;
        let mut ancestor = node.parent;
        while let Some(current) = ancestor {
            if current == 0 {
                break;
            }
            let element = &nodes[current];
            if is_word(element, b"sdt")
                || is_word(element, b"fldSimple")
                || is_word(element, b"txbxContent")
                || matches!(
                    element.local.as_slice(),
                    b"AlternateContent" | b"Choice" | b"Fallback"
                )
            {
                excluded = true;
                break;
            }
            ancestor = element.parent;
        }
        if !excluded {
            runs.push(&xml[nodes[index].start..nodes[index].end]);
        }
    }
    runs
}

#[derive(Debug, Default)]
struct Counts {
    fixtures: usize,
    paragraphs: usize,
    paragraphs_with_properties: usize,
    runs: usize,
    runs_with_properties: usize,
    addressable_runs: usize,
    outermost_runs: usize,
    tables: usize,
    rows: usize,
    cells: usize,
    spanned_cells: usize,
    invalid: usize,
}

/// The checked reader agrees with the lenient single-property readers
/// wherever the checked value is valid.
fn check_run(run: &Run, counts: &mut Counts, label: &str) {
    let checked = run
        .direct_properties()
        .unwrap_or_else(|error| panic!("{label}: {error}"));
    let invalid: Vec<&str> = checked.invalid_properties().collect();
    counts.invalid += invalid.len();
    counts.runs += 1;
    if checked != super::Run::default() {
        counts.runs_with_properties += 1;
    }
    if !invalid.contains(&"b") {
        assert_eq!(run.bold().unwrap(), checked.bold, "{label}: bold");
    }
    if !invalid.contains(&"i") {
        assert_eq!(run.italic().unwrap(), checked.italic, "{label}: italic");
    }
    if !invalid.contains(&"strike") {
        assert_eq!(
            run.strikethrough().unwrap(),
            checked.strike,
            "{label}: strike"
        );
    }
    if !invalid.contains(&"vertAlign") {
        let expected = match checked.vertical_align {
            Some(VerticalAlign::Superscript) => Some(VerticalPosition::Superscript),
            Some(VerticalAlign::Subscript) => Some(VerticalPosition::Subscript),
            Some(VerticalAlign::Baseline) | None => None,
        };
        assert_eq!(
            run.vertical_position().unwrap(),
            expected,
            "{label}: vertAlign"
        );
    }
    if !invalid.contains(&"sz") {
        assert_eq!(
            run.font_size().unwrap(),
            checked
                .size
                .map(|HalfPoints(size)| u32::try_from(size).unwrap()),
            "{label}: sz"
        );
    }
    if !invalid.contains(&"rFonts") {
        assert_eq!(
            run.font_name().unwrap(),
            checked.fonts.as_ref().and_then(|fonts| fonts.ascii.clone()),
            "{label}: rFonts"
        );
    }
    let legacy = run.get_properties().unwrap();
    assert_eq!(legacy.bold, run.bold().unwrap(), "{label}: get_properties");
    assert_eq!(
        legacy.underline,
        checked.underline.map(|underline| underline.style),
        "{label}: u"
    );
}

fn check_paragraph(paragraph: &Paragraph, counts: &mut Counts, label: &str) {
    counts.paragraphs += 1;
    let checked = paragraph
        .direct_properties()
        .unwrap_or_else(|error| panic!("{label}: {error}"));
    counts.invalid += checked.invalid_properties().count();
    if checked != super::Paragraph::default() {
        counts.paragraphs_with_properties += 1;
    }
    if let Ok(style) = paragraph.style_id() {
        assert_eq!(style, checked.style, "{label}: pStyle");
    }
    if let Ok(Some(spacing)) = paragraph.spacing() {
        assert_eq!(Some(spacing), checked.spacing, "{label}: spacing");
    }

    let addressable = paragraph.addressable_runs().unwrap();
    let expected = oracle_runs(paragraph.xml_bytes());
    let actual: Vec<&[u8]> = addressable.iter().map(Run::xml_bytes).collect();
    assert_eq!(actual, expected, "{label}: addressable runs");
    counts.addressable_runs += addressable.len();

    let outermost = paragraph.runs().unwrap();
    counts.outermost_runs += outermost.len();
    for (index, run) in outermost.iter().chain(addressable.iter()).enumerate() {
        check_run(run, counts, &format!("{label} run {index}"));
    }
}

fn check_table(table: &crate::Table, counts: &mut Counts, label: &str) {
    counts.tables += 1;
    let checked = table
        .properties()
        .unwrap_or_else(|error| panic!("{label}: {error}"));
    counts.invalid += checked.invalid_properties().count();
    for (row_index, row) in table.rows().unwrap().iter().enumerate() {
        counts.rows += 1;
        let row_properties = row.properties().unwrap();
        counts.invalid += row_properties.invalid_properties().count();
        for (cell_index, cell) in row.cells().unwrap().iter().enumerate() {
            counts.cells += 1;
            let label = format!("{label} row {row_index} cell {cell_index}");
            let properties = cell
                .properties()
                .unwrap_or_else(|error| panic!("{label}: {error}"));
            let invalid: Vec<&str> = properties.invalid_properties().collect();
            counts.invalid += invalid.len();
            if !invalid.contains(&"gridSpan") {
                let span = cell.grid_span().unwrap();
                assert_eq!(
                    span,
                    properties
                        .grid_span
                        .map_or(1, |span| usize::try_from(span).unwrap()),
                    "{label}: gridSpan"
                );
                if span > 1 {
                    counts.spanned_cells += 1;
                }
            }
            if !invalid.contains(&"vMerge") {
                assert_eq!(
                    cell.v_merge().unwrap(),
                    properties.vertical_merge.map(|merge| match merge {
                        super::Merge::Restart => VMergeState::Restart,
                        super::Merge::Continue => VMergeState::Continue,
                    }),
                    "{label}: vMerge"
                );
            }
            for paragraph in cell.paragraphs().unwrap() {
                check_paragraph(&paragraph, counts, &label);
            }
        }
    }
}

#[test]
fn the_corpus_reads_with_checked_and_legacy_readers_in_agreement() {
    let mut counts = Counts::default();
    for path in fixtures() {
        let Ok(package) = Package::open(&path) else {
            continue;
        };
        let Ok(document) = package.document() else {
            continue;
        };
        let label = path.display().to_string();
        counts.fixtures += 1;
        for (index, paragraph) in document.paragraphs().unwrap().iter().enumerate() {
            check_paragraph(
                paragraph,
                &mut counts,
                &format!("{label} paragraph {index}"),
            );
        }
        for (index, table) in document.tables().unwrap().iter().enumerate() {
            check_table(table, &mut counts, &format!("{label} table {index}"));
        }
    }
    // The corpus exercised every reader (59 fixtures, 913 paragraphs, 4,143
    // runs, 24 tables, 59 rows and 186 cells when this was written; no
    // fixture merges cells vertically, which the unit tests cover).
    assert!(counts.fixtures >= 50, "{counts:?}");
    assert!(counts.paragraphs_with_properties > 500, "{counts:?}");
    assert!(counts.runs_with_properties > 1_000, "{counts:?}");
    assert!(
        counts.tables >= 20 && counts.rows >= 50 && counts.cells >= 150,
        "{counts:?}"
    );
    assert!(counts.spanned_cells > 0, "{counts:?}");
    // Comment-reference, content-control and field runs leave the set.
    assert!(counts.addressable_runs < counts.outermost_runs);
    assert_eq!(counts.invalid, 0, "the corpus states only valid values");
    assert!(counts.paragraphs > 0 && counts.runs > 0, "{counts:?}");
}

/// Hyperlinks of source-owner paragraphs resolve through the source owner's
/// main-document relationships exactly as through the owning package's.
#[test]
fn source_owner_relationships_resolve_hyperlinks_like_the_owning_package() {
    let mut links = 0usize;
    for name in ["Hyperlink.docx", "drawing.docx"] {
        let path = Path::new("../../test-data/ooxml/docx").join(name);
        let owning = Package::open(&path).unwrap();
        let document = owning.document().unwrap();
        let owning_relationships = owning.opc_package().main_document_part().unwrap().rels();
        let source = source_backed::Package::open(&path).unwrap();
        let source_relationships = main_document_relationships(&source).unwrap();
        let source_document = source.document().unwrap();

        let owning_paragraphs = document.paragraphs().unwrap();
        let source_paragraphs = source_document.paragraphs().unwrap();
        assert_eq!(owning_paragraphs.len(), source_paragraphs.len());
        for (owning, source) in owning_paragraphs.iter().zip(source_paragraphs.iter()) {
            let left = owning
                .inlines_with_relationships(owning_relationships)
                .unwrap();
            let right = source
                .inlines_with_relationships(source_relationships)
                .unwrap();
            assert_eq!(left.len(), right.len());
            for (left, right) in left.iter().zip(right.iter()) {
                match (left, right) {
                    (Inline::Hyperlink(left), Inline::Hyperlink(right)) => {
                        links += 1;
                        assert_eq!(left.link().url(), right.link().url());
                        assert_eq!(left.link().anchor(), right.link().anchor());
                        assert_eq!(left.runs().len(), right.runs().len());
                    },
                    (Inline::Hyperlink(_), _) | (_, Inline::Hyperlink(_)) => {
                        panic!("{name}: hyperlink projections differ")
                    },
                    _ => {},
                }
            }
        }
    }
    assert!(links > 0, "the fixtures carry hyperlinks");
}

/// A source whose revision can be bumped, to observe a changed source.
#[derive(Debug)]
struct MutableSource {
    bytes: Vec<u8>,
    revision: std::sync::atomic::AtomicU64,
}

impl litchi_core::ReadAt for MutableSource {
    fn len(&self) -> std::io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let offset = usize::try_from(offset).unwrap_or(usize::MAX);
        let Some(available) = self.bytes.get(offset..) else {
            return Ok(0);
        };
        let length = available.len().min(output.len());
        output[..length].copy_from_slice(&available[..length]);
        Ok(length)
    }

    fn version(&self) -> std::io::Result<litchi_core::SourceVersion> {
        Ok(litchi_core::SourceVersion::new(
            0x0010,
            self.revision.load(std::sync::atomic::Ordering::SeqCst),
        ))
    }
}

#[test]
fn source_owner_relationships_refuse_a_changed_source() {
    let source = std::sync::Arc::new(MutableSource {
        bytes: std::fs::read("../../test-data/ooxml/docx/Hyperlink.docx").unwrap(),
        revision: std::sync::atomic::AtomicU64::new(0),
    });
    let package = source_backed::Package::from_read_at(source.clone()).unwrap();
    assert!(!main_document_relationships(&package).unwrap().is_empty());
    source
        .revision
        .fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    assert!(matches!(
        main_document_relationships(&package),
        Err(crate::Error::Opc(
            litchi_opc::OpcError::SourceChanged { .. }
        ))
    ));
}

// Corpus-only physical relationship access; ordinary leaf APIs do not gain a
// second raw relationship surface merely to support this comparison.
fn main_document_relationships(
    package: &source_backed::Package,
) -> crate::Result<&litchi_opc::rel::Relationships> {
    Ok(package
        .source_backed_package()?
        .main_document_part()?
        .rels())
}
