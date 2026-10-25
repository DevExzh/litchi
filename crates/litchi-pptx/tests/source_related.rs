//! Source-backed PPTX related-part snapshots.
//!
//! These tests exercise the source owner through its public, detached related
//! snapshots.  The fixture is assembled and serialized in memory so the
//! source path still opens a real ZIP package and validates the OPC graph.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration assertions intentionally panic on malformed fixtures"
)]

use std::io::{self, Cursor};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, Position,
    ReadAt, SourceVersion,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part, SourceLimitKind, TargetMode,
};
use litchi_pptx::chart::{Chart as EagerChart, Series, Type};
use litchi_pptx::presentation::related::raw;
use litchi_pptx::shape::owned::Scene as OwnedScene;
use litchi_pptx::{
    CodecLimitKind, Error, LimitKind, OpenOptions, Package, SlideLayoutKind,
    SourceBackedPresentation, SourceLimits,
};

const SLIDE_ONE: &str = "/ppt/slides/slide1.xml";
const LAYOUT: &str = "/ppt/slideLayouts/slideLayout1.xml";
const NOTES: &str = "/ppt/notesSlides/notesSlide1.xml";
const NOTES_MASTER: &str = "/ppt/notesMasters/notesMaster1.xml";
const NOTES_MASTER_TWO: &str = "/ppt/notesMasters/notesMaster2.xml";
const COMMENTS: &str = "/ppt/comments/comment1.xml";
const AUTHORS: &str = "/ppt/commentAuthors.xml";
const CHART: &str = "/ppt/charts/chart1.xml";
const CHART_TWO: &str = "/ppt/charts/chart2.xml";
const COMMENTS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const AUTHORS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors";

const COMMENT_LIST: &str = concat!(
    r#"<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
    r#"<p:cm authorId="0" dt="2026-01-02T03:04:05Z" idx="1"><p:pos x="120" y="240"/>"#,
    r#"<p:text>Review this slide</p:text></p:cm></p:cmLst>"#,
);
const COMMENT_AUTHORS: &str = concat!(
    r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
    r#"<p:cmAuthor id="0" name="Ada Lovelace" initials="AL" lastIdx="1" clrIdx="0"/>"#,
    r#"</p:cmAuthorLst>"#,
);

/// Build a two-slide in-memory ZIP.  The first slide has every related part
/// exercised below; the second deliberately has no notes, comments, or chart.
fn rich_deck() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    {
        let first = builder.presentation().unwrap().add_slide().unwrap();
        first.set_title("Overview");
        first.add_text_box("Body text", 10, 20, 30, 40);
        first.set_notes("First note");
    }
    {
        let second = builder.presentation().unwrap().add_slide().unwrap();
        second.set_title("Details");
    }

    let package = builder.build().unwrap();
    let chart = EagerChart::new(Type::Bar, 914_400, 914_400, 4_572_000, 2_743_200)
        .with_title("Quarterly sales")
        .with_legend(true)
        .add_series(
            Series::new("2026")
                .with_categories(vec!["Q1".into(), "Q2".into()])
                .with_values(vec![100.0, 150.0]),
        );
    // SLIDE_ONE is the first slide.
    let mut edit = package.edit_charts().unwrap();
    edit.add(0_usize, &chart).unwrap();
    let package = edit.commit().unwrap().into_package();
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();

    opc.try_add_part(Box::new(BlobPart::new(
        PackURI::new(COMMENTS).unwrap(),
        ct::PML_COMMENTS.into(),
        COMMENT_LIST.as_bytes().to_vec(),
    )))
    .unwrap();
    opc.try_add_part(Box::new(BlobPart::new(
        PackURI::new(AUTHORS).unwrap(),
        ct::PML_COMMENT_AUTHORS.into(),
        COMMENT_AUTHORS.as_bytes().to_vec(),
    )))
    .unwrap();
    opc.get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            AUTHORS_REL.into(),
            "commentAuthors.xml".into(),
            "rIdCommentAuthors".into(),
            TargetMode::Internal,
        )
        .unwrap();
    opc.get_part_mut(&PackURI::new(SLIDE_ONE).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            COMMENTS_REL.into(),
            "../comments/comment1.xml".into(),
            "rIdComments".into(),
            TargetMode::Internal,
        )
        .unwrap();

    // Give the shared generated layout stable semantic values to compare with
    // the eager owner.  The part keeps its normal master relationship.
    edit_blob(&mut opc, LAYOUT, |xml| {
        xml.replacen(" type=\"title\"", " type=\"chart\"", 1)
            .replacen("<p:cSld name=\"Title Slide\"", "<p:cSld name=\"Charts\"", 1)
    });

    PackageWriter::to_bytes(&opc).unwrap()
}

fn open(bytes: &[u8]) -> SourceBackedPresentation {
    SourceBackedPresentation::from_reader(Cursor::new(bytes)).unwrap()
}

fn rewrite(bytes: &[u8], edit: impl FnOnce(&mut OpcPackage)) -> Vec<u8> {
    let mut opc = OpcPackage::from_bytes(bytes, SourceLimits::default()).unwrap();
    edit(&mut opc);
    PackageWriter::to_bytes(&opc).unwrap()
}

fn edit_blob(opc: &mut OpcPackage, part_name: &str, edit: impl FnOnce(String) -> String) {
    let uri = PackURI::new(part_name).unwrap();
    let part = opc.get_part_mut(&uri).unwrap();
    let source = String::from_utf8(part.blob().to_vec()).unwrap();
    part.set_blob(edit(source).into_bytes());
}

fn selected_notes_bytes(bytes: &[u8]) -> usize {
    let opc = OpcPackage::from_bytes(bytes, SourceLimits::default()).unwrap();
    let notes = opc
        .get_part(&PackURI::new(NOTES).unwrap())
        .unwrap()
        .blob()
        .len();
    let master = opc.get_part(&PackURI::new(NOTES_MASTER).unwrap()).unwrap();
    let theme_uri = master
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::THEME)
        .unwrap()
        .target_partname()
        .unwrap();
    notes + master.blob().len() + opc.get_part(&theme_uri).unwrap().blob().len()
}

fn notes_limits(max_total_bytes: usize) -> litchi_pptx::notes::Limits {
    litchi_pptx::notes::Limits {
        max_total_bytes,
        ..litchi_pptx::notes::Limits::DEFAULT
    }
}

fn replace_relationship(
    opc: &mut OpcPackage,
    owner: &str,
    relationship_type: &str,
    target: &str,
    target_mode: TargetMode,
) {
    let owner = PackURI::new(owner).unwrap();
    let part = opc.get_part_mut(&owner).unwrap();
    let id = part
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == relationship_type)
        .unwrap()
        .r_id()
        .to_owned();
    part.rels_mut().remove(&id).unwrap();
    part.rels_mut()
        .try_add_relationship(
            relationship_type.to_owned(),
            target.to_owned(),
            id,
            target_mode,
        )
        .unwrap();
}

fn remove_relationship(opc: &mut OpcPackage, owner: &str, relationship_type: &str) {
    let owner = PackURI::new(owner).unwrap();
    let part = opc.get_part_mut(&owner).unwrap();
    let ids: Vec<_> = part
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type)
        .map(|relationship| relationship.r_id().to_owned())
        .collect();
    for id in ids {
        part.rels_mut().remove(&id).unwrap();
    }
}

fn add_duplicate_relationship(opc: &mut OpcPackage, owner: &str, relationship_type: &str) {
    let owner = PackURI::new(owner).unwrap();
    opc.get_part_mut(&owner)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            relationship_type.to_owned(),
            "../slideLayouts/slideLayout1.xml".into(),
            "rIdDuplicate".into(),
            TargetMode::Internal,
        )
        .unwrap();
}

fn owned_signature(scene: &OwnedScene) -> Vec<(Option<String>, Option<String>)> {
    scene
        .iter()
        .map(|handle| {
            let shape = handle.view();
            (
                shape.name().map(str::to_owned),
                shape.text().map(str::to_owned),
            )
        })
        .collect()
}

fn eager_signature(scene: &litchi_pptx::shape::Scene<'_>) -> Vec<(Option<String>, Option<String>)> {
    scene
        .iter()
        .map(|shape| {
            (
                shape.name().map(str::to_owned),
                shape.text().map(str::to_owned),
            )
        })
        .collect()
}

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(61, self.revision.load(Ordering::SeqCst)))
    }
}

struct ArmedRevision {
    inner: VersionedSource,
    checks: AtomicUsize,
    change_on: AtomicUsize,
}

impl ArmedRevision {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            inner: VersionedSource::new(bytes),
            checks: AtomicUsize::new(0),
            change_on: AtomicUsize::new(usize::MAX),
        }
    }
}

impl ReadAt for ArmedRevision {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.inner.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        let check = self.checks.fetch_add(1, Ordering::SeqCst) + 1;
        if check == self.change_on.load(Ordering::SeqCst) {
            self.inner.changed();
        }
        self.inner.version()
    }
}

fn managed_context(memory_limit: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "pptx-related-source-test",
        Limits::new(
            memory_limit,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(memory_limit.max(1)).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    (budget, cancellation_source, context)
}

#[test]
fn source_related_snapshots_match_eager_values_and_raw_identity() {
    let bytes = rich_deck();
    let eager = Package::from_bytes(&bytes).unwrap();
    let eager_presentation = eager.presentation().unwrap();
    let eager_slide = eager_presentation.slides().unwrap().remove(0);
    let source = open(&bytes);
    let slide = source.slide(0).unwrap();

    assert!(!slide.is_hidden().unwrap());
    assert_eq!(raw::slide_part_uri(&slide).as_str(), SLIDE_ONE);
    assert!(!raw::slide_relationship_id(&slide).is_empty());

    let layout = slide.layout().unwrap().unwrap();
    let eager_layout = eager_slide.layout().unwrap().unwrap();
    assert_eq!(layout.name(), Some("Charts"));
    assert_eq!(layout.kind(), Some(SlideLayoutKind::Chart));
    assert_eq!(layout.name(), Some(eager_layout.name().unwrap().as_str()));
    assert_eq!(
        layout.kind(),
        eager_layout
            .kind()
            .unwrap()
            .as_deref()
            .and_then(SlideLayoutKind::from_token)
    );
    assert_eq!(
        owned_signature(layout.scene()),
        eager_signature(&eager_layout.shapes().unwrap())
    );

    let notes = slide.notes().unwrap().unwrap();
    let eager_notes = eager_presentation.notes().unwrap().unwrap();
    let eager_note_text = eager_notes.slides()[0]
        .text(litchi_pptx::notes::Limits::DEFAULT)
        .unwrap();
    assert_eq!(notes.text(), eager_note_text.as_deref());
    assert!(!notes.scene().is_empty());

    let comments = slide.comments().unwrap().unwrap();
    let eager_comments = eager_slide
        .comments()
        .unwrap()
        .unwrap()
        .comments(litchi_pptx::comments::Limits::DEFAULT)
        .unwrap();
    assert_eq!(comments.entries().len(), eager_comments.len());
    let entry = &comments.entries()[0];
    let eager_entry = &eager_comments[0];
    assert_eq!(entry.text(), eager_entry.text);
    assert_eq!(entry.date_time(), eager_entry.date_time.as_deref());
    assert_eq!(entry.location(), (eager_entry.x, eager_entry.y));
    assert_eq!(entry.author_name(), Some("Ada Lovelace"));
    assert_eq!(entry.initials(), Some("AL"));

    let charts = slide.charts().unwrap();
    assert_eq!(charts.len(), 1);
    let chart = slide.chart(Position::new(0)).unwrap().unwrap();
    let eager_chart = eager_slide.charts().unwrap().remove(0);
    let eager_info = eager_chart.chart_info().unwrap();
    assert_eq!(chart.kind(), eager_info.chart_type);
    assert_eq!(chart.title(), eager_info.title.as_deref());
    assert_eq!(chart.has_legend(), eager_info.has_legend);
    assert_eq!(charts[0].kind(), Type::Bar);

    // The generic raw surface is the only place where the physical part and
    // relationship IDs are observable; semantic snapshots carry no owner.
    assert_eq!(raw::part_uri(&layout).as_str(), LAYOUT);
    assert_eq!(raw::part_uri(&notes).as_str(), NOTES);
    assert_eq!(raw::part_uri(&comments).as_str(), COMMENTS);
    assert!(!raw::part_uri(&chart).as_str().is_empty());
    assert!(!raw::relationship_id(&layout).is_empty());
    assert!(!raw::relationship_id(&notes).is_empty());
    assert!(!raw::relationship_id(&comments).is_empty());
    assert!(!raw::relationship_id(&chart).is_empty());
    assert!(!raw::data(&layout).is_empty());
    assert!(!raw::data(&notes).is_empty());
    assert!(!raw::data(&comments).is_empty());
    assert!(!raw::data(&chart).is_empty());

    assert!(slide.chart(Position::new(1)).unwrap().is_none());
}

#[test]
fn absent_related_relationships_return_none_or_empty() {
    let bytes = rich_deck();
    let source = open(&bytes);
    let slide = source.slide(1).unwrap();

    assert!(slide.notes().unwrap().is_none());
    assert!(slide.comments().unwrap().is_none());
    assert!(slide.charts().unwrap().is_empty());
    assert!(slide.chart(Position::new(0)).unwrap().is_none());
}

#[test]
fn indexed_chart_loads_only_the_selected_relationship() {
    let bytes = rewrite(&rich_deck(), |opc| {
        opc.try_add_part(Box::new(BlobPart::new(
            PackURI::new(CHART_TWO).unwrap(),
            "application/octet-stream".into(),
            b"<wrong/>".to_vec(),
        )))
        .unwrap();
        opc.get_part_mut(&PackURI::new(SLIDE_ONE).unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::CHART.into(),
                "../charts/chart2.xml".into(),
                "rIdSecondChart".into(),
                TargetMode::Internal,
            )
            .unwrap();
    });
    let slide = open(&bytes).slide(0).unwrap();

    let first = slide.chart(Position::new(0)).unwrap().unwrap();
    assert_eq!(first.kind(), Type::Bar);
    assert!(slide.charts().is_err());
    assert!(slide.chart(Position::new(1)).is_err());
}

#[test]
fn detached_snapshots_survive_source_editor_conversion_and_drop() {
    let bytes = rich_deck();
    let source = open(&bytes);
    let (layout, notes, comments, chart, layout_bytes) = {
        let slide = source.slide(0).unwrap();
        let layout = slide.layout().unwrap().unwrap();
        let notes = slide.notes().unwrap().unwrap();
        let comments = slide.comments().unwrap().unwrap();
        let chart = slide.chart(Position::new(0)).unwrap().unwrap();
        let layout_bytes = raw::data(&layout).to_vec();
        (layout, notes, comments, chart, layout_bytes)
    };

    let _package = litchi_pptx::source::Package::from_presentation(&source).unwrap();
    drop(source);
    assert_eq!(layout.name(), Some("Charts"));
    assert_eq!(notes.text(), Some("First note"));
    assert_eq!(comments.entries()[0].text(), "Review this slide");
    assert_eq!(chart.title(), Some("Quarterly sales"));
    assert!(!layout_bytes.is_empty());
}

#[test]
fn snapshot_debug_is_content_free_and_omits_physical_identity() {
    let bytes = rich_deck();
    let source = open(&bytes);
    let slide = source.slide(0).unwrap();
    let layout = slide.layout().unwrap().unwrap();
    let notes = slide.notes().unwrap().unwrap();
    let comments = slide.comments().unwrap().unwrap();
    let chart = slide.chart(Position::new(0)).unwrap().unwrap();
    let debug = format!("{layout:?} {notes:?} {comments:?} {chart:?}");

    for authored in [
        "Charts",
        "First note",
        "Review this slide",
        "Quarterly sales",
        "rId",
        "/ppt/",
    ] {
        assert!(
            !debug.contains(authored),
            "debug leaked {authored}: {debug}"
        );
    }
}

#[test]
fn related_reads_refuse_wrong_content_type_external_and_duplicate_edges() {
    let bytes = rich_deck();

    let wrong_layout_type = rewrite(&bytes, |opc| {
        replace_relationship(
            opc,
            SLIDE_ONE,
            rt::SLIDE_LAYOUT,
            "../comments/comment1.xml",
            TargetMode::Internal,
        );
    });
    assert!(open(&wrong_layout_type).slide(0).unwrap().layout().is_err());

    let external_chart = rewrite(&bytes, |opc| {
        replace_relationship(
            opc,
            SLIDE_ONE,
            rt::CHART,
            "https://example.invalid/chart.xml",
            TargetMode::External,
        );
    });
    let slide = open(&external_chart).slide(0).unwrap();
    assert!(slide.charts().is_err());
    assert!(slide.chart(Position::new(0)).is_err());

    let duplicate_layout = rewrite(&bytes, |opc| {
        add_duplicate_relationship(opc, SLIDE_ONE, rt::SLIDE_LAYOUT);
    });
    assert!(open(&duplicate_layout).slide(0).unwrap().layout().is_err());

    let external_comments = rewrite(&bytes, |opc| {
        replace_relationship(
            opc,
            SLIDE_ONE,
            COMMENTS_REL,
            "https://example.invalid/comments.xml",
            TargetMode::External,
        );
    });
    assert!(
        open(&external_comments)
            .slide(0)
            .unwrap()
            .comments()
            .is_err()
    );

    let external_authors = rewrite(&bytes, |opc| {
        replace_relationship(
            opc,
            "/ppt/presentation.xml",
            AUTHORS_REL,
            "https://example.invalid/commentAuthors.xml",
            TargetMode::External,
        );
    });
    assert!(
        open(&external_authors)
            .slide(0)
            .unwrap()
            .comments()
            .is_err()
    );
}

#[test]
fn source_comments_refuse_malformed_duplicate_and_dangling_authors() {
    let bytes = rich_deck();

    let malformed_authors = rewrite(&bytes, |opc| {
        edit_blob(opc, AUTHORS, |xml| xml.replace("clrIdx=\"0\"", ""));
    });
    assert!(
        open(&malformed_authors)
            .slide(0)
            .unwrap()
            .comments()
            .is_err()
    );

    let duplicate_authors = rewrite(&bytes, |opc| {
        edit_blob(opc, AUTHORS, |xml| {
            xml.replacen(
                "</p:cmAuthorLst>",
                r#"<p:cmAuthor id="0" name="Grace Hopper" initials="GH" lastIdx="1" clrIdx="1"/></p:cmAuthorLst>"#,
                1,
            )
        });
    });
    assert!(
        open(&duplicate_authors)
            .slide(0)
            .unwrap()
            .comments()
            .is_err()
    );

    let dangling_author = rewrite(&bytes, |opc| {
        edit_blob(opc, COMMENTS, |xml| {
            xml.replace("authorId=\"0\"", "authorId=\"99\"")
        });
    });
    assert!(open(&dangling_author).slide(0).unwrap().comments().is_err());
}

#[test]
fn notes_require_the_slide_backlink_and_notes_master_theme() {
    let bytes = rich_deck();

    let missing_backlink = rewrite(&bytes, |opc| {
        remove_relationship(opc, NOTES, rt::SLIDE);
    });
    assert!(open(&missing_backlink).slide(0).unwrap().notes().is_err());

    let missing_theme = rewrite(&bytes, |opc| {
        remove_relationship(opc, NOTES_MASTER, rt::THEME);
    });
    assert!(open(&missing_theme).slide(0).unwrap().notes().is_err());
}

#[test]
fn notes_selected_aggregate_is_tight_for_source_and_eager_reads() {
    let bytes = rich_deck();
    let selected = selected_notes_bytes(&bytes);
    assert!(selected > 0);

    let source = SourceBackedPresentation::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes.clone())),
        OpenOptions::default().notes_limits(notes_limits(selected)),
    )
    .unwrap();
    assert!(source.slide(0).unwrap().notes().unwrap().is_some());

    let source = SourceBackedPresentation::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes.clone())),
        OpenOptions::default().notes_limits(notes_limits(selected - 1)),
    )
    .unwrap();
    let error = source.slide(0).unwrap().notes().unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            kind: LimitKind::Codec(CodecLimitKind::XmlBytes),
            observed,
            maximum,
        } if observed == selected && maximum == selected - 1
    ));

    let package = Package::from_bytes_with_options(
        &bytes,
        OpenOptions::default().notes_limits(notes_limits(selected)),
    )
    .unwrap();
    assert!(
        package
            .presentation()
            .unwrap()
            .slide(0)
            .unwrap()
            .unwrap()
            .notes()
            .unwrap()
            .is_some()
    );

    let package = Package::from_bytes_with_options(
        &bytes,
        OpenOptions::default().notes_limits(notes_limits(selected - 1)),
    )
    .unwrap();
    let error = package
        .presentation()
        .unwrap()
        .slide(0)
        .unwrap()
        .unwrap()
        .notes()
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            kind: LimitKind::Codec(CodecLimitKind::XmlBytes),
            observed,
            maximum,
        } if observed == selected && maximum == selected - 1
    ));
}

#[test]
fn notes_selected_aggregate_does_not_scan_unrelated_notes_masters() {
    let bytes = rewrite(&rich_deck(), |opc| {
        opc.try_add_part(Box::new(BlobPart::new(
            PackURI::new(NOTES_MASTER_TWO).unwrap(),
            ct::PML_NOTES_MASTER.to_owned(),
            b"<broken-notes-master/>".to_vec(),
        )))
        .unwrap();
    });
    let selected = selected_notes_bytes(&bytes);

    let source = SourceBackedPresentation::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes.clone())),
        OpenOptions::default().notes_limits(notes_limits(selected)),
    )
    .unwrap();
    assert!(source.slide(0).unwrap().notes().unwrap().is_some());

    let package = Package::from_bytes_with_options(
        &bytes,
        OpenOptions::default().notes_limits(notes_limits(selected)),
    )
    .unwrap();
    assert!(
        package
            .presentation()
            .unwrap()
            .slide(0)
            .unwrap()
            .unwrap()
            .notes()
            .unwrap()
            .is_some()
    );
}

#[test]
fn notes_refuse_mixed_conformance_duplicate_graph_edges() {
    let bytes = rewrite(&rich_deck(), |opc| {
        let notes = opc.get_part_mut(&PackURI::new(NOTES).unwrap()).unwrap();
        notes
            .rels_mut()
            .try_add_relationship(
                "http://purl.oclc.org/ooxml/officeDocument/relationships/slide".into(),
                "../slides/slide1.xml".into(),
                "rIdStrictSlide".into(),
                TargetMode::Internal,
            )
            .unwrap();
        notes
            .rels_mut()
            .try_add_relationship(
                rt::STRICT_NOTES_MASTER.into(),
                "../notesMasters/notesMaster1.xml".into(),
                "rIdStrictMaster".into(),
                TargetMode::Internal,
            )
            .unwrap();
    });
    assert!(open(&bytes).slide(0).unwrap().notes().is_err());
}

#[test]
fn notes_refuse_a_notes_master_that_disagrees_with_the_presentation_catalog() {
    let bytes = rewrite(&rich_deck(), |opc| {
        let (content_type, blob) = {
            let master = opc.get_part(&PackURI::new(NOTES_MASTER).unwrap()).unwrap();
            (master.content_type().to_owned(), master.blob().to_vec())
        };
        let mut second = BlobPart::new(PackURI::new(NOTES_MASTER_TWO).unwrap(), content_type, blob);
        second
            .rels_mut()
            .try_add_relationship(
                rt::THEME.into(),
                "../theme/theme2.xml".into(),
                "rIdTheme".into(),
                TargetMode::Internal,
            )
            .unwrap();
        opc.try_add_part(Box::new(second)).unwrap();
        replace_relationship(
            opc,
            NOTES,
            rt::NOTES_MASTER,
            "../notesMasters/notesMaster2.xml",
            TargetMode::Internal,
        );
    });
    assert!(open(&bytes).slide(0).unwrap().notes().is_err());
}

#[test]
fn notes_empty_or_absent_body_have_no_projected_text() {
    let empty_text = rewrite(&rich_deck(), |opc| {
        edit_blob(opc, NOTES, |xml| {
            assert!(xml.contains("First note"));
            xml.replacen("First note", "", 1)
        });
    });
    let absent_body = rewrite(&rich_deck(), |opc| {
        edit_blob(opc, NOTES, |xml| {
            xml.replacen("type=\"body\"", "type=\"other\"", 1)
        });
    });
    for bytes in [empty_text, absent_body] {
        let notes = open(&bytes).slide(0).unwrap().notes().unwrap().unwrap();
        assert_eq!(notes.text(), None);
    }
}

#[test]
fn comments_refuse_outbound_relationships() {
    let bytes = rewrite(&rich_deck(), |opc| {
        opc.get_part_mut(&PackURI::new(COMMENTS).unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                "urn:test:comments-outbound".into(),
                "../slides/slide1.xml".into(),
                "rIdOutbound".into(),
                TargetMode::Internal,
            )
            .unwrap();
    });
    assert!(open(&bytes).slide(0).unwrap().comments().is_err());
}

#[test]
fn notes_refuse_duplicate_body_placeholders() {
    let duplicate_body = rewrite(&rich_deck(), |opc| {
        edit_blob(opc, NOTES, |xml| {
            let start = xml.find("<p:sp><p:nvSpPr>").unwrap();
            let end = start + xml[start..].find("</p:sp>").unwrap() + "</p:sp>".len();
            let duplicate = xml[start..end]
                .replacen("id=\"2\"", "id=\"3\"", 1)
                .replacen(
                    "name=\"Notes Placeholder\"",
                    "name=\"Notes Placeholder 2\"",
                    1,
                );
            xml.replacen("</p:spTree>", &format!("{duplicate}</p:spTree>"), 1)
        });
    });

    assert!(open(&duplicate_body).slide(0).unwrap().notes().is_err());
}

#[test]
fn source_related_scene_reuses_and_trims_the_public_scene_cache() {
    let source = open(&rich_deck());
    let slide = source.slide(0).unwrap();

    let first_layout = slide.layout().unwrap().unwrap();
    let first = first_layout.scene();
    let after_first = source.scene_cache_diagnostics();
    assert_eq!(after_first.misses, 1);
    assert_eq!(after_first.builds, 1);
    assert_eq!(after_first.retained_entries, 1);

    let second_layout = source.clone().slide(0).unwrap().layout().unwrap().unwrap();
    let second = second_layout.scene();
    let after_hit = source.scene_cache_diagnostics();
    assert_eq!(after_hit.hits, 1);
    assert_eq!(after_hit.misses, 1);
    assert_eq!(after_hit.builds, 1);
    assert_eq!(after_hit.retained_entries, 1);
    let expected_len = first.len();
    assert_eq!(expected_len, second.len());

    // Held detached scenes pin the entry.  Once both callers release them,
    // the public trim operation can reclaim the clean memo entry.
    assert_eq!(source.trim_scene_cache(), 0);
    drop((first_layout, second_layout));
    assert_eq!(source.trim_scene_cache(), 1);
    let after_trim = source.scene_cache_diagnostics();
    assert_eq!(after_trim.retained_entries, 0);
    assert_eq!(after_trim.trimmed, 1);

    let rebuilt_layout = slide.layout().unwrap().unwrap();
    assert_eq!(rebuilt_layout.scene().len(), expected_len);
    assert_eq!(source.scene_cache_diagnostics().builds, 2);
}

#[test]
fn managed_source_related_scenes_are_unretained_but_live_snapshots_remain_valid() {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(rich_deck()));
    let (budget, cancellation, context) = managed_context(16 * 1024 * 1024);
    let presentation = SourceBackedPresentation::from_read_at_with_options(
        source,
        OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let slide = presentation.slide(0).unwrap();
    let first_layout = slide.layout().unwrap().unwrap();
    let first = first_layout.scene();
    let second_layout = slide.layout().unwrap().unwrap();
    let second = second_layout.scene();
    let diagnostics = presentation.scene_cache_diagnostics();
    assert_eq!(diagnostics.max_bytes, 0);
    assert_eq!(diagnostics.max_entries, 0);
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.unretained_builds, 2);
    assert_eq!(first.len(), second.len());

    cancellation.cancel();
    assert!(matches!(
        slide.layout(),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    assert!(matches!(
        slide.notes(),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    assert!(matches!(
        slide.comments(),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    assert!(matches!(
        slide.charts(),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    assert!(matches!(
        slide.chart(Position::new(0)),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    assert_eq!(first.len(), second.len());
    drop((first_layout, second_layout));
    drop((slide, presentation));
    assert_eq!(budget.used(litchi_core::Resource::Memory), 0);
}

#[test]
fn source_related_layout_applies_physical_depth_limits_before_mce_and_does_not_cache_failures() {
    let nesting = "<x:n>".repeat(100) + &"</x:n>".repeat(100);
    let xml = format!(
        r#"<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unknown" type="title" preserve="1"><p:cSld name="Depth"><p:spTree><mc:AlternateContent><mc:Choice Requires="x">{nesting}</mc:Choice><mc:Fallback/></mc:AlternateContent></p:spTree></p:cSld></p:sldLayout>"#
    );
    let bytes = rewrite(&rich_deck(), |opc| edit_blob(opc, LAYOUT, |_| xml.clone()));
    let limits = SourceLimits::builder()
        .max_depth(64)
        .unwrap()
        .build()
        .unwrap();
    let presentation = SourceBackedPresentation::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes)),
        OpenOptions::default().source_limits(limits),
    )
    .unwrap();
    let slide = presentation.slide(0).unwrap();
    for _ in 0..2 {
        assert!(matches!(
            slide.layout(),
            Err(Error::Opc(OpcError::Limit {
                kind: SourceLimitKind::Depth,
                ..
            }))
        ));
    }
    let diagnostics = presentation.scene_cache_diagnostics();
    assert_eq!(diagnostics.builds, 0);
    assert_eq!(diagnostics.retained_entries, 0);
}

#[test]
fn source_related_mce_expansion_is_bounded_before_scene_publication() {
    let namespace = format!("urn:{}", "q".repeat(2_000));
    let mut shapes = String::new();
    for index in 0..80 {
        shapes.push_str(&format!(
            r#"<p:sp q:marker="{index}"><p:nvSpPr><p:cNvPr id="{}" name="Shape {index}"/></p:nvSpPr><p:spPr/></p:sp>"#,
            index + 10
        ));
    }
    let xml = format!(
        r#"<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported" type="title"><p:cSld><p:spTree><mc:AlternateContent><mc:Choice Requires="x"><x:drop/></mc:Choice><mc:Fallback xmlns:q="{namespace}">{shapes}</mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sldLayout>"#
    );
    assert!(xml.len() < 65_536);

    let bytes = rewrite(&rich_deck(), |opc| edit_blob(opc, LAYOUT, |_| xml));
    let limits = SourceLimits::builder()
        .max_part_bytes(65_536)
        .unwrap()
        .build()
        .unwrap();
    let presentation = SourceBackedPresentation::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes)),
        OpenOptions::default().source_limits(limits),
    )
    .unwrap();
    let slide = presentation.slide(0).unwrap();
    let error = slide.layout().unwrap_err();
    assert!(matches!(
        error,
        Error::MarkupCompatibility(
            litchi_ooxml_common::mce::Error::LimitExceeded { kind, .. }
        ) if kind == litchi_ooxml_common::mce::LimitKind::OutputBytes
    ));
    let diagnostics = presentation.scene_cache_diagnostics();
    assert_eq!(diagnostics.builds, 0);
    assert_eq!(diagnostics.retained_entries, 0);
}

#[test]
fn related_reads_refuse_source_changes_while_detached_snapshots_remain_readable() {
    let source = Arc::new(VersionedSource::new(rich_deck()));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let slide = presentation.slide(0).unwrap();

    let layout = slide.layout().unwrap().unwrap();
    let layout_len = layout.scene().len();
    let notes = slide.notes().unwrap().unwrap();
    let comments = slide.comments().unwrap().unwrap();
    let charts = slide.charts().unwrap();
    let hidden = slide.is_hidden().unwrap();
    assert!(!hidden);

    source.changed();
    assert!(matches!(
        slide.layout(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    assert!(matches!(
        slide.notes(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    assert!(matches!(
        slide.comments(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    assert!(matches!(
        slide.charts(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    assert!(matches!(
        slide.chart(Position::new(0)),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    assert!(matches!(
        slide.is_hidden(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));

    assert_eq!(layout.name(), Some("Charts"));
    assert_eq!(layout.scene().len(), layout_len);
    assert_eq!(notes.text(), Some("First note"));
    assert!(!notes.scene().is_empty());
    assert_eq!(comments.entries()[0].text(), "Review this slide");
    assert_eq!(charts.len(), 1);
    assert_eq!(charts[0].kind(), Type::Bar);
    assert!(!hidden);
}

#[test]
fn post_publication_layout_fence_discards_the_published_scene() {
    let probe = Arc::new(ArmedRevision::new(rich_deck()));
    let presentation = SourceBackedPresentation::from_read_at(probe.clone()).unwrap();
    let slide = presentation.slide(0).unwrap();
    let before = probe.checks.load(Ordering::SeqCst);
    slide.layout().unwrap();
    let checks = probe.checks.load(Ordering::SeqCst) - before;
    assert!(checks > 2);

    let source = Arc::new(ArmedRevision::new(rich_deck()));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let slide = presentation.slide(0).unwrap();
    source.change_on.store(
        source.checks.load(Ordering::SeqCst) + checks,
        Ordering::SeqCst,
    );
    assert!(matches!(
        slide.layout(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    let diagnostics = presentation.scene_cache_diagnostics();
    assert_eq!(diagnostics.builds, 1);
    assert_eq!(diagnostics.fence_discards, 1);
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.retained_bytes, 0);
}

#[test]
fn post_publication_notes_fence_discards_the_published_scene() {
    let probe = Arc::new(ArmedRevision::new(rich_deck()));
    let presentation = SourceBackedPresentation::from_read_at(probe.clone()).unwrap();
    let slide = presentation.slide(0).unwrap();
    let before = probe.checks.load(Ordering::SeqCst);
    slide.notes().unwrap();
    let checks = probe.checks.load(Ordering::SeqCst) - before;
    assert!(checks > 2);

    let source = Arc::new(ArmedRevision::new(rich_deck()));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let slide = presentation.slide(0).unwrap();
    source.change_on.store(
        source.checks.load(Ordering::SeqCst) + checks,
        Ordering::SeqCst,
    );
    assert!(matches!(
        slide.notes(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    let diagnostics = presentation.scene_cache_diagnostics();
    assert_eq!(diagnostics.builds, 1);
    assert_eq!(diagnostics.fence_discards, 1);
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.retained_bytes, 0);
}

#[test]
fn malformed_legacy_comments_are_refused_by_the_strict_related_reader() {
    let malformed = rewrite(&rich_deck(), |opc| {
        edit_blob(opc, COMMENTS, |xml| {
            xml.replacen(
                "<p:text>Review this slide</p:text>",
                "<p:text><p:futureText/></p:text>",
                1,
            )
        });
    });
    let error = open(&malformed).slide(0).unwrap().comments().unwrap_err();
    assert!(!format!("{error:?}").contains("Review this slide"));
}

#[test]
fn unsupported_layout_kind_is_refused_by_the_strict_related_reader() {
    let unsupported = rewrite(&rich_deck(), |opc| {
        edit_blob(opc, LAYOUT, |xml| {
            xml.replacen("type=\"chart\"", "type=\"futureVendorKind\"", 1)
        });
    });
    assert!(open(&unsupported).slide(0).unwrap().layout().is_err());
}

#[test]
fn unknown_chart_type_is_refused_by_the_strict_related_reader() {
    let unknown_chart = br#"<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:date1904 val="0"/><c:chart><c:plotArea><c:futureChart/></c:plotArea></c:chart></c:chartSpace>"#;
    let bytes = rewrite(&rich_deck(), |opc| {
        edit_blob(opc, CHART, |_| {
            String::from_utf8(unknown_chart.to_vec()).unwrap()
        });
    });
    let slide = open(&bytes).slide(0).unwrap();
    assert!(slide.charts().is_err());
    assert!(slide.chart(Position::new(0)).is_err());
}
