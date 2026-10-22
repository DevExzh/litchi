#![cfg(feature = "rtf")]

use litchi_core::Position;
use litchi_core::selector::{Path, Scope};
use litchi_ops::error::Source;
use litchi_ops::query::Query;
use litchi_ops::transaction::{Error, LimitKind, SemanticLimits};
use litchi_ops::{BatchItem, BatchMode, Code, File, NewNode, Props, open_bytes, props};
use litchi_rtf::edit::{ParagraphAlignmentUpdate, ParagraphLayoutPatch, ParagraphLayoutUpdate};

const INPUT: &[u8] = br"{\rtf1\ansi First\par Second\par Third}";
fn file() -> File {
    open_bytes(INPUT).unwrap()
}
fn property(file: &File, path: &str, name: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|property| property.key() == name)
        .unwrap()
        .to_string()
}
fn query(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .unwrap()
        .with_scope(Scope::Path(Box::new(Path::parse("/body").unwrap())))
}
fn all_props() -> Props<'static> {
    Props::new()
        .with("spaceBefore", "6pt")
        .with("spaceAfter", "3pt")
        .with("indent", "12pt")
        .with("rightIndent", "4pt")
        .with("firstLineIndent", "-2pt")
        .with("keepNext", true)
        .with("keepLines", true)
        .with("pageBreakBefore", true)
}
fn native_delta() -> ParagraphLayoutPatch {
    ParagraphLayoutPatch::new()
        .with_space_before(120)
        .with_space_after(60)
        .with_left_indent(240)
        .with_right_indent(80)
        .with_first_line_indent(-40)
        .with_keep_with_next(true)
        .with_keep_together(true)
        .with_page_break_before(true)
}

#[test]
fn every_layout_property_reads_native_values_and_drives_predicates() {
    let input =
        br"{\rtf1\ansi\sb120\sa60\li240\ri80\fi-40\keep\keepn\pagebb First\par\pard Second}";
    let source = open_bytes(input.as_slice()).unwrap();
    for (name, value, default) in [
        ("spaceBefore", "6pt", "0pt"),
        ("spaceAfter", "3pt", "0pt"),
        ("indent", "12pt", "0pt"),
        ("rightIndent", "4pt", "0pt"),
        ("firstLineIndent", "-2pt", "0pt"),
        ("keepNext", "true", "false"),
        ("keepLines", "true", "false"),
        ("pageBreakBefore", "true", "false"),
    ] {
        assert_eq!(property(&source, "/body/p[1]", name), value, "{name}");
        assert_eq!(property(&source, "/body/p[2]", name), default, "{name}");
        let paths = source
            .query(format!("p[{name}={value}]").as_str())
            .unwrap()
            .map(|node| node.unwrap().path().to_string())
            .collect::<Vec<_>>();
        assert_eq!(paths, ["/body/p[1]"], "{name}");
    }
}

#[test]
fn compound_layout_matches_native_bytes_and_preserves_unselected_paragraphs() {
    let source = file();
    let changed = source.set("/body/p[2]", all_props()).unwrap();
    let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
    let mut edit = native.edit();
    edit.patch_body_paragraph_layouts(&[ParagraphLayoutUpdate::new(
        Position::new(1),
        native_delta(),
    )])
    .unwrap();
    assert_eq!(
        changed.to_bytes().unwrap(),
        edit.commit().unwrap().package().to_bytes().unwrap()
    );
    let reopened = open_bytes(changed.to_bytes().unwrap()).unwrap();
    assert_eq!(property(&reopened, "/body/p[2]", "firstLineIndent"), "-2pt");
    assert_eq!(property(&reopened, "/body/p[1]", "spaceBefore"), "0pt");
    assert_eq!(property(&reopened, "/body/p[3]", "keepLines"), "false");
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn alignment_and_layout_fanout_share_one_native_publication_and_exact_history() {
    let source = file();
    let commit = source
        .edit(|edit| edit.set(query("p"), all_props().with("align", "center")))
        .unwrap();
    let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
    let mut edit = native.edit();
    edit.set_body_paragraph_alignments(
        &[0, 1, 2].map(|p| {
            ParagraphAlignmentUpdate::new(Position::new(p), litchi_rtf::Alignment::Center)
        }),
    )
    .unwrap();
    edit.patch_body_paragraph_layouts(
        &[0, 1, 2].map(|p| ParagraphLayoutUpdate::new(Position::new(p), native_delta())),
    )
    .unwrap();
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        edit.commit().unwrap().package().to_bytes().unwrap()
    );
    assert_eq!(
        commit
            .file()
            .query("p[keepNext=true][align=center]")
            .unwrap()
            .count(),
        3
    );
    assert_eq!(commit.diagnostics().operations(), 1);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        INPUT
    );
    assert_eq!(
        commit.patch().apply(&file()).unwrap().to_bytes().unwrap(),
        commit.file().to_bytes().unwrap()
    );
    let reopened = open_bytes(commit.file().to_bytes().unwrap()).unwrap();
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(&reopened)
            .unwrap()
            .to_bytes()
            .unwrap(),
        INPUT
    );
    assert!(matches!(
        commit.patch().apply(commit.file()).unwrap_err().leaf(),
        Source::Transaction(Error::PatchConflict)
    ));
}

#[test]
fn clearing_layout_preserves_mixed_runs_and_true_noops_keep_exact_bytes() {
    let input = br"{\rtf1\ansi\sb120\keep\b First\b0 rest\par\sb120\keep Second}";
    let source = open_bytes(input.as_slice()).unwrap();
    let commit = source
        .edit(|edit| edit.set(query("p"), props! {"spaceBefore"=>"0pt","keepLines"=>false}))
        .unwrap();
    assert_eq!(
        commit
            .file()
            .query("p[spaceBefore=0pt][keepLines=false]")
            .unwrap()
            .count(),
        2
    );
    let native = litchi_rtf::Package::from_bytes(&commit.file().to_bytes().unwrap()).unwrap();
    let runs: Vec<_> = native
        .body()
        .runs()
        .map(|run| (run.text().to_string(), run.format().bold()))
        .collect();
    assert!(
        runs.iter()
            .any(|(text, bold)| text.contains("First") && *bold)
    );
    assert!(
        runs.iter()
            .any(|(text, bold)| text.contains("rest") && !*bold)
    );
    let noop = commit
        .file()
        .edit(|edit| edit.set(query("p"), props! {"spaceBefore"=>"0pt","keepLines"=>false}))
        .unwrap();
    assert!(!noop.diagnostics().changed());
    assert_eq!(
        noop.file().to_bytes().unwrap(),
        commit.file().to_bytes().unwrap()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(noop.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        input
    );
}

#[test]
fn typed_lengths_and_flags_reject_invalid_input_and_native_text_conflicts_are_atomic() {
    let source = file();
    for properties in [
        Props::new().with("spaceBefore", "-1pt"),
        Props::new().with("spaceAfter", true),
        Props::new().with("indent", "999999999999999pt"),
        Props::new().with("keepNext", "maybe"),
    ] {
        assert_eq!(
            source.set("/body/p[1]", properties).unwrap_err().code(),
            Code::InvalidValue
        );
    }
    let changed = source
        .set(
            "/body/p[1]",
            props! {"indent"=>"1in","rightIndent"=>"-0.5in"},
        )
        .unwrap();
    assert_eq!(property(&changed, "/body/p[1]", "indent"), "72pt");
    assert_eq!(property(&changed, "/body/p[1]", "rightIndent"), "-36pt");
    let error = source
        .set(
            "/body/p[1]",
            props! {"spaceBefore"=>"2pt","text"=>"replacement"},
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert!(matches!(
        error.leaf(),
        Source::RtfEdit(litchi_rtf::edit::Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation && refusal.reason() == "paragraph-layout updates cannot compose with body text or character changes"
    ));
    assert!(
        source
            .add("/body", NewNode::new("p").prop("keepNext", true))
            .is_err()
    );
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn layout_fanout_limits_count_native_deltas_and_alignment_separately() {
    let source = file();
    for (limits, kind) in [
        (
            SemanticLimits::builder()
                .max_expanded_targets(2)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::ExpandedTargets,
        ),
        (
            SemanticLimits::builder()
                .max_operations(5)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::Operations,
        ),
        (
            SemanticLimits::builder()
                .max_staged_bytes(1)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::StagedBytes,
        ),
        (
            SemanticLimits::builder()
                .max_snapshot_bytes(1)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::SnapshotBytes,
        ),
    ] {
        let error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |edit| edit.set(query("p"), all_props().with("align", "right")),
            )
            .unwrap_err();
        assert!(
            matches!(error.leaf(),Source::Transaction(Error::Limit{kind:actual,..}) if *actual==kind),
            "{error:?}"
        );
    }
    let limits = SemanticLimits::builder()
        .max_operations(3)
        .unwrap()
        .build()
        .unwrap();
    let changed = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.set(query("p"), all_props()),
        )
        .unwrap();
    assert_eq!(
        changed.file().query("p[keepLines=true]").unwrap().count(),
        3
    );
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn batches_and_sequential_layout_predicates_preserve_atomicity() {
    let item = |value: &'static str| {
        let mut item = BatchItem::new("set");
        item.path = Some("/body/p[1]".into());
        item.props = Some(Props::new().with("spaceBefore", value));
        item
    };
    let source = file();
    assert_eq!(
        source
            .batch([item("6pt"), item("invalid")], BatchMode::Atomic)
            .unwrap_err()
            .item(),
        Some(1)
    );
    let best = source
        .batch([item("6pt"), item("invalid")], BatchMode::BestEffort)
        .unwrap();
    assert_eq!(property(best.file(), "/body/p[1]", "spaceBefore"), "6pt");
    let changed = source
        .edit(|edit| {
            edit.set("/body/p[2]", props! {"keepNext"=>true})?;
            edit.set(query("p[keepNext=true]"), props! {"pageBreakBefore"=>true})
        })
        .unwrap();
    assert_eq!(
        changed
            .file()
            .query("p[pageBreakBefore=true]")
            .unwrap()
            .count(),
        1
    );
    assert_eq!(
        property(changed.file(), "/body/p[2]", "pageBreakBefore"),
        "true"
    );
}

#[test]
fn layout_respects_protection_and_preserves_origin_for_atomic_save() {
    let input = br"{\rtf1\ansi\readprot\enforceprot1 First}";
    let protected = open_bytes(input.as_slice()).unwrap();
    let noop = protected
        .set("/body/p[1]", props! {"keepNext"=>false})
        .unwrap();
    assert_eq!(noop.to_bytes().unwrap(), input);
    assert_eq!(
        protected
            .set("/body/p[1]", props! {"keepNext"=>true})
            .unwrap_err()
            .code(),
        Code::DocumentProtected
    );
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("layout.rtf");
    std::fs::write(&path, INPUT).unwrap();
    litchi_ops::open(&path)
        .unwrap()
        .set("/body/p[2]", all_props())
        .unwrap()
        .save()
        .unwrap();
    assert_eq!(
        property(
            &litchi_ops::open(&path).unwrap(),
            "/body/p[2]",
            "spaceAfter"
        ),
        "3pt"
    );
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
}
