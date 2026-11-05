#![cfg(feature = "rtf")]

use litchi_core::Position;
use litchi_core::selector::{Path, Scope};
use litchi_ops::error::Source;
use litchi_ops::query::Query;
use litchi_ops::transaction::{Error, LimitKind, SemanticLimits};
use litchi_ops::{BatchItem, BatchMode, Code, File, NewNode, Props, open_bytes, props};
use litchi_rtf::Alignment;
use litchi_rtf::edit::{ParagraphAlignmentUpdate, ParagraphTextReplacement};

const INPUT: &[u8] = br"{\rtf1\ansi First\par Second\par Third}";
fn file() -> File {
    open_bytes(INPUT).unwrap()
}
fn align(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|property| property.key() == "align")
        .unwrap()
        .to_string()
}
fn query(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .unwrap()
        .with_scope(Scope::Path(Box::new(Path::parse("/body").unwrap())))
}

#[test]
fn alignment_reads_and_predicates_use_the_native_paragraph_format() {
    let source =
        open_bytes(br"{\rtf1\ansi\qr First\par\qc Second\par\qj Third\par\ql Fourth}".as_slice())
            .unwrap();
    for (i, value) in ["right", "center", "justify", "left"]
        .into_iter()
        .enumerate()
    {
        assert_eq!(align(&source, &format!("/body/p[{}]", i + 1)), value);
        let matched = source
            .query(format!("p[align={value}]").as_str())
            .unwrap()
            .map(|node| node.unwrap().path().to_string())
            .collect::<Vec<_>>();
        assert_eq!(matched, [format!("/body/p[{}]", i + 1)]);
    }
    assert_eq!(align(&file(), "/body/p[1]"), "left");
    assert_eq!(
        source
            .get("/body")
            .unwrap()
            .view()
            .format()
            .unwrap()
            .count(),
        0
    );
}

#[test]
fn every_alignment_matches_native_publication_and_reopens() {
    for (name, value) in [
        ("left", Alignment::Left),
        ("center", Alignment::Center),
        ("right", Alignment::Right),
        ("justify", Alignment::Justify),
    ] {
        let source = file();
        let result = source.set("/body/p[2]", props! {"align" => name}).unwrap();
        let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
        let mut edit = native.edit();
        edit.set_body_paragraph_alignments(&[ParagraphAlignmentUpdate::new(
            Position::new(1),
            value,
        )])
        .unwrap();
        assert_eq!(
            result.to_bytes().unwrap(),
            edit.commit().unwrap().package().to_bytes().unwrap()
        );
        let reopened = open_bytes(result.to_bytes().unwrap()).unwrap();
        assert_eq!(align(&reopened, "/body/p[2]"), name);
        assert_eq!(align(&reopened, "/body/p[1]"), "left");
        assert_eq!(align(&reopened, "/body/p[3]"), "left");
        assert_eq!(source.to_bytes().unwrap(), INPUT);
    }
}

#[test]
fn selector_fanout_and_mixed_properties_match_one_native_commit() {
    let source = file();
    let commit = source
        .edit(|edit| {
            edit.set(
                query("p[align=left]"),
                props! {"align" => "right", "text" => "same"},
            )
        })
        .unwrap();
    assert_eq!(commit.diagnostics().operations(), 1);
    let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
    let mut edit = native.edit();
    edit.replace_body_paragraph_texts(
        &[0, 1, 2].map(|p| ParagraphTextReplacement::new(Position::new(p), "same")),
    )
    .unwrap();
    edit.set_body_paragraph_alignments(
        &[0, 1, 2].map(|p| ParagraphAlignmentUpdate::new(Position::new(p), Alignment::Right)),
    )
    .unwrap();
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        edit.commit().unwrap().package().to_bytes().unwrap()
    );
    assert_eq!(
        commit
            .file()
            .query("p[align=right][text=same]")
            .unwrap()
            .count(),
        3
    );
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
    let alignment_only = source
        .set(query("p"), props! {"align" => "center"})
        .unwrap();
    assert_eq!(alignment_only.query("p[align=center]").unwrap().count(), 3);
    assert_eq!(
        source
            .set(query("p[align=right]"), props! {"align" => "left"})
            .unwrap_err()
            .code(),
        Code::ZeroMatches
    );
}

#[test]
fn same_alignment_preserves_exact_bytes_and_protection_policy() {
    for input in [
        INPUT,
        br"{\rtf1\ansi\readprot\enforceprot1 First\par Second}".as_slice(),
    ] {
        let source = open_bytes(input).unwrap();
        let commit = source
            .edit(|edit| edit.set("/body/p[1]", props! {"align" => "left"}))
            .unwrap();
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 1);
        assert_eq!(commit.file().to_bytes().unwrap(), input);
    }
    let protected = open_bytes(br"{\rtf1\ansi\readprot\enforceprot1 First}".as_slice()).unwrap();
    let error = protected
        .set("/body/p[1]", props! {"align"=>"center"})
        .unwrap_err();
    assert_eq!(error.code(), Code::DocumentProtected);
    assert!(matches!(error.leaf(), Source::RtfEdit(_)));
}

#[test]
fn malformed_and_unsupported_properties_fail_without_partial_publication() {
    let source = file();
    for properties in [
        Props::new().with("align", "diagonal"),
        Props::new().with("align", true),
        Props::new().with("align", 42),
    ] {
        assert_eq!(
            source.set("/body/p[1]", properties).unwrap_err().code(),
            Code::InvalidValue
        );
    }
    assert!(
        source
            .set("/body/p[1]", props! {"align"=>"center","bold"=>true})
            .is_err()
    );
    assert!(
        source
            .set(
                "/body/p[1]",
                props! {"align"=>"center","text"=>"two\nparagraphs"}
            )
            .is_err()
    );
    assert!(
        source
            .add("/body", NewNode::new("p").prop("align", "center"))
            .is_err()
    );
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn mixed_assignment_budgets_count_both_native_facets() {
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
                |edit| edit.set(query("p"), props! {"text"=>"updated","align"=>"center"}),
            )
            .unwrap_err();
        assert!(
            matches!(error.leaf(),Source::Transaction(Error::Limit{kind:actual,..}) if *actual==kind),
            "{error:?}"
        );
    }
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn alignment_batches_keep_error_indices_and_sequential_queries() {
    let command = |path: &'static str, align: &'static str| {
        let mut item = BatchItem::new("set");
        item.path = Some(path.into());
        item.props = Some(Props::new().with("align", align));
        item
    };
    let commands = || {
        [
            command("/body/p[1]", "right"),
            command("/body/p[2]", "bad"),
            command("/body/p[3]", "center"),
        ]
    };
    let source = file();
    assert_eq!(
        source
            .batch(commands(), BatchMode::Atomic)
            .unwrap_err()
            .item(),
        Some(1)
    );
    let best = source.batch(commands(), BatchMode::BestEffort).unwrap();
    assert_eq!(align(best.file(), "/body/p[1]"), "right");
    assert_eq!(align(best.file(), "/body/p[2]"), "left");
    assert_eq!(align(best.file(), "/body/p[3]"), "center");
    let sequential = source
        .edit(|edit| {
            edit.set("/body/p[1]", props! {"align"=>"right"})?;
            edit.set(query("p[align=right]"), props! {"align"=>"justify"})
        })
        .unwrap();
    assert_eq!(align(sequential.file(), "/body/p[1]"), "justify");
    assert_eq!(
        sequential.file().query("p[align=justify]").unwrap().count(),
        1
    );
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn changed_alignment_retains_origin_for_atomic_save() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("align.rtf");
    std::fs::write(&path, INPUT).unwrap();
    litchi_ops::open(&path)
        .unwrap()
        .set("/body/p[2]", props! {"align"=>"justify"})
        .unwrap()
        .save()
        .unwrap();
    assert_eq!(
        align(&litchi_ops::open(&path).unwrap(), "/body/p[2]"),
        "justify"
    );
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
}
