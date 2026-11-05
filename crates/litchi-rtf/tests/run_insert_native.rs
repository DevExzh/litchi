#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_core::Position;
use litchi_rtf::Package;
use litchi_rtf::edit::{Error, History, HistoryLimits, RunFormatPatch, RunTextInsertion};
use litchi_rtf::opaque::Anchor;

const SOURCE: &[u8] = br"{\rtf1\ansi{\fonttbl{\f0\fnil Example Sans;}}{\colortbl;\red255\green0\blue0;}\f0\cf1\i\ul\fs24\strike\striked\v\caps\scaps\outl Plain{\b Bold}\par Tail}";

fn insertion(paragraph: usize, run: usize, text: &str) -> RunTextInsertion {
    RunTextInsertion::new(Position::new(paragraph), Position::new(run), text)
}

fn paragraph_text(package: &Package, paragraph: usize) -> String {
    package
        .body()
        .paragraphs()
        .nth(paragraph)
        .unwrap()
        .runs()
        .map(|run| run.text())
        .collect()
}

fn body_anchors(package: &Package) -> Vec<usize> {
    package
        .opaque()
        .iter()
        .filter_map(|node| match node.anchor() {
            Anchor::Body(position) => Some(position),
            Anchor::Structural { .. } => None,
        })
        .collect()
}

#[test]
fn native_run_insertion_inherits_full_neighbor_format_and_applies_patch() {
    let source = Package::from_bytes(SOURCE).unwrap();
    let source_runs = source
        .body()
        .paragraphs()
        .next()
        .unwrap()
        .runs()
        .collect::<Vec<_>>();
    assert_eq!(source_runs.len(), 2);

    let mut edit = source.edit();
    edit.insert_body_runs(&[
        insertion(0, 0, "start"),
        insertion(0, 1, "middle").with_format(RunFormatPatch::new().with_bold(false)),
        insertion(0, 2, "end"),
    ])
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        paragraph_text(commit.package(), 0),
        "startPlainmiddleBoldend"
    );

    let runs = commit
        .package()
        .body()
        .paragraphs()
        .next()
        .unwrap()
        .runs()
        .collect::<Vec<_>>();
    let start = runs.iter().find(|run| run.text() == "start").unwrap();
    let middle = runs.iter().find(|run| run.text() == "middle").unwrap();
    let end = runs.iter().find(|run| run.text() == "end").unwrap();
    assert_eq!(start.format().bold(), source_runs[0].format().bold());
    assert_eq!(start.format().italic(), source_runs[0].format().italic());
    assert_eq!(
        start.format().underline(),
        source_runs[0].format().underline()
    );
    assert_eq!(start.format().size(), source_runs[0].format().size());
    assert_eq!(start.format().font(), source_runs[0].format().font());
    assert_eq!(
        start.format().foreground(),
        source_runs[0].format().foreground()
    );
    assert_eq!(start.format().strike(), source_runs[0].format().strike());
    assert_eq!(
        start.format().double_strike(),
        source_runs[0].format().double_strike()
    );
    assert_eq!(start.format().hidden(), source_runs[0].format().hidden());
    assert_eq!(
        start.format().all_caps(),
        source_runs[0].format().all_caps()
    );
    assert_eq!(
        start.format().small_caps(),
        source_runs[0].format().small_caps()
    );
    assert_eq!(start.format().outline(), source_runs[0].format().outline());
    assert_eq!(end.format().bold(), source_runs[1].format().bold());
    assert_eq!(end.format().italic(), source_runs[1].format().italic());
    assert_eq!(
        end.format().underline(),
        source_runs[1].format().underline()
    );
    assert_eq!(end.format().size(), source_runs[1].format().size());
    assert_eq!(end.format().font(), source_runs[1].format().font());
    assert_eq!(
        end.format().foreground(),
        source_runs[1].format().foreground()
    );
    assert_eq!(end.format().strike(), source_runs[1].format().strike());
    assert_eq!(
        end.format().double_strike(),
        source_runs[1].format().double_strike()
    );
    assert_eq!(end.format().hidden(), source_runs[1].format().hidden());
    assert_eq!(end.format().all_caps(), source_runs[1].format().all_caps());
    assert_eq!(
        end.format().small_caps(),
        source_runs[1].format().small_caps()
    );
    assert_eq!(end.format().outline(), source_runs[1].format().outline());
    assert!(!middle.format().bold());
    assert!(middle.format().italic());
    assert_eq!(middle.format().font(), source_runs[1].format().font());
    assert_eq!(
        middle.format().foreground(),
        source_runs[1].format().foreground()
    );
    assert_eq!(middle.format().size(), source_runs[1].format().size());
    assert_eq!(middle.format().strike(), source_runs[1].format().strike());
    assert_eq!(
        middle.format().double_strike(),
        source_runs[1].format().double_strike()
    );
    assert_eq!(middle.format().hidden(), source_runs[1].format().hidden());
    assert_eq!(
        middle.format().all_caps(),
        source_runs[1].format().all_caps()
    );
    assert_eq!(
        middle.format().small_caps(),
        source_runs[1].format().small_caps()
    );
    assert_eq!(middle.format().outline(), source_runs[1].format().outline());
}

#[test]
fn native_run_insertion_supports_empty_paragraph_defaults() {
    for (source_bytes, expected_bold, expected_size) in [
        (br"{\rtf1\ansi\par}".as_slice(), false, 24),
        (br"{\rtf1\ansi\b\fs28\par}".as_slice(), true, 28),
    ] {
        let source = Package::from_bytes(source_bytes).unwrap();
        assert_eq!(source.paragraph_count(), 1);
        assert_eq!(source.body().paragraphs().next().unwrap().runs().count(), 0);

        let mut edit = source.edit();
        edit.insert_run(
            Position::new(0),
            Position::new(0),
            "empty",
            RunFormatPatch::new(),
        )
        .unwrap();
        let commit = edit.commit().unwrap();
        let run = commit
            .package()
            .body()
            .paragraphs()
            .next()
            .unwrap()
            .runs()
            .next()
            .unwrap();
        assert_eq!(run.text(), "empty");
        assert_eq!(run.format().bold(), expected_bold);
        assert_eq!(run.format().size().get(), expected_size);
        assert!(!run.format().italic());
        assert_eq!(
            run.format().underline(),
            litchi_rtf::UnderlineStyle::default()
        );
        assert!(!run.format().strike());
        assert!(!run.format().double_strike());
        assert!(!run.format().hidden());
        assert!(!run.format().all_caps());
        assert!(!run.format().small_caps());
        assert!(!run.format().outline());
        assert_eq!(
            run.format().baseline(),
            litchi_rtf::CharacterBaseline::Normal
        );
        assert_eq!(
            run.format().expansion(),
            litchi_rtf::CharacterExpansion::None
        );
        assert_eq!(run.format().horizontal_scale_percent(), 100);
        assert_eq!(run.format().kerning_half_points(), 0);
    }

    let source = Package::from_bytes(br"{\rtf1\ansi A\par\par B}").unwrap();
    assert_eq!(source.paragraph_count(), 3);
    let mut edit = source.edit();
    edit.insert_run(
        Position::new(1),
        Position::new(0),
        "empty",
        RunFormatPatch::new(),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(paragraph_text(commit.package(), 0), "A");
    assert_eq!(paragraph_text(commit.package(), 1), "empty");
    assert_eq!(paragraph_text(commit.package(), 2), "B");
}

#[test]
fn native_run_insertion_keeps_unicode_input_as_one_run() {
    let source = Package::from_bytes(br"{\rtf1\ansi\uc1 \u233?{\b B}C}").unwrap();
    let text = "雪 😀 \\ {}\tX";
    let mut edit = source.edit();
    edit.insert_run(
        Position::new(0),
        Position::new(1),
        text,
        RunFormatPatch::new().with_italic(true),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(paragraph_text(commit.package(), 0), format!("é{text}BC"));
    let inserted = commit
        .package()
        .body()
        .paragraphs()
        .next()
        .unwrap()
        .runs()
        .find(|run| run.text() == text)
        .unwrap();
    assert!(inserted.format().italic());
}

#[test]
fn native_run_insertion_equal_boundaries_are_stable_and_atomic() {
    let source = Package::from_bytes(br"{\rtf1\ansi A}").unwrap();
    let mut edit = source.edit();
    edit.insert_body_runs(&[insertion(0, 0, "x"), insertion(0, 0, "y")])
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(paragraph_text(commit.package(), 0), "xyA");

    let mut staged = source.edit();
    staged
        .insert_run(
            Position::new(0),
            Position::new(0),
            "x",
            RunFormatPatch::new(),
        )
        .unwrap()
        .insert_run(
            Position::new(0),
            Position::new(0),
            "y",
            RunFormatPatch::new(),
        )
        .unwrap();
    assert_eq!(paragraph_text(staged.commit().unwrap().package(), 0), "xyA");

    let mut invalid = source.edit();
    assert!(matches!(
        invalid.insert_body_runs(&[insertion(0, 0, "valid"), insertion(0, 99, "invalid"),]),
        Err(Error::RunInsertionOutOfRange {
            paragraph: 0,
            run: 99,
            ..
        })
    ));
    assert_eq!(invalid.operation_count(), 0);

    let mut empty = source.edit();
    assert!(matches!(
        empty.insert_body_runs(&[]),
        Err(Error::EmptyRunInsertionBatch)
    ));
    assert!(matches!(
        empty.insert_run(
            Position::new(0),
            Position::new(0),
            "",
            RunFormatPatch::new()
        ),
        Err(Error::EmptyRunTextInsertion {
            paragraph: 0,
            run: 0
        })
    ));
    assert!(matches!(
        empty.insert_body_runs(&[insertion(0, 1, "later"), insertion(0, 0, "earlier")]),
        Err(Error::RunInsertionBatchOutOfOrder { .. })
    ));
    assert_eq!(empty.operation_count(), 0);
}

#[test]
fn native_run_insertion_projects_opaque_anchors_by_source_tie_order() {
    let source_bytes =
        br"{\rtf1\ansi{\*\before opaque}A{\i B}{\*\middle opaque}C{\*\after opaque}}";
    let source = Package::from_bytes(source_bytes).unwrap();
    assert_eq!(body_anchors(&source), [0, 2, 3]);

    for (boundary, expected) in [(0, vec![0, 3, 4]), (2, vec![0, 2, 4]), (3, vec![0, 2, 4])] {
        let mut edit = source.edit();
        edit.insert_run(
            Position::new(0),
            Position::new(boundary),
            "X",
            RunFormatPatch::new(),
        )
        .unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(body_anchors(commit.package()), expected);
        assert!(commit.package().opaque().iter().all(|node| {
            source.opaque().iter().any(|original| {
                original.kind() == node.kind() && original.source() == node.source()
            })
        }));
    }
}

#[test]
fn native_run_insertion_has_exact_inverse_history_and_durable_refusal() {
    let source = Package::from_bytes(SOURCE).unwrap();
    let source_bytes = source.to_bytes().unwrap();
    let mut edit = source.edit();
    edit.insert_run(
        Position::new(0),
        Position::new(1),
        "inserted",
        RunFormatPatch::new(),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source_bytes
    );
    assert!(
        commit
            .patch()
            .to_durable(litchi_core::patch::PatchLimits::new(
                litchi_core::patch::BlobLimits::new(0, 0, 0),
                1024 * 1024,
                10,
                8,
                256 * 1024,
                512 * 1024,
            ))
            .is_err()
    );

    let mut history = History::new(source.clone(), HistoryLimits::new(2, 1024 * 1024));
    let mut history_edit = history.current().edit();
    history_edit
        .insert_run(
            Position::new(0),
            Position::new(1),
            "inserted",
            RunFormatPatch::new(),
        )
        .unwrap();
    let history_commit = history.commit(history_edit).unwrap();
    let changed_bytes = history_commit.package().to_bytes().unwrap();
    assert!(history.undo());
    assert_eq!(history.current().to_bytes().unwrap(), source_bytes);
    assert!(history.redo());
    assert_eq!(history.current().to_bytes().unwrap(), changed_bytes);
}
