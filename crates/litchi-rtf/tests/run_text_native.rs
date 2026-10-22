#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_core::Position;
use litchi_rtf::Package;
use litchi_rtf::edit::{Error, History, HistoryLimits, RunFormatPatch, RunTextReplacement};
use litchi_rtf::opaque::Anchor;

const SOURCE: &[u8] =
    br"{\rtf1\ansi{\fonttbl{\f0\fnil Example Sans;}}{\colortbl;\red255\green0\blue0;}{\info{\title Keep title}}{\header Keep header}\f0\cf1\qr Plain{\i Italic}\b Bold\par\pard\qc Tail}";

fn replacement(paragraph: usize, run: usize, text: &str) -> RunTextReplacement {
    RunTextReplacement::new(Position::new(paragraph), Position::new(run), text)
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

#[test]
fn native_run_text_replacement_preserves_omitted_facets_and_metadata() {
    let source = Package::from_bytes(SOURCE).unwrap();
    let mut edit = source.edit();
    edit.replace_body_run_texts(&[
        replacement(0, 1, "new").with_format(RunFormatPatch::new().with_bold(true))
    ])
    .unwrap();
    let commit = edit.commit().unwrap();
    let paragraph = commit.package().body().paragraphs().next().unwrap();
    let mut runs = paragraph.runs();
    assert_eq!(runs.next().unwrap().text(), "Plain");
    let changed = runs.next().unwrap();
    assert_eq!(changed.text(), "new");
    assert!(changed.format().bold());
    assert!(changed.format().italic());
    assert_eq!(changed.format().font().unwrap().name(), "Example Sans");
    assert_eq!(
        changed.format().foreground(),
        source
            .body()
            .paragraphs()
            .next()
            .unwrap()
            .runs()
            .nth(1)
            .unwrap()
            .format()
            .foreground()
    );
    assert_eq!(runs.next().unwrap().text(), "Bold");
    assert_eq!(paragraph_text(commit.package(), 1), "Tail");
    let bytes = commit.package().to_bytes().unwrap();
    assert!(
        bytes
            .windows(b"Keep title".len())
            .any(|window| window == b"Keep title")
    );
    assert!(
        bytes
            .windows(b"Keep header".len())
            .any(|window| window == b"Keep header")
    );
}

#[test]
fn native_run_text_replacement_supports_unicode_escape_tab_and_unicode_fallback_scope() {
    for source_bytes in [
        br"{\rtf1\ansi\uc0 \u38634{\i inner} tail}".as_slice(),
        br"{\rtf1\ansi\uc2 \u233??{\i inner} tail}".as_slice(),
        br"{\rtf1\ansi\uc1 \u-10179?\u-8704?{\i inner} tail}".as_slice(),
        br"{\rtf1\ansi first{same}last}".as_slice(),
        br"{\rtf1\ansi first{\b same}last}".as_slice(),
    ] {
        let source = Package::from_bytes(source_bytes).unwrap();
        let expected_runs = source.body().runs().count();
        let mut edit = source.edit();
        let updates = source
            .body()
            .paragraphs()
            .enumerate()
            .flat_map(|(paragraph, value)| {
                value
                    .runs()
                    .enumerate()
                    .map(move |(run, _)| replacement(paragraph, run, "雪 😀 \\ {}\tX"))
            })
            .collect::<Vec<_>>();
        edit.replace_body_run_texts(&updates).unwrap();
        let changed = edit.commit().unwrap();
        let formats = source
            .body()
            .runs()
            .map(|run| (run.format().bold(), run.format().italic()))
            .collect::<Vec<_>>();
        let mut offset = 0usize;
        for run in changed.package().body().runs() {
            let expected = formats[offset / "雪 😀 \\ {}\tX".len()];
            assert_eq!((run.format().bold(), run.format().italic()), expected);
            offset += run.text().len();
        }
        assert_eq!(
            paragraph_text(changed.package(), 0),
            "雪 😀 \\ {}\tX".repeat(expected_runs)
        );
    }
}

#[test]
fn native_run_text_empty_replacement_removes_only_selected_runs() {
    let source = Package::from_bytes(SOURCE).unwrap();
    let mut edit = source.edit();
    edit.replace_body_run_texts(&[replacement(0, 1, ""), replacement(1, 0, "")])
        .unwrap();
    let changed = edit.commit().unwrap();
    assert_eq!(paragraph_text(changed.package(), 0), "PlainBold");
    assert_eq!(paragraph_text(changed.package(), 1), "");
    assert_eq!(changed.package().paragraph_count(), 2);
}

#[test]
fn native_run_text_all_empty_preserves_terminal_paragraph_axis() {
    for source_bytes in [
        br"{\rtf1\ansi One}".as_slice(),
        br"{\rtf1\ansi One\par Two}".as_slice(),
        br"{\rtf1\ansi One\par Two\par Three}".as_slice(),
    ] {
        let source = Package::from_bytes(source_bytes).unwrap();
        let paragraph_count = source.paragraph_count();
        let replacements = source
            .body()
            .paragraphs()
            .enumerate()
            .flat_map(|(paragraph, value)| {
                value
                    .runs()
                    .enumerate()
                    .map(move |(run, _)| replacement(paragraph, run, ""))
            })
            .collect::<Vec<_>>();
        let mut edit = source.edit();
        edit.replace_body_run_texts(&replacements).unwrap();
        let changed = edit.commit().unwrap();
        assert_eq!(changed.package().paragraph_count(), paragraph_count);
        assert!(
            changed
                .package()
                .body()
                .paragraphs()
                .all(|paragraph| paragraph.is_empty())
        );
    }
}

#[test]
fn native_run_text_batch_is_atomic_and_typed() {
    let source = Package::from_bytes(SOURCE).unwrap();
    let valid = replacement(0, 0, "changed");
    let invalid = replacement(1, 99, "never");
    let mut edit = source.edit();
    assert!(matches!(
        edit.replace_body_run_texts(&[valid.clone(), invalid]),
        Err(Error::RunOutOfRange {
            paragraph: 1,
            run: 99,
            ..
        })
    ));
    assert_eq!(edit.operation_count(), 0);
    assert!(matches!(
        edit.replace_body_run_texts(&[]),
        Err(Error::EmptyRunTextBatch)
    ));
    assert!(matches!(
        edit.replace_body_run_texts(&[replacement(1, 0, "a"), valid]),
        Err(Error::RunTextBatchOutOfOrder {
            previous_paragraph: 1,
            previous_run: 0,
            incoming_paragraph: 0,
            incoming_run: 0,
        })
    ));
    assert!(matches!(
        edit.replace_body_run_texts(&[replacement(0, 0, "a\nb")]),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
    ));
    assert_eq!(edit.operation_count(), 0);
}

#[test]
fn native_run_text_noop_inverse_history_and_durable_refusal_are_exact() {
    let source = Package::from_bytes(SOURCE).unwrap();
    let source_bytes = source.to_bytes().unwrap();
    let original = source
        .body()
        .paragraphs()
        .next()
        .unwrap()
        .runs()
        .nth(1)
        .unwrap();

    let mut noop_edit = source.edit();
    noop_edit
        .replace_body_run_texts(&[replacement(0, 1, original.text())])
        .unwrap();
    let noop = noop_edit.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.package().to_bytes().unwrap(), source_bytes);

    let mut changed_edit = source.edit();
    changed_edit
        .replace_body_run_texts(&[replacement(0, 1, "changed")])
        .unwrap();
    let changed = changed_edit.commit().unwrap();
    assert_eq!(
        changed
            .package()
            .apply(&changed.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source_bytes
    );
    assert!(
        changed
            .patch()
            .to_durable(litchi_core::patch::PatchLimits::new(
                litchi_core::patch::BlobLimits::new(0, 0, 0),
                1024 * 1024,
                10,
                8,
                256 * 1024,
                512 * 1024
            ))
            .is_err()
    );

    let mut history = History::new(source.clone(), HistoryLimits::new(2, 1024 * 1024));
    let mut history_edit = history.current().edit();
    history_edit
        .replace_body_run_texts(&[replacement(0, 1, "changed")])
        .unwrap();
    let history_commit = history.commit(history_edit).unwrap();
    let changed_bytes = history_commit.package().to_bytes().unwrap();
    assert!(history.undo());
    assert_eq!(history.current().to_bytes().unwrap(), source_bytes);
    assert!(history.redo());
    assert_eq!(history.current().to_bytes().unwrap(), changed_bytes);
}

#[test]
fn native_run_text_replacement_budget_preflights_whole_batch_and_accumulates() {
    let source = Package::from_bytes(SOURCE).unwrap();
    let limits = litchi_rtf::SemanticLimits::builder()
        .max_replacement_bytes(3)
        .unwrap()
        .build()
        .unwrap();
    let mut edit = source.edit_with_options(litchi_rtf::EditOptions::new().semantic_limits(limits));
    assert!(matches!(
        edit.replace_body_run_texts(&[replacement(0, 0, "aa"), replacement(0, 1, "bb")]),
        Err(Error::Limit {
            kind: litchi_rtf::SemanticLimitKind::ReplacementBytes,
            observed: 4,
            maximum: 3
        })
    ));
    assert_eq!(edit.operation_count(), 0);
    edit.replace_run_text(Position::new(0), Position::new(0), "aa")
        .unwrap();
    assert!(matches!(
        edit.replace_run_text(Position::new(0), Position::new(1), "bb"),
        Err(Error::Limit {
            kind: litchi_rtf::SemanticLimitKind::ReplacementBytes,
            observed: 4,
            maximum: 3
        })
    ));
    assert_eq!(edit.operation_count(), 1);
    edit.replace_run_text(Position::new(0), Position::new(1), "b")
        .unwrap();
    let changed = edit.commit().unwrap();
    assert_eq!(paragraph_text(changed.package(), 0), "aabBold");
    assert_eq!(source.to_bytes().unwrap(), SOURCE);
}

#[test]
fn native_run_text_replacement_projects_opaque_anchors_across_utf8_runs() {
    let source_bytes =
        br"{\rtf1\ansi{\*\before opaque}A{\i B}{\*\middle opaque}\u38634?{\b C}{\*\after opaque}}";
    let source = Package::from_bytes(source_bytes).unwrap();
    assert_eq!(source.text(), "AB雪C");
    assert_eq!(source.body().runs().count(), 4);
    assert_eq!(
        source
            .opaque()
            .iter()
            .filter_map(|node| match node.anchor() {
                Anchor::Body(position) => Some(position),
                Anchor::Structural { .. } => None,
            })
            .collect::<Vec<_>>(),
        [0, 2, 6]
    );

    let opaque_sources = source
        .opaque()
        .iter()
        .map(|node| node.source().to_vec())
        .collect::<Vec<_>>();
    let mut edit = source.edit();
    edit.replace_body_run_texts(&[
        replacement(0, 0, "alpha"),
        replacement(0, 1, "β"),
        replacement(0, 2, "水"),
        replacement(0, 3, "omega"),
    ])
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.package().text(), "alphaβ水omega");
    assert_eq!(
        commit
            .package()
            .opaque()
            .iter()
            .filter_map(|node| match node.anchor() {
                Anchor::Body(position) => Some(position),
                Anchor::Structural { .. } => None,
            })
            .collect::<Vec<_>>(),
        [0, 7, 15]
    );
    assert_eq!(
        commit
            .package()
            .opaque()
            .iter()
            .map(|node| node.source().to_vec())
            .collect::<Vec<_>>(),
        opaque_sources
    );
    let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.text(), "alphaβ水omega");
    assert_eq!(
        reopened
            .opaque()
            .iter()
            .filter_map(|node| match node.anchor() {
                Anchor::Body(position) => Some(position),
                Anchor::Structural { .. } => None,
            })
            .collect::<Vec<_>>(),
        [0, 7, 15]
    );
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
}

#[test]
fn native_run_text_removal_projects_opaque_anchors_through_terminal_empty_paragraph() {
    let source_bytes =
        br"{\rtf1\ansi{\*\before opaque}A{\i B}{\*\middle opaque}{\b C}{\*\after opaque}}";
    let source = Package::from_bytes(source_bytes).unwrap();
    assert_eq!(source.text(), "ABC");
    assert_eq!(source.body().runs().count(), 3);
    let opaque_sources = source
        .opaque()
        .iter()
        .map(|node| node.source().to_vec())
        .collect::<Vec<_>>();

    let mut edit = source.edit();
    edit.replace_body_run_texts(&[
        replacement(0, 0, ""),
        replacement(0, 1, ""),
        replacement(0, 2, ""),
    ])
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.package().paragraph_count(), 1);
    assert!(
        commit
            .package()
            .body()
            .paragraphs()
            .next()
            .unwrap()
            .is_empty()
    );
    assert_eq!(
        commit
            .package()
            .opaque()
            .iter()
            .filter_map(|node| match node.anchor() {
                Anchor::Body(position) => Some(position),
                Anchor::Structural { .. } => None,
            })
            .collect::<Vec<_>>(),
        [0, 0, 1]
    );
    assert_eq!(
        commit
            .package()
            .opaque()
            .iter()
            .map(|node| node.source().to_vec())
            .collect::<Vec<_>>(),
        opaque_sources
    );
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
}
