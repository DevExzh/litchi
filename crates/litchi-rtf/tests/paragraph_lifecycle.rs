#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_rtf::{
    Document,
    edit::{Composition, CompositionError, CompositionLimits, Error, Limits},
    transport::compress,
};
use std::io::{self, Write};

fn durable_limits() -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(0, 0, 0),
        1024 * 1024,
        1,
        8,
        256 * 1024,
        512 * 1024,
    )
}

fn source() -> Document {
    Document::parse(r"{\rtf1\ansi First\par Second\par Third\par Fourth}").unwrap()
}

fn paragraph_texts(document: &Document) -> Vec<String> {
    document
        .body()
        .paragraphs()
        .map(|paragraph| paragraph.to_text())
        .collect()
}

fn assert_complete_reopen(document: &Document, expected: &[&str]) {
    let bytes = document.to_bytes().unwrap();
    let reopened = Document::from_bytes(&bytes).unwrap();
    let expected = expected
        .iter()
        .map(|value| (*value).to_string())
        .collect::<Vec<_>>();
    assert_eq!(paragraph_texts(&reopened), expected);
    assert_eq!(reopened.paragraph_count(), expected.len());
}

#[test]
fn remove_first_middle_and_last_are_exact_reversible_and_reopen() {
    for (position, expected) in [
        (0, vec!["Second", "Third", "Fourth"]),
        (1, vec!["First", "Third", "Fourth"]),
        (3, vec!["First", "Second", "Third"]),
    ] {
        let source = source();
        let mut edit = source.edit();
        edit.remove_paragraph(position).unwrap();
        let commit = edit.commit().unwrap();

        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operation_count(), 1);
        assert_complete_reopen(commit.snapshot(), &expected);
        assert!(
            commit
                .patch()
                .apply(&source)
                .unwrap()
                .same_snapshot(commit.snapshot())
        );
        assert!(
            commit
                .patch()
                .inverse()
                .apply(commit.snapshot())
                .unwrap()
                .same_snapshot(&source)
        );

        if position == 0 {
            let durable = commit.patch().to_durable(durable_limits()).unwrap();
            let removed = source.apply_durable(&durable).unwrap();
            assert_eq!(
                removed.to_bytes().unwrap(),
                commit.snapshot().to_bytes().unwrap()
            );
            let restored = removed.apply_durable(&durable.inverse()).unwrap();
            assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
        }

        let foreign = Document::parse(r"{\rtf1\ansi Foreign}").unwrap();
        assert!(matches!(
            commit.patch().apply(&foreign),
            Err(Error::PatchConflict)
        ));
    }
}

#[test]
fn removing_the_only_paragraph_yields_an_empty_story_and_durable_inverse_restores_it() {
    let source = Document::parse(r"{\rtf1\ansi Only}").unwrap();
    let mut edit = source.edit();
    edit.remove_paragraph(0).unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(commit.snapshot().text(), "");
    assert_eq!(commit.snapshot().paragraph_count(), 0);
    assert_complete_reopen(commit.snapshot(), &[]);

    let durable = commit.patch().to_durable(durable_limits()).unwrap();
    let json = durable.to_deterministic_json().unwrap();
    let text = std::str::from_utf8(&json).unwrap();
    assert!(text.contains("paragraph.remove"));
    assert!(text.contains("paragraph.insert"));
    let decoded =
        litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
            &json,
            durable_limits(),
        )
        .unwrap();
    let removed = source.apply_durable(&decoded).unwrap();
    assert_eq!(removed.text(), "");
    let restored = removed.apply_durable(&decoded.inverse()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
    assert_complete_reopen(&restored, &["Only"]);
}

#[test]
fn move_uses_final_list_positions_in_both_directions_and_same_position_is_exact_noop() {
    let source = source();

    let mut forward = source.edit();
    forward.move_paragraph(0, 3).unwrap();
    let forward = forward.commit().unwrap();
    assert_complete_reopen(forward.snapshot(), &["Second", "Third", "Fourth", "First"]);

    let mut backward = source.edit();
    backward.move_paragraph(3, 0).unwrap();
    let backward = backward.commit().unwrap();
    assert_complete_reopen(backward.snapshot(), &["Fourth", "First", "Second", "Third"]);

    let durable = forward.patch().to_durable(durable_limits()).unwrap();
    let json = durable.to_deterministic_json().unwrap();
    assert!(
        std::str::from_utf8(&json)
            .unwrap()
            .contains("paragraph.move")
    );
    let moved = source.apply_durable(&durable).unwrap();
    assert_eq!(
        moved.to_bytes().unwrap(),
        forward.snapshot().to_bytes().unwrap()
    );
    let restored = moved.apply_durable(&durable.inverse()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
    let foreign = Document::parse(r"{\rtf1\ansi Foreign\par Source}").unwrap();
    assert!(matches!(
        foreign.apply_durable(&durable),
        Err(Error::PatchConflict)
    ));

    let mut noop = source.edit();
    noop.move_paragraph(2, 2).unwrap();
    let noop = noop.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert!(noop.snapshot().same_snapshot(&source));
    assert!(
        noop.patch()
            .to_durable(durable_limits())
            .unwrap()
            .operations()
            .is_empty()
    );
}

#[test]
fn lifecycle_selectors_limits_and_structural_composition_fail_before_publication() {
    let source = source();
    let mut invalid = source.edit();
    assert!(matches!(
        invalid.remove_paragraph(4),
        Err(Error::ParagraphOutOfRange {
            position: 4,
            count: 4
        })
    ));
    assert!(matches!(
        invalid.move_paragraph(0, 4),
        Err(Error::ParagraphOutOfRange {
            position: 4,
            count: 4
        })
    ));

    let mut bounded = source.edit_with_limits(Limits::new(0));
    assert!(matches!(
        bounded.remove_paragraph(0),
        Err(Error::OperationLimit {
            observed: 1,
            limit: 0
        })
    ));

    let mut structural = source.edit();
    structural.remove_paragraph(1).unwrap();
    assert!(matches!(
        structural.move_paragraph(2, 0),
        Err(Error::Conflict {
            existing: 0,
            incoming: 1
        })
    ));
    assert!(matches!(
        structural.replace_paragraph_text(0, "changed"),
        Err(Error::Conflict {
            existing: 0,
            incoming: 1
        })
    ));

    let mut property = source.edit();
    property
        .set_paragraph_alignment(0, litchi_rtf::Alignment::Center)
        .unwrap();
    assert!(matches!(
        property.remove_paragraph(1),
        Err(Error::StructuralPropertyConflict)
    ));
    assert_eq!(source.text(), "First\nSecond\nThird\nFourth");

    let limits = CompositionLimits::new(4, 8, 16, 8);
    let mut remove = source.edit();
    remove.remove_paragraph(0).unwrap();
    let mut reorder = source.edit();
    reorder.move_paragraph(3, 0).unwrap();
    let mut composition = Composition::new(&source, limits);
    composition
        .join(remove.into_sub_edit("remove", limits).unwrap())
        .unwrap();
    assert!(matches!(
        composition.join(reorder.into_sub_edit("move", limits).unwrap()),
        Err(CompositionError::Conflicts(conflicts)) if !conflicts.is_empty()
    ));
    assert_eq!(source.text(), "First\nSecond\nThird\nFourth");
}

#[test]
fn changed_lifecycle_refuses_opaque_formatted_ambiguous_and_nonplain_transports() {
    let cases = [
        Document::parse(r"{\rtf1\ansi One\future42 Two\par Three}").unwrap(),
        Document::parse(r"{\rtf1\ansi \b One\par Two}").unwrap(),
        Document::parse(r"{\rtf1\ansi One\line wrapped\par Two}").unwrap(),
        Document::parse(r"{\rtf1\ansi One\par\trowd\cellx1000\intbl Cell\cell\row}").unwrap(),
    ];
    for source in cases {
        let exact = source.to_bytes().unwrap();
        let mut edit = source.edit();
        edit.remove_paragraph(0).unwrap();
        assert!(matches!(edit.commit(), Err(Error::UnsupportedSource(_))));
        assert_eq!(source.to_bytes().unwrap(), exact);
    }

    let mut cp1252 = br"{\rtf1\ansi\ansicpg1252 caf".to_vec();
    cp1252.push(0xe9);
    cp1252.extend_from_slice(br"\par Two}");
    let source = Document::from_bytes(&cp1252).unwrap();
    let mut edit = source.edit();
    edit.remove_paragraph(0).unwrap();
    assert!(matches!(edit.commit(), Err(Error::UnsupportedSource(_))));
    assert_eq!(source.to_bytes().unwrap(), cp1252);

    let raw = br"{\rtf1\ansi One\par Two}";
    let compressed = compress(raw, true).unwrap();
    let source = Document::from_bytes(&compressed).unwrap();
    let mut edit = source.edit();
    edit.move_paragraph(0, 1).unwrap();
    assert!(matches!(edit.commit(), Err(Error::UnsupportedSource(_))));
    assert_eq!(source.to_bytes().unwrap(), compressed);
}

struct PartialSink {
    remaining: usize,
    bytes: Vec<u8>,
}

impl Write for PartialSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::other("injected partial sink failure"));
        }
        let accepted = input.len().min(self.remaining);
        self.bytes.extend_from_slice(&input[..accepted]);
        self.remaining -= accepted;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn partial_sink_failure_does_not_mutate_the_published_snapshot() {
    let source = source();
    let mut edit = source.edit();
    edit.remove_paragraph(1).unwrap();
    let commit = edit.commit().unwrap();
    let exact = commit.snapshot().to_bytes().unwrap();
    let mut sink = PartialSink {
        remaining: 7,
        bytes: Vec::new(),
    };

    assert!(commit.snapshot().write_to(&mut sink).is_err());
    assert!(!sink.bytes.is_empty());
    assert!(sink.bytes.len() < exact.len());
    assert_eq!(commit.snapshot().to_bytes().unwrap(), exact);
    assert_complete_reopen(commit.snapshot(), &["First", "Third", "Fourth"]);
}
