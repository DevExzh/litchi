#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_core::Position;
use litchi_rtf::{
    EditOptions, Package, SemanticLimitKind, SemanticLimits,
    edit::{Composition, CompositionError, CompositionLimits, Error},
    transport::{CompressionLimits, compress},
};
use std::io::{self, Write};

fn semantic_limits(max_operations: usize) -> SemanticLimits {
    SemanticLimits::builder()
        .max_operations(max_operations)
        .unwrap()
        .build()
        .unwrap()
}

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

fn source() -> Package {
    Package::parse(r"{\rtf1\ansi First\par Second\par Third\par Fourth}").unwrap()
}

fn paragraph_texts(document: &Package) -> Vec<String> {
    document
        .body()
        .paragraphs()
        .map(|paragraph| paragraph.to_text())
        .collect()
}

fn assert_complete_reopen(document: &Package, expected: &[&str]) {
    let bytes = document.to_bytes().unwrap();
    let reopened = Package::from_bytes(&bytes).unwrap();
    let expected = expected
        .iter()
        .map(|value| (*value).to_string())
        .collect::<Vec<_>>();
    assert_eq!(paragraph_texts(&reopened), expected);
    assert_eq!(reopened.paragraph_count(), expected.len());
}

#[test]
fn positioned_insert_preserves_empty_paragraphs_and_terminal_boundaries() {
    let cases = [
        (
            Package::parse(r"{\rtf1\ansi A\par B}").unwrap(),
            ["", "Hello", "café 水 🦀"].as_slice(),
        ),
        (
            Package::parse(r"{\rtf1\ansi A\par\par B}").unwrap(),
            ["", "Inserted", "", "Tail"].as_slice(),
        ),
        (
            Package::parse(r"{\rtf1\ansi A\par B\par}").unwrap(),
            ["", "Inserted", "Tail"].as_slice(),
        ),
        (
            Package::parse(r"{\rtf1\ansi }").unwrap(),
            ["", "Hello", ""].as_slice(),
        ),
    ];
    for (source, values) in cases {
        let count = source.paragraph_count();
        for (position, value) in values.iter().enumerate().take(count.saturating_add(1)) {
            let mut edit = source.edit();
            edit.insert_paragraph(Position::new(position), *value)
                .unwrap();
            let commit = edit.commit().unwrap_or_else(|error| {
                panic!(
                    "source={:?}, position={position}, value={value:?}: {error:?}",
                    source.text()
                )
            });
            let mut expected = paragraph_texts(&source);
            expected.insert(position, (*value).to_string());
            assert_eq!(paragraph_texts(commit.package()), expected);
            assert_eq!(commit.package().paragraph_count(), count + 1);
            assert_complete_reopen(
                commit.package(),
                &expected.iter().map(String::as_str).collect::<Vec<_>>(),
            );
            let restored = commit
                .package()
                .apply(&commit.patch().inverse())
                .unwrap()
                .into_package();
            assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
            let source_bound = (source.text() == "A\n\nB" && position == 1)
                || (source.text() == "A\nB\n" && position == count);
            if source_bound {
                assert!(matches!(
                    commit.patch().to_durable(durable_limits()),
                    Err(litchi_core::patch::PatchError::InvalidText { .. })
                ));
                assert!(matches!(
                    commit.patch().inverse().to_durable(durable_limits()),
                    Err(litchi_core::patch::PatchError::InvalidText { .. })
                ));
            } else if !value.is_empty() {
                let durable = commit.patch().to_durable(durable_limits()).unwrap();
                let applied = source.apply_durable(&durable).unwrap().into_package();
                assert_eq!(
                    applied.to_bytes().unwrap(),
                    commit.package().to_bytes().unwrap()
                );
                let restored = applied
                    .apply_durable(&durable.inverse())
                    .unwrap()
                    .into_package();
                assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
            }
        }
    }
}

#[test]
fn positioned_insert_into_an_empty_story_creates_a_real_empty_paragraph() {
    let source = Package::parse(r"{\rtf1\ansi }").unwrap();
    assert_eq!(source.paragraph_count(), 0);
    let mut edit = source.edit();
    edit.insert_paragraph(Position::new(0), "").unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(paragraph_texts(commit.package()), [""]);
    assert_eq!(commit.package().paragraph_count(), 1);
    assert_complete_reopen(commit.package(), &[""]);
    let durable = commit.patch().to_durable(durable_limits()).unwrap();
    let applied = source.apply_durable(&durable).unwrap().into_package();
    assert_eq!(
        applied.to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    let restored = applied
        .apply_durable(&durable.inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
}

#[test]
fn positioned_insert_rejects_protected_sources_after_staging() {
    let source = Package::parse(r"{\rtf1\ansi\readprot\enforceprot1 First\par Second}").unwrap();
    let mut edit = source.edit();
    edit.insert_paragraph(Position::new(1), "").unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::Protected
    ));
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
        edit.remove_paragraph(Position::new(position)).unwrap();
        let commit = edit.commit().unwrap();

        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operation_count(), 1);
        assert_complete_reopen(commit.package(), &expected);
        assert!(
            source
                .apply(commit.patch())
                .unwrap()
                .into_package()
                .same_snapshot(commit.package())
        );
        assert!(
            commit
                .package()
                .apply(&commit.patch().inverse())
                .unwrap()
                .into_package()
                .same_snapshot(&source)
        );

        if position == 0 {
            let durable = commit.patch().to_durable(durable_limits()).unwrap();
            let removed = source.apply_durable(&durable).unwrap().into_package();
            assert_eq!(
                removed.to_bytes().unwrap(),
                commit.package().to_bytes().unwrap()
            );
            let restored = removed
                .apply_durable(&durable.inverse())
                .unwrap()
                .into_package();
            assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
        }

        let foreign = Package::parse(r"{\rtf1\ansi Foreign}").unwrap();
        assert!(matches!(
            foreign.apply(commit.patch()),
            Err(Error::PatchConflict)
        ));
    }
}

#[test]
fn removing_the_only_paragraph_yields_an_empty_story_and_durable_inverse_restores_it() {
    let source = Package::parse(r"{\rtf1\ansi Only}").unwrap();
    let mut edit = source.edit();
    edit.remove_paragraph(Position::new(0)).unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(commit.package().text(), "");
    assert_eq!(commit.package().paragraph_count(), 0);
    assert_complete_reopen(commit.package(), &[]);

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
    let removed = source.apply_durable(&decoded).unwrap().into_package();
    assert_eq!(removed.text(), "");
    let restored = removed
        .apply_durable(&decoded.inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
    assert_complete_reopen(&restored, &["Only"]);
}

#[test]
fn move_uses_final_list_positions_in_both_directions_and_same_position_is_exact_noop() {
    let source = source();

    let mut forward = source.edit();
    forward
        .move_paragraph(Position::new(0), Position::new(3))
        .unwrap();
    let forward = forward.commit().unwrap();
    assert_complete_reopen(forward.package(), &["Second", "Third", "Fourth", "First"]);

    let mut backward = source.edit();
    backward
        .move_paragraph(Position::new(3), Position::new(0))
        .unwrap();
    let backward = backward.commit().unwrap();
    assert_complete_reopen(backward.package(), &["Fourth", "First", "Second", "Third"]);

    let durable = forward.patch().to_durable(durable_limits()).unwrap();
    let json = durable.to_deterministic_json().unwrap();
    assert!(
        std::str::from_utf8(&json)
            .unwrap()
            .contains("paragraph.move")
    );
    let moved = source.apply_durable(&durable).unwrap().into_package();
    assert_eq!(
        moved.to_bytes().unwrap(),
        forward.package().to_bytes().unwrap()
    );
    let restored = moved
        .apply_durable(&durable.inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
    let foreign = Package::parse(r"{\rtf1\ansi Foreign\par Source}").unwrap();
    assert!(matches!(
        foreign.apply_durable(&durable),
        Err(Error::PatchConflict)
    ));

    let mut noop = source.edit();
    noop.move_paragraph(Position::new(2), Position::new(2))
        .unwrap();
    let noop = noop.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert!(noop.package().same_snapshot(&source));
    assert!(
        noop.patch()
            .to_durable(durable_limits())
            .unwrap()
            .operations()
            .is_empty()
    );
}

#[test]
fn swap_is_one_atomic_native_lifecycle_operation_with_exact_inverse() {
    let source = source();
    let mut edit = source.edit();
    edit.swap_paragraphs(Position::new(0), Position::new(3))
        .unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(
        paragraph_texts(commit.package()),
        ["Fourth", "Second", "Third", "First"]
    );
    assert_eq!(commit.diagnostics().operation_count(), 1);
    assert!(commit.diagnostics().changed());
    assert_complete_reopen(commit.package(), &["Fourth", "Second", "Third", "First"]);
    let restored = commit
        .package()
        .apply(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
    assert!(matches!(
        commit.patch().to_durable(durable_limits()),
        Err(litchi_core::patch::PatchError::InvalidText {
            field: "RTF paragraph swaps are source-bound"
        })
    ));
}

#[test]
fn swap_preserves_empty_terminal_and_decoded_unicode_paragraphs() {
    for input in [
        br"{\rtf1\ansi First\par\par Third}".as_slice(),
        br"{\rtf1\ansi First\par\par Third\par}".as_slice(),
        br"{\rtf1\ansi First\par caf\'e9\par Third}".as_slice(),
    ] {
        let source = Package::from_bytes(input).unwrap();
        let expected = paragraph_texts(&source);
        let mut edit = source.edit();
        edit.swap_paragraphs(Position::new(0), Position::new(2))
            .unwrap();
        let commit = edit.commit().unwrap();
        let mut reordered = expected.clone();
        reordered.swap(0, 2);
        assert_eq!(paragraph_texts(commit.package()), reordered);
        assert_complete_reopen(
            commit.package(),
            &reordered.iter().map(String::as_str).collect::<Vec<_>>(),
        );
        let restored = commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package();
        assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
        assert!(matches!(
            commit.patch().to_durable(durable_limits()),
            Err(litchi_core::patch::PatchError::InvalidText { .. })
        ));
    }
}

#[test]
fn swap_same_position_is_noop_even_for_protected_or_unsupported_sources() {
    for source in [
        Package::parse(r"{\rtf1\ansi\readprot\enforceprot1 First\par Second}").unwrap(),
        Package::parse(r"{\rtf1\ansi \b First\par Second}").unwrap(),
    ] {
        let exact = source.to_bytes().unwrap();
        let mut edit = source.edit();
        edit.swap_paragraphs(Position::new(1), Position::new(1))
            .unwrap();
        let commit = edit.commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.package().to_bytes().unwrap(), exact);
        assert!(
            commit
                .patch()
                .to_durable(durable_limits())
                .unwrap()
                .operations()
                .is_empty()
        );
    }
}

#[test]
fn moving_empty_paragraphs_preserves_tail_and_terminal_boundaries() {
    for input in [
        br"{\rtf1\ansi First\par\par Third}".as_slice(),
        br"{\rtf1\ansi First\par\par Third\par}".as_slice(),
    ] {
        let source = Package::from_bytes(input).unwrap();
        assert_eq!(paragraph_texts(&source), ["First", "", "Third"]);
        for (position, final_position, expected) in [
            (1, 0, ["", "First", "Third"]),
            (1, 2, ["First", "Third", ""]),
            (2, 0, ["Third", "First", ""]),
        ] {
            let mut edit = source.edit();
            edit.move_paragraph(Position::new(position), Position::new(final_position))
                .unwrap();
            let commit = edit.commit().unwrap();
            assert_eq!(commit.package().paragraph_count(), expected.len());
            assert_eq!(paragraph_texts(commit.package()), expected);
            assert_complete_reopen(commit.package(), &expected);
            assert!(matches!(
                commit.patch().to_durable(durable_limits()),
                Err(litchi_core::patch::PatchError::InvalidText {
                    field: "RTF paragraph moves involving empty or terminal-boundary paragraphs are source-bound"
                })
            ));
            let restored = commit
                .package()
                .apply(&commit.patch().inverse())
                .unwrap()
                .into_package();
            assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
        }
    }
}

#[test]
fn terminal_boundary_moves_are_source_bound_even_without_empty_paragraphs() {
    let source = Package::parse(r"{\rtf1\ansi First\par Second\par Third\par}").unwrap();
    let mut edit = source.edit();
    edit.move_paragraph(Position::new(0), Position::new(2))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        paragraph_texts(commit.package()),
        ["Second", "Third", "First"]
    );
    assert!(matches!(
        commit.patch().to_durable(durable_limits()),
        Err(litchi_core::patch::PatchError::InvalidText {
            field: "RTF paragraph moves involving empty or terminal-boundary paragraphs are source-bound"
        })
    ));
}

#[test]
fn lifecycle_selectors_limits_and_structural_composition_fail_before_publication() {
    let source = source();
    let mut invalid = source.edit();
    assert!(matches!(
        invalid.remove_paragraph(Position::new(4)),
        Err(Error::ParagraphOutOfRange {
            position: 4,
            count: 4
        })
    ));
    assert!(matches!(
        invalid.move_paragraph(Position::new(0), Position::new(4)),
        Err(Error::ParagraphOutOfRange {
            position: 4,
            count: 4
        })
    ));

    let mut bounded =
        source.edit_with_options(EditOptions::new().semantic_limits(semantic_limits(0)));
    assert!(matches!(
        bounded.remove_paragraph(Position::new(0)),
        Err(Error::Limit {
            kind: SemanticLimitKind::Operations,
            observed: 1,
            maximum: 0
        })
    ));

    let mut structural = source.edit();
    structural.remove_paragraph(Position::new(1)).unwrap();
    assert!(matches!(
        structural.move_paragraph(Position::new(2), Position::new(0)),
        Err(Error::Conflict {
            existing: 0,
            incoming: 1
        })
    ));
    assert!(matches!(
        structural.replace_paragraph_text(Position::new(0), "changed"),
        Err(Error::Conflict {
            existing: 0,
            incoming: 1
        })
    ));

    let mut property = source.edit();
    property
        .set_paragraph_alignment(Position::new(0), litchi_rtf::Alignment::Center)
        .unwrap();
    assert!(matches!(
        property.remove_paragraph(Position::new(1)),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation && refusal.reason() == "paragraph-structure and property changes cannot compose safely"
    ));
    assert_eq!(source.text(), "First\nSecond\nThird\nFourth");

    let limits = CompositionLimits::new(4, 8, 16, 8);
    let mut remove = source.edit();
    remove.remove_paragraph(Position::new(0)).unwrap();
    let mut reorder = source.edit();
    reorder
        .move_paragraph(Position::new(3), Position::new(0))
        .unwrap();
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
    // Formatted paragraphs and line breaks are removed by the producer
    // splice (record 0293): the survivors read back unchanged.
    for (input, expected) in [
        (r"{\rtf1\ansi \b One\par Two}", r"{\rtf1\ansi \b Two}"),
        (
            r"{\rtf1\ansi One\line wrapped\par Two}",
            r"{\rtf1\ansi Two}",
        ),
    ] {
        let source = Package::parse(input).unwrap();
        let mut edit = source.edit();
        edit.remove_paragraph(Position::new(0)).unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(commit.package().to_bytes().unwrap(), expected.as_bytes());
        assert_eq!(source.to_bytes().unwrap(), input.as_bytes());
    }
    // A table row after the paragraph stays (record 0308).
    let table =
        Package::parse(r"{\rtf1\ansi One\par\trowd\cellx1000\intbl Cell\cell\row}").unwrap();
    let mut edit = table.edit();
    edit.remove_paragraph(Position::new(0)).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit.package().to_bytes().unwrap(),
        br"{\rtf1\ansi \trowd\cellx1000\intbl Cell\cell\row}"
    );
    // Unknown syntax stays refused.
    let cases = [Package::parse(r"{\rtf1\ansi One\future42 Two\par Three}").unwrap()];
    for source in cases {
        let exact = source.to_bytes().unwrap();
        let mut edit = source.edit();
        edit.remove_paragraph(Position::new(0)).unwrap();
        assert!(
            matches!(edit.commit(), Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent)
        );
        assert_eq!(source.to_bytes().unwrap(), exact);
    }

    let mut cp1252 = br"{\rtf1\ansi\ansicpg1252 caf".to_vec();
    cp1252.push(0xe9);
    cp1252.extend_from_slice(br"\par Two}");
    let source = Package::from_bytes(&cp1252).unwrap();
    let mut edit = source.edit();
    edit.remove_paragraph(Position::new(0)).unwrap();
    assert!(
        matches!(edit.commit(), Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent)
    );
    assert_eq!(source.to_bytes().unwrap(), cp1252);

    let raw = br"{\rtf1\ansi One\par Two}";
    let compressed = compress(raw, true, CompressionLimits::default()).unwrap();
    let source = Package::from_bytes(&compressed).unwrap();
    let mut edit = source.edit();
    edit.move_paragraph(Position::new(0), Position::new(1))
        .unwrap();
    assert!(
        matches!(edit.commit(), Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent)
    );
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
    edit.remove_paragraph(Position::new(1)).unwrap();
    let commit = edit.commit().unwrap();
    let exact = commit.package().to_bytes().unwrap();
    let mut sink = PartialSink {
        remaining: 7,
        bytes: Vec::new(),
    };

    assert!(commit.package().write_to(&mut sink).is_err());
    assert!(!sink.bytes.is_empty());
    assert!(sink.bytes.len() < exact.len());
    assert_eq!(commit.package().to_bytes().unwrap(), exact);
    assert_complete_reopen(commit.package(), &["First", "Third", "Fourth"]);
}

#[test]
fn positioned_insertion_admits_bounds_before_retaining_payload_and_stays_reusable() {
    let source = Package::new().unwrap();
    let limits = SemanticLimits::builder()
        .max_replacement_bytes(3)
        .unwrap()
        .build()
        .unwrap();
    let mut edit = source.edit_with_options(EditOptions::new().semantic_limits(limits));
    assert!(matches!(
        edit.insert_paragraph(Position::new(0), "abc"),
        Err(Error::Limit {
            kind: SemanticLimitKind::ReplacementBytes,
            ..
        })
    ));
    assert!(matches!(
        edit.insert_paragraph(Position::new(1), "ab"),
        Err(Error::ParagraphOutOfRange { .. })
    ));
    edit.insert_paragraph(Position::new(0), "ab").unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(paragraph_texts(commit.package()), ["ab"]);
    assert_eq!(commit.diagnostics().operation_count(), 1);
    assert_eq!(source.paragraph_count(), 0);
    let mut denied =
        source.edit_with_options(EditOptions::new().semantic_limits(semantic_limits(0)));
    assert!(matches!(
        denied.insert_paragraph(Position::new(0), ""),
        Err(Error::Limit {
            kind: SemanticLimitKind::Operations,
            ..
        })
    ));
    assert!(!denied.commit().unwrap().diagnostics().changed());
}

#[test]
fn swap_admission_failures_leave_the_edit_reusable() {
    use litchi_core::Position;
    let source = source();
    let mut edit = source.edit();
    for (first, second) in [(4, 0), (0, 4), (usize::MAX, usize::MAX)] {
        assert!(matches!(
            edit.swap_paragraphs(Position::new(first), Position::new(second)),
            Err(Error::ParagraphOutOfRange { .. })
        ));
    }
    edit.swap_paragraphs(Position::new(0), Position::new(1))
        .unwrap();
    assert!(matches!(
        edit.swap_paragraphs(Position::new(1), Position::new(2)),
        Err(Error::Conflict { .. })
    ));
    let commit = edit.commit().unwrap();
    assert_eq!(
        paragraph_texts(commit.package()),
        ["Second", "First", "Third", "Fourth"]
    );
    assert_eq!(commit.diagnostics().operation_count(), 1);
    assert!(matches!(
        commit.patch().inverse().to_durable(durable_limits()),
        Err(litchi_core::patch::PatchError::InvalidText { .. })
    ));
    let mut denied =
        source.edit_with_options(EditOptions::new().semantic_limits(semantic_limits(0)));
    assert!(matches!(
        denied.swap_paragraphs(Position::new(0), Position::new(0)),
        Err(Error::Limit {
            kind: SemanticLimitKind::Operations,
            ..
        })
    ));
    assert!(!denied.commit().unwrap().diagnostics().changed());
}
