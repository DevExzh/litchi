#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_core::Position;
use litchi_rtf::{
    EditOptions, HeaderFooterType, OpenOptions, Package, SemanticLimitKind, SemanticLimits,
    SourceLimits,
    edit::{Error, HeaderFooterParagraph, ParagraphTextReplacement, TextSpan, TransferPlan},
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

fn replacements(values: &[(usize, &str)]) -> Vec<ParagraphTextReplacement> {
    values
        .iter()
        .map(|(position, text)| ParagraphTextReplacement::new(Position::new(*position), *text))
        .collect()
}

fn durable_limits(max_operations: usize) -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(0, 0, 0),
        1024 * 1024,
        max_operations,
        8,
        256 * 1024,
        512 * 1024,
    )
}

#[test]
fn paragraph_batch_matches_scalar_bytes_for_first_middle_and_last() {
    let source = Package::parse(r"{\rtf1\ansi First\par Second\par Third\par Fourth}").unwrap();
    let batch = replacements(&[(0, "One"), (2, "Three"), (3, "Four")]);
    assert_eq!(batch[0].position(), Position::new(0));
    assert_eq!(batch[0].replacement(), "One");

    let mut batched = source.edit();
    batched.replace_body_paragraph_texts(&batch).unwrap();
    let batched = batched.commit().unwrap();

    let mut scalar = source.edit();
    scalar
        .replace_paragraph_text(Position::new(0), "One")
        .unwrap();
    scalar
        .replace_paragraph_text(Position::new(2), "Three")
        .unwrap();
    scalar
        .replace_paragraph_text(Position::new(3), "Four")
        .unwrap();
    let scalar = scalar.commit().unwrap();

    assert_eq!(batched.package().text(), "One\nSecond\nThree\nFour");
    assert_eq!(
        batched.package().to_bytes().unwrap(),
        scalar.package().to_bytes().unwrap()
    );
    assert_eq!(batched.diagnostics().operation_count(), 3);
}

#[test]
fn paragraph_batch_preserves_inline_line_breaks_outside_selected_paragraphs() {
    let source = Package::parse(r"{\rtf1\ansi First\line second\par\par Third}").unwrap();
    let mut edit = source.edit();
    edit.replace_body_paragraph_texts(&replacements(&[(1, "Saved")]))
        .unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(commit.package().text(), "First\nsecond\nSaved\nThird");
    assert_eq!(
        commit.package().to_bytes().unwrap(),
        br"{\rtf1\ansi First\line second\par Saved\par Third}"
    );
}

#[test]
fn paragraph_batch_ascii_fanout_delimits_empty_paragraph_replacements() {
    let source = Package::parse(r"{\rtf1\ansi First\line second\par\par Third}").unwrap();
    let mut edit = source.edit();
    edit.replace_body_paragraph_texts(&replacements(&[(0, "same"), (1, "same"), (2, "same")]))
        .unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(commit.package().text(), "same\nsame\nsame");
    assert_eq!(
        commit.package().to_bytes().unwrap(),
        br"{\rtf1\ansi same\par same\par same}"
    );
}

#[test]
fn paragraph_batch_rejects_bad_shape_and_late_selector_atomically() {
    let source = Package::parse(r"{\rtf1\ansi First\par Second\par Third}").unwrap();

    let mut empty = source.edit();
    assert!(matches!(
        empty.replace_body_paragraph_texts(&[]),
        Err(Error::EmptyParagraphBatch)
    ));
    assert_eq!(empty.operation_count(), 0);

    for batch in [
        replacements(&[(2, "Three"), (1, "Two")]),
        replacements(&[(1, "Two"), (1, "Again")]),
    ] {
        let mut edit = source.edit();
        assert!(matches!(
            edit.replace_body_paragraph_texts(&batch),
            Err(Error::ParagraphBatchOutOfOrder { .. })
        ));
        assert_eq!(edit.operation_count(), 0);
        assert!(edit.commit().unwrap().package().same_snapshot(&source));
    }

    let mut late = source.edit();
    assert!(matches!(
        late.replace_body_paragraph_texts(&replacements(&[(0, "One"), (99, "Never")])),
        Err(Error::ParagraphOutOfRange {
            position: 99,
            count: 3
        })
    ));
    assert_eq!(late.operation_count(), 0);
    late.replace_paragraph_text(Position::new(1), "Changed")
        .unwrap();
    assert_eq!(
        late.commit().unwrap().package().text(),
        "First\nChanged\nThird"
    );
}

#[test]
fn paragraph_batch_preflights_operation_and_replacement_limits_atomically() {
    let source = Package::parse(r"{\rtf1\ansi First\par Second\par Third}").unwrap();

    let mut exact =
        source.edit_with_options(EditOptions::new().semantic_limits(semantic_limits(3)));
    exact
        .replace_body_paragraph_texts(&replacements(&[(0, "1"), (1, "2"), (2, "3")]))
        .unwrap();
    assert_eq!(exact.operation_count(), 3);

    let mut above =
        source.edit_with_options(EditOptions::new().semantic_limits(semantic_limits(2)));
    assert!(matches!(
        above.replace_body_paragraph_texts(&replacements(&[(0, "1"), (1, "2"), (2, "3")])),
        Err(Error::Limit {
            kind: SemanticLimitKind::Operations,
            observed: 3,
            maximum: 2
        })
    ));
    assert_eq!(above.operation_count(), 0);

    let mut existing =
        source.edit_with_options(EditOptions::new().semantic_limits(semantic_limits(2)));
    existing
        .replace_text(TextSpan::new(0, 1).unwrap(), "F")
        .unwrap();
    assert!(matches!(
        existing.replace_body_paragraph_texts(&replacements(&[(1, "Two"), (2, "Three")])),
        Err(Error::Limit {
            kind: SemanticLimitKind::Operations,
            observed: 3,
            maximum: 2
        })
    ));
    assert_eq!(existing.operation_count(), 1);

    let limited = SourceLimits::builder()
        .max_input_bytes(64)
        .expect("source-byte limit is representable")
        .build()
        .expect("source profile is consistent");
    let limited_source = Package::parse_with_options(
        r"{\rtf1 A\par B}",
        OpenOptions::new().source_limits(limited),
    )
    .unwrap();
    let mut oversized = limited_source.edit();
    assert!(matches!(
        oversized.replace_body_paragraph_texts(&replacements(&[(0, &"x".repeat(65))])),
        Err(Error::InputTooLarge {
            observed: 65,
            limit: 64
        })
    ));
    assert_eq!(oversized.operation_count(), 0);
}

#[test]
fn paragraph_batch_preserves_noops_and_conflict_rules() {
    let source = Package::parse(r"{\rtf1\ansi First\par Second\par Third}").unwrap();
    let mut mixed = source.edit();
    mixed
        .replace_body_paragraph_texts(&replacements(&[(0, "First"), (1, "Changed")]))
        .unwrap();
    let mixed = mixed.commit().unwrap();
    assert!(mixed.diagnostics().changed());
    assert_eq!(mixed.diagnostics().operation_count(), 2);
    assert_eq!(mixed.package().text(), "First\nChanged\nThird");

    let mut noop = source.edit();
    noop.replace_body_paragraph_texts(&replacements(&[(0, "First"), (1, "Second"), (2, "Third")]))
        .unwrap();
    let noop = noop.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert!(noop.package().same_snapshot(&source));
    assert_eq!(
        noop.package().to_bytes().unwrap(),
        source.to_bytes().unwrap()
    );

    let mut property = source.edit();
    property
        .set_paragraph_alignment(Position::new(0), litchi_rtf::Alignment::Center)
        .unwrap();
    assert!(matches!(
        property.replace_body_paragraph_texts(&replacements(&[(1, "Two\nInserted")])),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation && refusal.reason() == "paragraph-structure and property changes cannot compose safely"
    ));
    assert_eq!(property.operation_count(), 1);

    let mut overlap = source.edit();
    overlap
        .replace_text(TextSpan::new(6, 12).unwrap(), "changed")
        .unwrap();
    assert!(matches!(
        overlap.replace_body_paragraph_texts(&replacements(&[(1, "Two")])),
        Err(Error::Conflict {
            existing: 0,
            incoming: 1
        })
    ));
    assert_eq!(overlap.operation_count(), 1);
}

#[test]
fn paragraph_batch_is_durable_reversible_and_source_checked() {
    let source = Package::parse(r"{\rtf1\ansi First\par Second\par Third}").unwrap();
    let mut edit = source.edit();
    edit.replace_body_paragraph_texts(&replacements(&[(0, "One"), (2, "Three")]))
        .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().to_durable(durable_limits(2)).unwrap();
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

    let foreign = Package::parse(r"{\rtf1\ansi Foreign\par Source}").unwrap();
    assert!(matches!(
        foreign.apply_durable(&durable),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn protected_documents_refuse_changed_body_destination_and_root_publication() {
    let protected_body =
        Package::parse(r"{\rtf1\ansi\readprot\enforceprot1 First\par Second}").unwrap();
    let mut body = protected_body.edit();
    body.replace_body_paragraph_texts(&replacements(&[(0, "Changed")]))
        .unwrap();
    assert!(matches!(
        body.commit(),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::Protected
    ));

    let protected_destination =
        Package::parse(r"{\rtf1\ansi\allprot\enforceprot1{\header Head}Body}").unwrap();
    let mut destination = protected_destination.edit();
    destination
        .set_header_footer_text(
            HeaderFooterParagraph::new(0, HeaderFooterType::Header, 0),
            "Changed head",
        )
        .unwrap();
    assert!(matches!(
        destination.commit(),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::Protected
    ));

    let transfer_source = Package::parse(r"{\rtf1{\field{\*\fldinst PAGE}{\fldrslt 1}}}").unwrap();
    let protected_target = Package::parse(r"{\rtf1\ansi\formprot\enforceprot1 Target}").unwrap();
    let transfer = TransferPlan::field(&transfer_source, 0, &protected_target).unwrap();
    assert!(matches!(
        transfer.commit(),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::Protected
    ));
}

#[test]
fn protected_exact_noops_remain_byte_identical() {
    for source in [
        r"{\rtf1\ansi\readprot\enforceprot1 Same}",
        r"{\rtf1\ansi\revprot\enforceprot1 Same}",
        r"{\rtf1\ansi\annotprot\enforceprot1 Same}",
        r"{\rtf1\ansi\formprot\enforceprot1 Same}",
        r"{\rtf1\ansi\allprot\enforceprot1 Same}",
    ] {
        let source = Package::parse(source).unwrap();
        let exact = source.to_bytes().unwrap();
        let mut edit = source.edit();
        edit.replace_body_paragraph_texts(&replacements(&[(0, "Same")]))
            .unwrap();
        let commit = edit.commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert!(commit.package().same_snapshot(&source));
        assert_eq!(commit.package().to_bytes().unwrap(), exact);
    }

    let destination =
        Package::parse(r"{\rtf1\ansi\allprot\enforceprot1{\header Head}Same}").unwrap();
    let exact = destination.to_bytes().unwrap();
    let mut edit = destination.edit();
    edit.set_header_footer_text(
        HeaderFooterParagraph::new(0, HeaderFooterType::Header, 0),
        "Head",
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.package().same_snapshot(&destination));
    assert_eq!(commit.package().to_bytes().unwrap(), exact);
}

#[test]
fn paragraph_batch_preserves_transport_encoding_and_opaque_guards() {
    let cp1252 = Package::from_bytes(br"{\rtf1\ansi\ansicpg1252 Caf\'e9\par Suite}").unwrap();
    let mut cp1252_edit = cp1252.edit();
    cp1252_edit
        .replace_body_paragraph_texts(&replacements(&[(0, "Changed")]))
        .unwrap();
    let cp1252_commit = cp1252_edit.commit().unwrap();
    assert_eq!(cp1252_commit.package().text(), "Changed\nSuite");
    assert!(
        cp1252_commit
            .package()
            .to_bytes()
            .unwrap()
            .starts_with(br"{\rtf1\ansi\ansicpg1252 ")
    );

    let raw = br"{\rtf1\ansi First\par Second}";
    let compressed = compress(raw, true, CompressionLimits::default()).unwrap();
    let compressed = Package::from_bytes(&compressed).unwrap();
    let mut compressed_edit = compressed.edit();
    compressed_edit
        .replace_body_paragraph_texts(&replacements(&[(1, "Changed")]))
        .unwrap();
    assert!(matches!(
        compressed_edit.commit(),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
    ));

    let opaque = Package::parse(r"{\rtf1\ansi A\future42 B\par C}").unwrap();
    let exact = opaque.to_bytes().unwrap();
    let mut opaque_edit = opaque.edit();
    opaque_edit
        .replace_body_paragraph_texts(&replacements(&[(1, "Changed")]))
        .unwrap();
    assert!(matches!(
        opaque_edit.commit(),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
    ));
    assert_eq!(opaque.to_bytes().unwrap(), exact);
}

struct FailAfter {
    accepted: usize,
    limit: usize,
}

impl Write for FailAfter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.limit {
            return Err(io::Error::other("injected RTF sink failure"));
        }
        let accepted = bytes.len().min(self.limit - self.accepted);
        self.accepted += accepted;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn batch_commit_is_immutable_across_partial_sink_failure() {
    let source = Package::parse(r"{\rtf1\ansi First\par Second\par Third}").unwrap();
    let mut edit = source.edit();
    edit.replace_body_paragraph_texts(&replacements(&[(0, "One"), (2, "Three")]))
        .unwrap();
    let commit = edit.commit().unwrap();
    let exact = commit.package().to_bytes().unwrap();
    let mut sink = FailAfter {
        accepted: 0,
        limit: exact.len() / 2,
    };
    assert!(commit.package().write_to(&mut sink).is_err());
    assert_eq!(commit.package().to_bytes().unwrap(), exact);
    assert_eq!(source.text(), "First\nSecond\nThird");
}
