//! Native field/validation evidence and explicitly library-produced transport.
use litchi_core::Position;
use litchi_rtf::{Package, ValidationStatus};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
#[test]
fn native_form_instructions_exclude_bookmark_metadata_and_retain_complete_spans() {
    use litchi_rtf::FieldType;
    let original = native("word-legacy-form.rtf");
    let p = Package::from_bytes(&original).unwrap();
    let fields = p.fields();
    assert_eq!(fields.len(), 3);
    let expected = [
        (
            FieldType::FormText,
            "FORMTEXT",
            "FixtureText010",
            70,
            "FORM_VALUE_010",
            false,
        ),
        (
            FieldType::FormCheckbox,
            "FORMCHECKBOX",
            "FixtureCheck010",
            105,
            "",
            true,
        ),
        (
            FieldType::FormDropdown,
            "FORMDROPDOWN",
            "FixtureDropdown010",
            126,
            "",
            true,
        ),
    ];
    for (field, (kind, instruction, name, position, text, private)) in fields.iter().zip(expected) {
        assert_eq!(field.field_type, kind);
        assert_eq!(field.instruction.trim(), instruction);
        assert_eq!(field.result, text);
        assert_eq!(field.status.private, private);
        let bookmark = p.raw().bookmarks().get(name).unwrap();
        assert_eq!(bookmark.position, position);
        assert_eq!(bookmark.content, text);
        assert_eq!(&p.text()[position..position + text.len()], text);
    }
    assert_eq!(p.raw().bookmarks().bookmarks().len(), 3);
    assert_eq!(
        p.validation_report().fields().status(),
        ValidationStatus::Present
    );
    let mut output = Vec::new();
    litchi_rtf::write::Writer::new(&mut output)
        .write_document(p.raw())
        .unwrap();
    let read = Package::from_bytes(&output).unwrap();
    assert_eq!(read.text(), p.text());
    assert_eq!(read.raw().form_fields(), p.raw().form_fields());
    for (before, after) in p.fields().iter().zip(read.fields()) {
        assert_eq!(before.field_type, after.field_type);
        assert_eq!(before.status, after.status);
        assert_eq!(before.instruction.trim(), after.instruction.trim());
    }
    for before in p.raw().bookmarks().bookmarks() {
        let after = read.raw().bookmarks().get(&before.name).unwrap();
        assert_eq!(
            (after.position, &after.content),
            (before.position, &before.content)
        );
    }
    assert!(p.to_bytes().unwrap() == original);
}
fn patch_limits() -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(0, 0, 0),
        128 * 1024,
        16,
        8,
        64 * 1024,
        256 * 1024,
    )
}
#[test]
fn native_validation_counts_and_redaction_plans_keep_unresolved_scope_explicit() {
    use litchi_rtf::redact::{Mode, UnsupportedReference};
    for (name, fields, objects, pictures, forms, required) in [
        (
            "word-rtf-minimal.rtf",
            0,
            0,
            0,
            0,
            UnsupportedReference::OpaqueSyntax,
        ),
        (
            "word-tables-layout.rtf",
            2,
            0,
            0,
            0,
            UnsupportedReference::OpaqueSyntax,
        ),
        (
            "word-mailmerge-fields.rtf",
            6,
            0,
            0,
            0,
            UnsupportedReference::MailMerge,
        ),
        (
            "word-legacy-form.rtf",
            3,
            0,
            0,
            3,
            UnsupportedReference::Protection,
        ),
        (
            "word-visual-objects.rtf",
            0,
            2,
            6,
            0,
            UnsupportedReference::Object,
        ),
        (
            "word-write-reservation.rtf",
            0,
            0,
            0,
            0,
            UnsupportedReference::WriteReservation,
        ),
    ] {
        let original = native(name);
        let p = Package::from_bytes(&original).unwrap();
        let report = p.validation_report();
        assert_eq!(
            report,
            litchi_rtf::validation::ValidationReport::from_bytes(
                &original,
                litchi_rtf::OpenOptions::default()
            )
            .unwrap()
        );
        for status in [
            report.syntax().status(),
            report.root().status(),
            report.document().status(),
        ] {
            assert_eq!(status, ValidationStatus::Valid);
        }
        let c = report.counts();
        assert_eq!(
            (c.fields(), c.objects(), c.pictures(), c.form_fields()),
            (fields, objects, pictures, forms)
        );
        assert_eq!(c.source_bytes(), original.len());
        assert!(c.opaque_nodes() > 0);
        assert_eq!(
            report.fields().status(),
            if fields == 0 {
                ValidationStatus::Absent
            } else {
                ValidationStatus::Present
            }
        );
        assert_eq!(
            report.unsupported_syntax().status(),
            ValidationStatus::Present
        );
        for status in [
            report.security().status(),
            report.external_links().status(),
            report.active_content().status(),
            report.external_resolution().status(),
            report.execution().status(),
        ] {
            assert_eq!(status, ValidationStatus::Unknown);
        }
        assert!(!report.is_conservatively_clean());
        let inventory = p.external_reference_redaction_snapshot().unwrap();
        assert!(inventory.references().is_empty());
        assert!(inventory.unsupported().contains(&required));
        assert!(
            matches!(inventory.plan(Mode::Strict,patch_limits()),Err(litchi_rtf::edit::Error::Refused(r)) if r.kind()==litchi_core::edit::RefusalKind::UnsupportedContent && r.reason()=="external-reference redaction has unsupported remaining surfaces")
        );
        let commit = inventory
            .plan(Mode::BestEffort, patch_limits())
            .unwrap()
            .apply()
            .unwrap();
        assert!(commit.diagnostics().is_incomplete());
        assert_eq!(commit.diagnostics().removed_references(), 0);
        assert!(commit.diagnostics().unsupported().contains(&required));
        assert!(commit.package().to_bytes().unwrap() == original);
        assert!(p.to_bytes().unwrap() == original);
    }
}
#[test]
fn native_payload_transport_derivatives_round_trip_with_bounds_and_checksum_checks() {
    use litchi_rtf::transport::{CompressionLimits, DecompressionLimits, compress, decompress};
    // Native specimens are ordinary RTF. The LZFu/MELA frames here are
    // library-created derivatives, not evidence of native Outlook emission.
    for name in [
        "word-rtf-minimal.rtf",
        "word-bidi-eastasia.rtf",
        "word-rtf-table-grid.rtf",
    ] {
        let original = native(name);
        let source = Package::from_bytes(&original).unwrap();
        for compressed in [false, true] {
            let frame = compress(
                &original,
                compressed,
                CompressionLimits::new(2 * original.len() + 16),
            )
            .unwrap();
            assert_eq!(&frame[8..12], if compressed { b"LZFu" } else { b"MELA" });
            assert!(
                decompress(&frame, DecompressionLimits::new(original.len())).unwrap() == original
            );
            assert!(matches!(
                decompress(&frame, DecompressionLimits::new(original.len() - 1)),
                Err(litchi_rtf::Error::Limit {
                    kind: litchi_rtf::SourceLimitKind::DecompressedBytes,
                    ..
                })
            ));
            let p = Package::from_bytes(&frame).unwrap();
            assert_eq!(p.text(), source.text());
            assert_eq!(p.paragraph_count(), source.paragraph_count());
            assert_eq!(
                p.validation_report().compressed_transport().status(),
                ValidationStatus::Present
            );
            let mut file = tempfile::NamedTempFile::with_suffix(".lzfu").unwrap();
            p.write_to(&mut file).unwrap();
            assert!(std::fs::read(file.path()).unwrap() == frame);
            assert!(p.edit().commit().unwrap().package().to_bytes().unwrap() == frame);
            let mut edit = p.edit();
            edit.replace_paragraph_text(Position::new(0), "refused compressed edit")
                .unwrap();
            assert!(
                matches!(edit.commit(),Err(litchi_rtf::edit::Error::Refused(r)) if r.kind()==litchi_core::edit::RefusalKind::UnsupportedContent)
            );
            let mut bad_crc = frame.clone();
            bad_crc[12] ^= 1;
            assert!(
                matches!(decompress(&bad_crc,DecompressionLimits::new(original.len())),Err(litchi_rtf::Error::InvalidStructure(s)) if s.contains("CRC32"))
            );
        }
        assert!(source.to_bytes().unwrap() == original);
    }
}
