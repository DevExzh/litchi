#![allow(
    clippy::unwrap_used,
    reason = "focused integration tests use panic-on-failure assertions"
)]

use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits,
    Limits as BudgetLimits, ReadAt, Resource, SourceVersion,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, SourceCacheLimits, TargetMode,
};
use litchi_xlsx::tab_state::{Commit, SourceBackedEditor};
use litchi_xlsx::{Error, Package, ReadLimits, Visibility};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const TRANSITIONAL: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const TRANSITIONAL_REL_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAIN: &str = "/xl/workbook.xml";
const ONE: &str = "/xl/worksheets/sheet1.xml";
const TWO: &str = "/xl/worksheets/sheet2.xml";
const THREE: &str = "/xl/worksheets/sheet3.xml";
const MEDIA: &str = "/xl/media/exact.bin";
const CHART_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const STRICT_CHART_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
const DIALOG_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
const MACRO_REL: &str = "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";
const CHART_CT: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

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
        Ok(SourceVersion::new(
            811,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

struct FailingSink {
    accepted: usize,
    maximum: usize,
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.maximum {
            return Err(io::Error::other("injected sink failure"));
        }
        let written = bytes.len().min(self.maximum - self.accepted);
        self.accepted += written;
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Clone, Copy, Default)]
struct Fixture<'a> {
    strict: bool,
    workbook_prefix: &'a str,
    workbook_suffix: &'a str,
    one_xml: Option<&'a str>,
    three_kind: Option<(&'a str, &'a str, &'a str)>,
    signed: bool,
}

fn fixture(options: Fixture<'_>) -> Vec<u8> {
    let sml = if options.strict { STRICT } else { TRANSITIONAL };
    let rel_ns = if options.strict {
        STRICT_REL_NS
    } else {
        TRANSITIONAL_REL_NS
    };
    let worksheet_rel = if options.strict {
        rt::STRICT_WORKSHEET
    } else {
        rt::WORKSHEET
    };
    let office_rel = if options.strict {
        rt::STRICT_OFFICE_DOCUMENT
    } else {
        rt::OFFICE_DOCUMENT
    };
    let workbook = format!(
        r#"{prefix}<workbook xmlns="{sml}" xmlns:r="{rel_ns}"><bookViews><workbookView activeTab="0" firstSheet="0"/></bookViews>{workbook_prefix}<sheets><sheet name="One" sheetId="1" r:id="rIdOne"/><sheet name="Two" sheetId="2" state="hidden" r:id="rIdTwo"/><sheet name="Three" sheetId="3" r:id="rIdThree"/></sheets><calcPr calcId="7"/>{suffix}</workbook>"#,
        prefix = "<?xml version=\"1.0\"?>",
        workbook_prefix = options.workbook_prefix,
        suffix = options.workbook_suffix,
    );
    let one = options.one_xml.map_or_else(
        || {
            format!(
                r#"<worksheet xmlns="{sml}" xmlns:r="{rel_ns}"><sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews><sheetData/><drawing r:id="rIdMedia"/></worksheet>"#
            )
        },
        ToOwned::to_owned,
    );
    let two = format!(
        r#"<worksheet xmlns="{sml}"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/></worksheet>"#
    );
    let (three_rel, three_ct, three) = options.three_kind.map_or_else(
        || {
            (
                worksheet_rel,
                ct::SML_WORKSHEET,
                format!(
                    r#"<worksheet xmlns="{sml}"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/><extLst><ext uri="urn:exact">{}</ext></extLst></worksheet>"#,
                    "x".repeat(96 * 1024)
                ),
            )
        },
        |(relationship, content_type, root)| {
            (
                relationship,
                content_type,
                format!(
                    r#"<{root} xmlns="{sml}"><sheetViews><sheetView workbookViewId="0"/></sheetViews></{root}>"#
                ),
            )
        },
    );

    let mut package = OpcPackage::new();
    for (name, content_type, bytes) in [
        (MAIN, ct::SML_SHEET_MAIN, workbook.into_bytes()),
        (ONE, ct::SML_WORKSHEET, one.into_bytes()),
        (TWO, ct::SML_WORKSHEET, two.into_bytes()),
        (THREE, three_ct, three.into_bytes()),
        (
            MEDIA,
            "application/octet-stream",
            (0..128 * 1024).map(|value| (value % 251) as u8).collect(),
        ),
    ] {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(name).unwrap(),
                content_type.to_owned(),
                bytes,
            )))
            .unwrap();
    }
    let workbook_part = package.get_part_mut(&PackURI::new(MAIN).unwrap()).unwrap();
    for (relationship_type, target, id) in [
        (worksheet_rel, "worksheets/sheet1.xml", "rIdOne"),
        (worksheet_rel, "worksheets/sheet2.xml", "rIdTwo"),
        (three_rel, "worksheets/sheet3.xml", "rIdThree"),
    ] {
        workbook_part
            .rels_mut()
            .try_add_relationship(
                relationship_type.to_owned(),
                target.to_owned(),
                id.to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    package
        .get_part_mut(&PackURI::new(ONE).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/exact.bin".to_owned(),
            "rIdMedia".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package.relate_to("xl/workbook.xml", office_rel);
    if options.signed {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                b"<origin/>".to_vec(),
            )))
            .unwrap();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    }
    PackageWriter::to_bytes(&package).unwrap()
}

fn open_editor(bytes: Vec<u8>) -> SourceBackedEditor {
    SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap()
}

fn active_commit(editor: &SourceBackedEditor, name: &str) -> Commit {
    let mut edit = editor.edit().unwrap();
    assert!(edit.activate(name).unwrap());
    edit.commit().unwrap()
}

fn relationship_signatures(relationships: &litchi_opc::Relationships) -> Vec<String> {
    let mut signatures = relationships
        .iter()
        .map(|relationship| {
            format!(
                "{}\u{1f}{}\u{1f}{}\u{1f}{:?}",
                relationship.r_id(),
                relationship.reltype(),
                relationship.target_ref(),
                relationship.target_mode()
            )
        })
        .collect::<Vec<_>>();
    signatures.sort_unstable();
    signatures
}

fn part_len(bytes: &[u8], member: &str) -> u64 {
    OpcPackage::from_bytes(bytes)
        .unwrap()
        .get_part(&PackURI::new(member).unwrap())
        .unwrap()
        .blob()
        .len() as u64
}

fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "xlsx-tab-state-managed-test",
        BudgetLimits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(memory.max(1)).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    (budget, cancellation_source, context)
}

fn replace_zip_member_unchecked(source: &[u8], name: &str, replacement: &[u8]) -> Vec<u8> {
    let archive = ArchiveReader::new(source).unwrap();
    let names = archive
        .file_names()
        .map(ToOwned::to_owned)
        .collect::<Vec<_>>();
    let mut writer = StreamingArchiveWriter::new();
    for member in names {
        let bytes = if member == name {
            replacement.to_vec()
        } else {
            archive.read(&member).unwrap()
        };
        writer.write_deflated(&member, &bytes).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

#[test]
fn visibility_operations_use_one_part_and_preserve_every_unrelated_member() {
    let source = fixture(Fixture::default());
    let editor = open_editor(source.clone());
    assert_eq!(editor.cache_diagnostics().successful_loads, 0);
    let mut edit = editor.edit().unwrap();
    assert_eq!(editor.cache_diagnostics().successful_loads, 1);
    assert!(edit.show("Two").unwrap());
    assert!(!edit.show(1usize).unwrap());
    assert!(edit.very_hide("Three").unwrap());
    assert!(!edit.very_hide(2usize).unwrap());
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.diagnostics().touched_parts(), 1);
    assert_eq!(commit.diagnostics().touched_worksheets(), 0);
    assert_eq!(editor.cache_diagnostics().successful_loads, 1);

    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    let before = OpcPackage::from_bytes(&source).unwrap();
    let after = OpcPackage::from_bytes(&output).unwrap();
    for part in before.iter_parts() {
        let actual = after.get_part(part.partname()).unwrap();
        assert_eq!(actual.content_type(), part.content_type());
        assert_eq!(
            relationship_signatures(actual.rels()),
            relationship_signatures(part.rels())
        );
        if part.partname() != &PackURI::new(MAIN).unwrap() {
            assert_eq!(actual.blob(), part.blob());
        }
    }
    let workbook = Package::from_bytes(output)
        .unwrap()
        .into_workbook()
        .unwrap();
    assert_eq!(
        workbook.sheet("Two").unwrap().unwrap().visibility(),
        &Visibility::Visible
    );
    assert_eq!(
        workbook.sheet("Three").unwrap().unwrap().visibility(),
        &Visibility::VeryHidden
    );
    assert_eq!(workbook.active_sheet().unwrap().name(), "One");
}

#[test]
fn activation_and_implicit_relocation_use_three_parts_and_are_reversible() {
    for strict in [false, true] {
        let source = fixture(Fixture {
            strict,
            ..Fixture::default()
        });
        let editor = open_editor(source.clone());
        let commit = active_commit(&editor, "Three");
        assert_eq!(commit.diagnostics().touched_parts(), 3);
        assert_eq!(commit.diagnostics().touched_worksheets(), 2);
        assert_eq!(editor.cache_diagnostics().successful_loads, 3);

        let mut replay = OpcPackage::from_bytes(&source).unwrap();
        commit.patch().apply(&mut replay).unwrap();
        assert_eq!(
            litchi_xlsx::tab_state::Snapshot::load(&replay)
                .unwrap()
                .active_tab()
                .unwrap()
                .name(),
            "Three"
        );
        let one = std::str::from_utf8(replay.get_part(&PackURI::new(ONE).unwrap()).unwrap().blob())
            .unwrap();
        let three = std::str::from_utf8(
            replay
                .get_part(&PackURI::new(THREE).unwrap())
                .unwrap()
                .blob(),
        )
        .unwrap();
        assert!(!one.contains("tabSelected=\"1\""));
        assert!(three.contains("tabSelected=\"1\""));
        commit.patch().inverse().apply(&mut replay).unwrap();
        let original = OpcPackage::from_bytes(&source).unwrap();
        for name in [MAIN, ONE, THREE] {
            let uri = PackURI::new(name).unwrap();
            assert_eq!(
                replay.get_part(&uri).unwrap().blob(),
                original.get_part(&uri).unwrap().blob()
            );
        }

        let mut output = Vec::new();
        editor
            .publish_commit_to_stream(&mut output, &commit)
            .unwrap();
        let published = Package::from_bytes(output.clone())
            .unwrap()
            .into_workbook()
            .unwrap();
        assert_eq!(published.active_sheet().unwrap().name(), "Three");
        let output = OpcPackage::from_bytes(&output).unwrap();
        for name in [TWO, MEDIA] {
            let uri = PackURI::new(name).unwrap();
            assert_eq!(
                output.get_part(&uri).unwrap().blob(),
                original.get_part(&uri).unwrap().blob()
            );
        }

        let editor = open_editor(source);
        let mut edit = editor.edit().unwrap();
        assert!(edit.hide("One").unwrap());
        let commit = edit.commit().unwrap();
        assert_eq!(commit.snapshot().active_tab().unwrap().name(), "Three");
        assert_eq!(commit.diagnostics().touched_parts(), 3);
    }
}

#[test]
fn selectors_visibility_invariants_and_hidden_activation_fail_closed() {
    let editor = open_editor(fixture(Fixture::default()));
    let mut edit = editor.edit().unwrap();
    assert!(edit.show(1usize).unwrap());
    assert!(edit.activate("two").unwrap());
    assert!(edit.hide("Two").unwrap());
    assert!(matches!(
        edit.commit(),
        Err(Error::TabEditBlocked {
            reason: litchi_xlsx::TabEditBlock::NotVisible,
            ..
        })
    ));

    let editor = open_editor(fixture(Fixture::default()));
    let mut edit = editor.edit().unwrap();
    assert!(edit.hide("One").unwrap());
    assert!(edit.very_hide("Three").unwrap());
    assert!(matches!(
        edit.commit(),
        Err(Error::TabEditBlocked {
            reason: litchi_xlsx::TabEditBlock::LastVisibleTab,
            ..
        })
    ));

    let editor = open_editor(fixture(Fixture::default()));
    let mut edit = editor.edit().unwrap();
    assert!(edit.show("Two").unwrap());
    assert!(edit.activate(1usize).unwrap());
    assert!(edit.commit().unwrap().changed());
}

#[test]
fn staged_activation_can_return_to_source_and_noop_activation_does_not_pin_relocation() {
    let source = fixture(Fixture::default());
    let editor = open_editor(source.clone());
    let mut edit = editor.edit().unwrap();
    assert!(edit.activate("Three").unwrap());
    assert!(edit.activate("One").unwrap());
    assert!(!edit.is_changed());
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.diagnostics().touched_parts(), 0);

    let editor = open_editor(source);
    let mut edit = editor.edit().unwrap();
    assert!(!edit.activate("One").unwrap());
    assert!(edit.hide("One").unwrap());
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.snapshot().active_tab().unwrap().name(), "Three");
    assert_eq!(commit.diagnostics().touched_parts(), 3);
}

#[test]
fn noop_signed_source_is_exact_and_changed_signatures_are_refused() {
    let source = fixture(Fixture {
        signed: true,
        ..Fixture::default()
    });
    let editor = open_editor(source.clone());
    let snapshot = editor.snapshot().unwrap();
    let mut edit = editor.edit().unwrap();
    assert!(!edit.show("One").unwrap());
    assert!(!edit.activate("One").unwrap());
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert!(Arc::ptr_eq(
        &snapshot.workbook_source_arc().unwrap(),
        &commit.snapshot().workbook_source_arc().unwrap()
    ));
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source);

    let editor = open_editor(fixture(Fixture {
        signed: true,
        ..Fixture::default()
    }));
    let mut edit = editor.edit().unwrap();
    edit.very_hide("Three").unwrap();
    let commit = edit.commit().unwrap();
    assert!(
        editor
            .publish_commit_to_stream(Vec::new(), &commit)
            .is_err()
    );
}

#[test]
fn protection_mce_dtd_pi_owner_and_sheet_kind_refusals_are_typed_or_closed() {
    let protected = fixture(Fixture {
        workbook_prefix: r#"<workbookProtection lockStructure="1"/>"#,
        ..Fixture::default()
    });
    let editor = open_editor(protected);
    let mut edit = editor.edit().unwrap();
    edit.hide("Three").unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::TabEditBlocked {
            reason: litchi_xlsx::TabEditBlock::ProtectedWorkbook,
            ..
        })
    ));

    let mce = fixture(Fixture {
        workbook_suffix: r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><mc:Choice Requires="x15"><x15:future/></mc:Choice><mc:Fallback><extLst/></mc:Fallback></mc:AlternateContent>"#,
        ..Fixture::default()
    });
    let editor = open_editor(mce);
    let mut edit = editor.edit().unwrap();
    edit.hide("Three").unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::TabEditBlocked {
            reason: litchi_xlsx::TabEditBlock::MarkupCompatibility,
            ..
        })
    ));

    for prefix in ["<!DOCTYPE workbook>", "<?audit refuse?>"] {
        let source = fixture(Fixture {
            workbook_prefix: "",
            ..Fixture::default()
        });
        let package = OpcPackage::from_bytes(&source).unwrap();
        let workbook = package.get_part(&PackURI::new(MAIN).unwrap()).unwrap();
        let mut bytes = prefix.as_bytes().to_vec();
        bytes.extend_from_slice(workbook.blob());
        let editor = open_editor(replace_zip_member_unchecked(
            &source,
            "xl/workbook.xml",
            &bytes,
        ));
        let mut edit = editor.edit().unwrap();
        edit.hide("Three").unwrap();
        assert!(edit.commit().is_err());
    }

    let mce_sheet = format!(
        r#"<worksheet xmlns="{TRANSITIONAL}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews><sheetData/><mc:AlternateContent><mc:Choice Requires="x15"><x15:future/></mc:Choice><mc:Fallback><extLst/></mc:Fallback></mc:AlternateContent></worksheet>"#
    );
    let editor = open_editor(fixture(Fixture {
        one_xml: Some(&mce_sheet),
        ..Fixture::default()
    }));
    let mut edit = editor.edit().unwrap();
    edit.activate("Three").unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::TabEditBlocked {
            reason: litchi_xlsx::TabEditBlock::MarkupCompatibility,
            ..
        })
    ));

    for prefix in ["<!DOCTYPE worksheet>", "<?audit refuse?>"] {
        let source = fixture(Fixture::default());
        let package = OpcPackage::from_bytes(&source).unwrap();
        let worksheet = package.get_part(&PackURI::new(ONE).unwrap()).unwrap();
        let mut bytes = prefix.as_bytes().to_vec();
        bytes.extend_from_slice(worksheet.blob());
        let editor = open_editor(replace_zip_member_unchecked(
            &source,
            "xl/worksheets/sheet1.xml",
            &bytes,
        ));
        let mut edit = editor.edit().unwrap();
        edit.activate("Three").unwrap();
        assert!(edit.commit().is_err());
    }

    for (relationship, content_type, root) in [
        (CHART_REL, CHART_CT, "chartsheet"),
        (DIALOG_REL, "application/xml", "dialogsheet"),
        (MACRO_REL, "application/xml", "macrosheet"),
    ] {
        let editor = open_editor(fixture(Fixture {
            three_kind: Some((relationship, content_type, root)),
            ..Fixture::default()
        }));
        let mut edit = editor.edit().unwrap();
        assert!(edit.activate("Three").unwrap());
        assert!(matches!(
            edit.commit(),
            Err(Error::TabEditBlocked {
                reason: litchi_xlsx::TabEditBlock::MarkupCompatibility,
                ..
            })
        ));
    }

    let editor = open_editor(fixture(Fixture {
        three_kind: Some(("urn:unknown-sheet-owner", "application/xml", "futureSheet")),
        ..Fixture::default()
    }));
    assert!(editor.edit().unwrap().hide("Three").is_err());

    let strict_chart = open_editor(fixture(Fixture {
        strict: true,
        three_kind: Some((STRICT_CHART_REL, CHART_CT, "chartsheet")),
        ..Fixture::default()
    }));
    let mut edit = strict_chart.edit().unwrap();
    assert!(edit.activate("Three").unwrap());
    assert!(matches!(
        edit.commit(),
        Err(Error::TabEditBlocked {
            reason: litchi_xlsx::TabEditBlock::MarkupCompatibility,
            ..
        })
    ));
}

#[test]
fn relationship_source_version_limits_foreign_closure_and_partial_sink_are_checked() {
    let source = fixture(Fixture::default());
    let editor = open_editor(source.clone());
    let mut edit = editor.edit().unwrap();
    edit.very_hide("Three").unwrap();
    let commit = edit.commit().unwrap();

    let foreign_editor = open_editor(source.clone());
    assert!(matches!(
        foreign_editor.publish_commit_to_stream(Vec::new(), &commit),
        Err(Error::PatchConflict { .. })
    ));

    let mut foreign = OpcPackage::from_bytes(&source).unwrap();
    foreign
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            "urn:foreign".to_owned(),
            "media/exact.bin".to_owned(),
            "rIdForeign".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    assert!(matches!(
        commit.patch().apply(&mut foreign),
        Err(Error::PatchConflict { .. })
    ));

    let mut foreign_owner = OpcPackage::from_bytes(&source).unwrap();
    let owner_id = foreign_owner
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::OFFICE_DOCUMENT)
        .unwrap()
        .r_id()
        .to_owned();
    foreign_owner.relationships_mut().remove(&owner_id);
    foreign_owner
        .relationships_mut()
        .try_add_relationship(
            rt::OFFICE_DOCUMENT.to_owned(),
            "xl/workbook.xml".to_owned(),
            "rIdForeignOwner".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    assert!(matches!(
        commit.patch().apply(&mut foreign_owner),
        Err(Error::PatchConflict { .. })
    ));

    let source_handle = Arc::new(VersionedSource::new(source.clone()));
    let editor = SourceBackedEditor::from_read_at(source_handle.clone()).unwrap();
    let mut edit = editor.edit().unwrap();
    edit.hide("Three").unwrap();
    let commit = edit.commit().unwrap();
    source_handle.changed();
    assert!(
        editor
            .publish_commit_to_stream(Vec::new(), &commit)
            .is_err()
    );

    let limits = ReadLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        SourceBackedEditor::from_read_at_with_limits(
            Arc::new(VersionedSource::new(source.clone())),
            limits,
        )
        .is_err()
    );

    let editor = open_editor(source);
    let commit = active_commit(&editor, "Three");
    assert!(matches!(
        editor.publish_commit_to_stream(
            FailingSink {
                accepted: 0,
                maximum: 64,
            },
            &commit,
        ),
        Err(Error::Package(OpcError::IncompleteOutput { .. }))
    ));
}

#[test]
fn malformed_duplicate_external_and_retargeted_sheet_relationships_are_refused() {
    let source = fixture(Fixture::default());
    let mut package = OpcPackage::from_bytes(&source).unwrap();
    let rels = package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut();
    rels.remove("rIdThree");
    rels.try_add_relationship(
        rt::WORKSHEET.to_owned(),
        "worksheets/sheet1.xml".to_owned(),
        "rIdThree".to_owned(),
        TargetMode::Internal,
    )
    .unwrap();
    let duplicate = PackageWriter::to_bytes(&package).unwrap();
    assert!(open_editor(duplicate).edit().is_err());

    let mut package = OpcPackage::from_bytes(&source).unwrap();
    let rels = package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut();
    rels.remove("rIdThree");
    rels.try_add_relationship(
        rt::WORKSHEET.to_owned(),
        "https://example.invalid/sheet.xml".to_owned(),
        "rIdThree".to_owned(),
        TargetMode::External,
    )
    .unwrap();
    let external = PackageWriter::to_bytes(&package).unwrap();
    assert!(open_editor(external).edit().is_err());

    let editor = open_editor(source.clone());
    let commit = active_commit(&editor, "Three");
    let mut retargeted = OpcPackage::from_bytes(&source).unwrap();
    let rels = retargeted
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut();
    rels.remove("rIdThree");
    rels.try_add_relationship(
        rt::WORKSHEET.to_owned(),
        "worksheets/sheet2.xml".to_owned(),
        "rIdThree".to_owned(),
        TargetMode::Internal,
    )
    .unwrap();
    assert!(matches!(
        commit.patch().apply(&mut retargeted),
        Err(Error::PatchConflict { .. })
    ));

    let mut touched_relationship = OpcPackage::from_bytes(&source).unwrap();
    touched_relationship
        .get_part_mut(&PackURI::new(ONE).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            "urn:foreign-sheet-relationship".to_owned(),
            "../media/exact.bin".to_owned(),
            "rIdForeignSheet".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    assert!(matches!(
        commit.patch().apply(&mut touched_relationship),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn managed_workbook_payload_is_budgeted_streamed_and_released() {
    let bytes = fixture(Fixture::default());
    let exact = part_len(&bytes, MAIN);
    let (budget, _cancellation_source, context) = managed_context(exact);
    let editor =
        SourceBackedEditor::from_read_at_with_limits_and_cache_limits_and_execution_context(
            Arc::new(VersionedSource::new(bytes.clone())),
            ReadLimits::default(),
            SourceCacheLimits::new(usize::try_from(exact).unwrap(), 4).unwrap(),
            context,
        )
        .unwrap();
    assert!(editor.cache_diagnostics().budget_managed);
    assert_eq!(budget.used(Resource::Memory), 0);
    let mut edit = editor.edit().unwrap();
    assert_eq!(budget.used(Resource::Memory), exact);
    assert!(edit.show("Two").unwrap());
    let commit = edit.commit().unwrap();

    let mut replay = OpcPackage::from_bytes(&bytes).unwrap();
    commit.patch().apply(&mut replay).unwrap();
    assert!(matches!(
        commit.patch().inverse().apply(&mut replay),
        Err(Error::Package(OpcError::ManagedPartDataArcEscape))
    ));
    let mut output = Vec::new();
    let published = editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(published.tabs()[1].visibility(), &Visibility::Visible);
    let reopened = Package::from_bytes(output)
        .unwrap()
        .into_workbook()
        .unwrap();
    assert_eq!(
        reopened.sheet("Two").unwrap().unwrap().visibility(),
        &Visibility::Visible
    );
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_one_under_budget_fails_before_payload_retention_and_cancellation_stops_output() {
    let bytes = fixture(Fixture::default());
    let exact = part_len(&bytes, MAIN);
    let (budget, _cancellation_source, context) = managed_context(exact - 1);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes.clone())),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    assert!(matches!(
        editor.edit(),
        Err(Error::Package(OpcError::Execution(
            ExecutionError::ResourceLimit(_)
        )))
    ));
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes.clone())),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let edit = editor.edit().unwrap();
    cancellation_source.cancel();
    assert!(matches!(
        edit.commit(),
        Err(Error::Package(OpcError::Cancelled))
    ));
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes)),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut edit = editor.edit().unwrap();
    edit.show("Two").unwrap();
    let commit = edit.commit().unwrap();
    cancellation_source.cancel();
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::Cancelled))
    ));
    assert!(output.is_empty());
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}
