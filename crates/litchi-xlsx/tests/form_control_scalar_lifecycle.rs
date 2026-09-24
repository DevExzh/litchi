//! Batch 2–4 form-control scalar lifecycle coverage.
//!
//! These tests use the retained local design fixtures as provenance evidence.
//! They exercise the public worksheet/source-backed owners and assert package
//! preservation around the two-part properties/VML scalar closure.  Nothing in
//! this file claims acceptance by Excel or another native application.
#![allow(
    clippy::expect_used,
    reason = "focused fixture assertions panic on failure"
)]
#![allow(
    clippy::unwrap_used,
    reason = "focused fixture assertions panic on failure"
)]

use std::io::{self, Cursor, Write};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::patch::{EffectAccess, SubEditConflict};
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as BudgetLimits, ReadAt,
    Resource, SourceVersion,
};
use litchi_opc::{OpcError, OpcPackage, PackageWriter, ReadLimits};
use litchi_xlsx::form_control::{
    Checked, ControlSelector, FormControlFormula, OwnerLimits, OwnerProfile, ScalarField,
    ScalarValue, SourceBackedFormControlEditor,
};
use litchi_xlsx::{
    Change, Conflict, Error, JoinFailure, MergeChoice, MergeLimits, PackageChange,
    SourceBackedWorkbook, Workbook,
};
use serde_json::Value;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const CONTROL_PROPERTIES_MEMBER: &str = "xl/ctrlProps/ctrlProp1.xml";
const VML_MEMBER: &str = "xl/drawings/vmlDrawing1.vml";
const DRAWING_MEMBER: &str = "xl/drawings/drawing1.xml";
const WORKSHEET_MEMBER: &str = "xl/worksheets/sheet1.xml";
const WORKSHEET_RELS_MEMBER: &str = "xl/worksheets/_rels/sheet1.xml.rels";
const CORPUS: &str = include_str!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../docs/report/spec-gap-validation-evidence/xlsx-form-control-properties/native-corpus.json"
));

fn fixture_path(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/form_control_properties")
        .join(name)
}

fn fixture_bytes(name: &str) -> Vec<u8> {
    std::fs::read(fixture_path(name)).expect("local design fixture")
}

fn rewrite_member(source: &[u8], member: &str, edit: impl FnOnce(String) -> String) -> Vec<u8> {
    let archive = ArchiveReader::new(source).expect("fixture archive");
    let replacement = edit(
        String::from_utf8(
            archive
                .read(member)
                .expect("member exists in local design fixture"),
        )
        .expect("member is UTF-8 in local design fixture"),
    )
    .into_bytes();
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        let bytes = if name == member {
            replacement.clone()
        } else {
            archive.read(name).expect("copy fixture member")
        };
        writer.write_deflated(name, &bytes).expect("rewrite member");
    }
    writer.finish_to_bytes().expect("finish fixture rewrite")
}

fn add_member(source: &[u8], name: &str, bytes: &[u8]) -> Vec<u8> {
    let archive = ArchiveReader::new(source).expect("fixture archive");
    let mut writer = StreamingArchiveWriter::new();
    for member in archive.file_names() {
        writer
            .write_deflated(member, &archive.read(member).expect("copy fixture member"))
            .expect("copy archive member");
    }
    writer
        .write_deflated(name, bytes)
        .expect("add archive member");
    writer.finish_to_bytes().expect("finish added member")
}

fn signed_fixture(source: &[u8]) -> Vec<u8> {
    let source = rewrite_member(source, "[Content_Types].xml", |xml| {
        xml.replace(
            "</Types>",
            "<Override PartName=\"/_xmlsignatures/origin.sigs\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-origin\"/></Types>",
        )
    });
    let source = rewrite_member(&source, "_rels/.rels", |xml| {
        xml.replace(
            "</Relationships>",
            "<Relationship Id=\"rIdSignature\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin\" Target=\"_xmlsignatures/origin.sigs\"/></Relationships>",
        )
    });
    add_member(&source, "_xmlsignatures/origin.sigs", b"<origin/>")
}

fn member(source: &[u8], name: &str) -> Vec<u8> {
    ArchiveReader::new(source)
        .expect("fixture archive")
        .read(name)
        .expect("fixture member")
}

fn package_members(source: &[u8]) -> Vec<(String, Vec<u8>)> {
    let archive = ArchiveReader::new(source).expect("fixture archive");
    archive
        .file_names()
        .map(|name| {
            (
                name.to_owned(),
                archive.read(name).expect("read fixture member"),
            )
        })
        .collect()
}

fn assert_only_members_changed(before: &[u8], after: &[u8], allowed: &[&str]) {
    let before = package_members(before);
    let after = package_members(after);
    assert_eq!(
        before.len(),
        after.len(),
        "scalar edit changed ZIP topology"
    );
    for ((before_name, before_bytes), (after_name, after_bytes)) in before.iter().zip(after.iter())
    {
        assert_eq!(before_name, after_name, "scalar edit reordered ZIP members");
        if allowed.contains(&before_name.as_str()) {
            assert_ne!(
                before_bytes, after_bytes,
                "expected changed member {before_name}"
            );
        } else {
            assert_eq!(
                before_bytes, after_bytes,
                "unrelated member changed: {before_name}"
            );
        }
    }
}

fn corpus_entry(name: &str) -> Value {
    let corpus: Value = serde_json::from_str(CORPUS).expect("native corpus manifest");
    corpus["fixtures"]
        .as_array()
        .expect("corpus fixtures")
        .iter()
        .find(|entry| {
            entry["retained_path"]
                .as_str()
                .and_then(|path| Path::new(path).file_name())
                .and_then(|name| name.to_str())
                == Some(name)
        })
        .cloned()
        .expect("fixture is registered in local design corpus")
}

fn assert_local_provenance(name: &str, source: &[u8]) {
    let entry = corpus_entry(name);
    assert_eq!(
        source.len(),
        entry["bytes"].as_u64().expect("fixture byte count") as usize
    );
    let expected = entry["sha256"].as_str().expect("fixture digest");
    assert_eq!(
        litchi_core::EvidenceDigest::of(source).to_string(),
        expected
    );
    for part in entry["form_control_parts"]
        .as_array()
        .expect("fixture properties manifest")
    {
        let member_name = part["member"].as_str().expect("properties member name");
        let bytes = member(source, member_name);
        assert_eq!(
            bytes.len(),
            part["bytes"].as_u64().expect("properties byte count") as usize,
            "retained properties byte count drifted for {name}:{member_name}"
        );
        assert_eq!(
            litchi_core::EvidenceDigest::of(&bytes).to_string(),
            part["sha256"].as_str().expect("properties digest"),
            "retained properties digest drifted for {name}:{member_name}"
        );
    }
}

#[derive(Debug)]
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

    fn change(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(73, self.revision.load(Ordering::SeqCst)))
    }
}

struct FailingSink {
    accepted: usize,
    limit: usize,
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.limit {
            return Err(io::Error::other("injected output refusal"));
        }
        let count = bytes.len().min(self.limit - self.accepted);
        self.accepted += count;
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn canceled_context() -> ExecutionContext {
    let (context, cancellation) = live_context();
    cancellation.cancel();
    context
}

fn live_context() -> (ExecutionContext, CancellationSource) {
    let budget = Budget::root(
        "xlsx-form-control-scalar-cancel-test",
        litchi_core::Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (source, cancellation) = CancellationSource::pair();
    let context = ExecutionContext::new(
        budget,
        cancellation,
        ExecutionLimits::new(
            std::num::NonZeroUsize::new(1).expect("non-zero depth"),
            std::num::NonZeroUsize::new(1).expect("non-zero breadth"),
            std::num::NonZeroU64::new(u64::MAX).expect("non-zero bytes"),
            0,
        )
        .expect("execution limits"),
    );
    (context, source)
}

fn managed_scalar_context(
    memory: u64,
    objects: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "xlsx-form-control-scalar-managed-test",
        BudgetLimits::new(memory, u64::MAX, u64::MAX, objects, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        std::num::NonZeroUsize::new(1).expect("non-zero depth"),
        std::num::NonZeroUsize::new(1).expect("non-zero breadth"),
        std::num::NonZeroU64::new(u64::MAX).expect("non-zero bytes"),
        0,
    )
    .expect("execution limits");
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    (budget, cancellation_source, context)
}

fn scalar_resource_usage(budget: &Budget) -> [u64; 2] {
    [
        budget.used(Resource::Memory),
        budget.used(Resource::Objects),
    ]
}

fn checked_scalar(value: Checked) -> Option<ScalarValue> {
    Some(ScalarValue::Checked(value))
}

fn source_commit(
    source: &[u8],
    field: ScalarField,
    value: Option<ScalarValue>,
) -> (Vec<u8>, litchi_xlsx::form_control::FormControlCommit) {
    let editor = SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(
        source.to_vec(),
    )))
    .expect("open source scalar editor");
    let mut edit = editor.edit("Sheet1").expect("select source worksheet");
    edit.set_scalar(ControlSelector::position(0), field, value)
        .expect("stage source scalar");
    let commit = edit.commit().expect("commit source scalar");
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .expect("publish source scalar");
    (output, commit)
}

#[test]
fn local_design_fixture_provenance_and_baseline_owner_views_are_stable() {
    for name in [
        "button-form-control.xlsx",
        "checkbox-form-control.xlsx",
        "singlecontrol.xlsx",
        "tdf134769.xlsx",
        "tdf120301_xmlSpaceParsing.xlsx",
        "tdf161365.xlsx",
        "tdf60673.xlsx",
    ] {
        let source = fixture_bytes(name);
        assert_local_provenance(name, &source);
        let workbook = Workbook::from_bytes(source.clone()).expect("eager fixture owner");
        let sheet = workbook.sheet(0).expect("sheet lookup").expect("sheet");
        assert!(!sheet.form_controls().expect("form controls").is_empty());
        let source_workbook =
            SourceBackedWorkbook::from_reader(Cursor::new(source)).expect("source owner");
        assert!(
            !source_workbook
                .sheet(0)
                .expect("source sheet lookup")
                .expect("source sheet")
                .form_controls()
                .expect("source form controls")
                .is_empty()
        );
    }
}

#[test]
fn source_snapshot_retains_profile_read_set_and_commit_guard() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let versioned = Arc::new(VersionedSource::new(source));
    let editor =
        SourceBackedFormControlEditor::from_read_at(versioned).expect("open source scalar editor");

    let snapshot = editor.snapshot("Sheet1").expect("capture source snapshot");
    assert_eq!(snapshot.sheet_name(), "Sheet1");
    assert_eq!(snapshot.sheet_position(), 0);
    assert_eq!(snapshot.form_controls().len(), 1);
    assert_eq!(snapshot.profile(), OwnerProfile::canonical());
    assert_eq!(snapshot.source_version(), Some(SourceVersion::new(73, 0)));
    let _ = snapshot.has_execution_budget();

    let read_set = snapshot.read_set().expect("source read set retained");
    assert_eq!(read_set.source_version(), snapshot.source_version());
    assert!(read_set.worksheet().is_source_backed());
    assert!(read_set.vml().is_some_and(|part| part.is_source_backed()));
    assert!(
        read_set
            .drawing()
            .is_some_and(|part| part.is_source_backed())
    );
    assert_eq!(read_set.properties().len(), 1);
    assert!(
        read_set
            .properties()
            .iter()
            .all(|part| part.is_source_backed())
    );
    let control = snapshot
        .control(ControlSelector::position(0))
        .expect("resolve retained control")
        .expect("retained control");
    assert!(control.read_set().is_some());
    assert_eq!(
        control.read_set().and_then(|set| set.source_version()),
        snapshot.source_version()
    );

    let mut edit = editor.edit("Sheet1").expect("start source edit");
    assert_eq!(edit.before().profile(), snapshot.profile());
    assert_eq!(edit.before().read_set(), snapshot.read_set());
    assert!(!edit.is_changed());
    edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::Checked,
        checked_scalar(Checked::Checked),
    )
    .expect("stage exact source no-op");
    assert!(edit.is_changed());
    let commit = edit.commit().expect("commit source no-op");
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().profile(), OwnerProfile::canonical());
    assert_eq!(commit.snapshot().read_set(), snapshot.read_set());
    assert_eq!(
        commit.patch().before().source_version(),
        snapshot.source_version()
    );
    assert_eq!(
        commit.patch().after().source_version(),
        snapshot.source_version()
    );
    assert!(commit.patch().inverse().is_empty());

    let (committed_snapshot, patch) = commit.into_parts();
    assert_eq!(committed_snapshot.sheet_position(), 0);
    assert!(patch.is_empty());
}

#[test]
fn source_scalar_changes_properties_and_vml_as_one_write_set() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let (output, commit) = source_commit(
        &source,
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    );
    assert!(commit.changed());
    assert!(!commit.patch().is_empty());
    assert_only_members_changed(&source, &output, &[CONTROL_PROPERTIES_MEMBER, VML_MEMBER]);

    let properties = member(&output, CONTROL_PROPERTIES_MEMBER);
    let vml = member(&output, VML_MEMBER);
    let source_properties =
        String::from_utf8(member(&source, CONTROL_PROPERTIES_MEMBER)).expect("properties XML");
    assert!(source_properties.contains("checked=\"Checked\""));
    let expected_properties = source_properties
        .replace("checked=\"Checked\"", "checked=\"Unchecked\"")
        .into_bytes();
    let expected_vml = String::from_utf8(member(&source, VML_MEMBER))
        .expect("VML XML")
        .replace("<x:Checked>1</x:Checked>", "<x:Checked>0</x:Checked>")
        .into_bytes();
    assert_eq!(properties, expected_properties.as_slice());
    assert_eq!(vml, expected_vml.as_slice());
    assert!(
        properties
            .windows(b"checked=\"Unchecked\"".len())
            .any(|window| window == b"checked=\"Unchecked\"")
    );
    assert!(
        vml.windows(b"<x:Checked>0</x:Checked>".len())
            .any(|window| window == b"<x:Checked>0</x:Checked>")
    );
    assert!(
        !properties
            .windows(b"checked=\"Checked\"".len())
            .any(|window| window == b"checked=\"Checked\"")
    );
    assert!(
        !vml.windows(b"<x:Checked>1</x:Checked>".len())
            .any(|window| window == b"<x:Checked>1</x:Checked>")
    );

    let reopened = Workbook::from_bytes(output).expect("reopen changed scalar");
    let control = reopened
        .sheet(0)
        .expect("sheet")
        .expect("worksheet")
        .form_control(ControlSelector::position(0))
        .expect("control readback")
        .expect("control");
    assert_eq!(
        control.properties().checked(),
        Some(&litchi_xlsx::form_control::KnownOrUnknown::Known(
            Checked::Unchecked
        ))
    );
}

#[test]
fn source_noop_is_exact_and_inverse_restores_lexical_pair() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let editor =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(source.clone())))
            .expect("open source scalar editor");
    let mut noop_edit = editor.edit("Sheet1").expect("select source worksheet");
    noop_edit
        .set_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Checked),
        )
        .expect("stage exact scalar no-op");
    let noop = noop_edit.commit().expect("commit exact no-op");
    assert!(!noop.changed());
    assert!(noop.patch().is_empty());
    let mut no_op_bytes = Vec::new();
    editor
        .publish_commit_to_stream(&mut no_op_bytes, &noop)
        .expect("publish exact no-op");
    assert_eq!(no_op_bytes, source);

    let (changed_bytes, changed) = source_commit(
        &source,
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    );
    let mut package = OpcPackage::from_bytes(&source).expect("source OPC package");
    changed
        .patch()
        .apply(&mut package)
        .expect("apply forward patch");
    assert_eq!(
        package
            .get_part(&litchi_opc::PackURI::new(format!("/{CONTROL_PROPERTIES_MEMBER}")).unwrap())
            .unwrap()
            .blob(),
        member(&changed_bytes, CONTROL_PROPERTIES_MEMBER).as_slice()
    );
    let inverse = changed.patch().inverse();
    inverse.apply(&mut package).expect("apply inverse patch");
    assert_eq!(
        package
            .get_part(&litchi_opc::PackURI::new(format!("/{CONTROL_PROPERTIES_MEMBER}")).unwrap())
            .unwrap()
            .blob(),
        member(&source, CONTROL_PROPERTIES_MEMBER).as_slice()
    );
    assert_eq!(
        package
            .get_part(&litchi_opc::PackURI::new(format!("/{VML_MEMBER}")).unwrap())
            .unwrap()
            .blob(),
        member(&source, VML_MEMBER).as_slice()
    );
}

#[test]
fn source_scalar_preserves_unrelated_mce_sidecars_and_opaque_formula_bytes() {
    let source = fixture_bytes("tdf134769.xlsx");
    let (output, commit) = source_commit(
        &source,
        ScalarField::NoThreeD,
        Some(ScalarValue::Boolean(false)),
    );
    assert!(commit.changed());
    assert_only_members_changed(&source, &output, &[CONTROL_PROPERTIES_MEMBER, VML_MEMBER]);
    assert!(
        member(&output, CONTROL_PROPERTIES_MEMBER)
            .windows(b"noThreeD=\"0\"".len())
            .any(|window| window == b"noThreeD=\"0\"")
    );
    assert!(
        member(&output, VML_MEMBER)
            .windows(b"<x:NoThreeD>False</x:NoThreeD>".len())
            .any(|window| window == b"<x:NoThreeD>False</x:NoThreeD>")
    );
    assert!(
        member(&source, CONTROL_PROPERTIES_MEMBER)
            .windows(b"fmlaLink=\"#REF!\"".len())
            .any(|window| window == b"fmlaLink=\"#REF!\"")
    );
    assert!(
        member(&output, CONTROL_PROPERTIES_MEMBER)
            .windows(b"fmlaLink=\"#REF!\"".len())
            .any(|window| window == b"fmlaLink=\"#REF!\"")
    );
    assert!(
        member(&source, VML_MEMBER)
            .windows(b"FmlaLink".len())
            .any(|window| window == b"FmlaLink")
    );
    assert!(
        member(&output, VML_MEMBER)
            .windows(b"FmlaLink".len())
            .any(|window| window == b"FmlaLink")
    );
    assert_eq!(
        member(&source, WORKSHEET_MEMBER),
        member(&output, WORKSHEET_MEMBER)
    );
    assert_eq!(
        member(&source, DRAWING_MEMBER),
        member(&output, DRAWING_MEMBER)
    );
    assert_eq!(
        member(&source, WORKSHEET_RELS_MEMBER),
        member(&output, WORKSHEET_RELS_MEMBER)
    );
}

#[test]
fn source_only_formula_allows_exact_noop_and_unrelated_scalar_but_refuses_replacement() {
    let source = fixture_bytes("tdf134769.xlsx");

    let editor =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(source.clone())))
            .expect("open source-only formula fixture");
    let snapshot = editor
        .snapshot("Sheet1")
        .expect("snapshot source-only formula fixture");
    let exact_formula = snapshot
        .control(ControlSelector::position(0))
        .expect("source-only control lookup")
        .expect("source-only control")
        .properties()
        .scalar(ScalarField::FmlaLink)
        .expect("source-only formula scalar");
    assert!(matches!(
        &exact_formula,
        ScalarValue::Formula(formula) if formula.as_str() == "#REF!"
    ));

    let mut exact_edit = editor.edit("Sheet1").expect("select exact no-op sheet");
    exact_edit
        .set_scalar(
            ControlSelector::position(0),
            ScalarField::FmlaLink,
            Some(exact_formula),
        )
        .expect("stage exact source-only formula no-op");
    let exact_commit = exact_edit.commit().expect("commit exact source-only no-op");
    assert!(!exact_commit.changed());
    assert!(exact_commit.patch().is_empty());
    let mut exact_output = Vec::new();
    editor
        .publish_commit_to_stream(&mut exact_output, &exact_commit)
        .expect("publish exact source-only formula no-op");
    assert_eq!(exact_output, source);

    let (unrelated_output, unrelated_commit) = source_commit(
        &source,
        ScalarField::NoThreeD,
        Some(ScalarValue::Boolean(false)),
    );
    assert!(unrelated_commit.changed());
    assert!(
        member(&unrelated_output, CONTROL_PROPERTIES_MEMBER)
            .windows(b"fmlaLink=\"#REF!\"".len())
            .any(|window| window == b"fmlaLink=\"#REF!\""),
        "unrelated scalar edit discarded source-only formula bytes"
    );
    assert!(
        member(&unrelated_output, VML_MEMBER)
            .windows(b"<x:FmlaLink>#REF!</x:FmlaLink>".len())
            .any(|window| window == b"<x:FmlaLink>#REF!</x:FmlaLink>"),
        "unrelated scalar edit discarded source-only VML formula bytes"
    );

    let replacement_editor =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(source.clone())))
            .expect("reopen source-only formula fixture for replacement");
    let mut replacement = replacement_editor
        .edit("Sheet1")
        .expect("select replacement sheet");
    let staged = replacement.set_scalar(
        ControlSelector::position(0),
        ScalarField::FmlaLink,
        Some(ScalarValue::Formula(
            FormControlFormula::new("Sheet1!A1").expect("valid authored formula"),
        )),
    );
    let mut refused_output = Vec::new();
    let error = match staged {
        Err(error) => error,
        Ok(_) => match replacement.commit() {
            Err(error) => error,
            Ok(commit) => replacement_editor
                .publish_commit_to_stream(&mut refused_output, &commit)
                .expect_err("source-only formula replacement was accepted"),
        },
    };
    assert!(
        matches!(
            error,
            Error::Unsupported { .. }
                | Error::Invalid(_)
                | Error::FormControl(litchi_xlsx::form_control::FormControlError::Invalid(_))
        ),
        "source-only formula replacement returned the wrong error: {error:?}"
    );
    assert!(
        refused_output.is_empty(),
        "source-only formula replacement published partial output"
    );
}

#[test]
fn source_scalar_ignores_foreign_namespace_client_data_before_real_owner() {
    let source = rewrite_member(&fixture_bytes("singlecontrol.xlsx"), VML_MEMBER, |xml| {
        let marker = "  <x:ClientData ObjectType=\"Checkbox\">";
        assert!(xml.contains(marker), "real x:ClientData marker missing");
        xml.replacen(
            marker,
            "  <q:ClientData xmlns:q=\"urn:opaque\"><q:Checked>1</q:Checked></q:ClientData>\n  <x:ClientData ObjectType=\"Checkbox\">",
            1,
        )
    });
    let (output, commit) = source_commit(
        &source,
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    );
    assert!(commit.changed());
    assert_only_members_changed(&source, &output, &[CONTROL_PROPERTIES_MEMBER, VML_MEMBER]);
    let output_vml = member(&output, VML_MEMBER);
    assert!(
        output_vml
            .windows(b"<q:ClientData xmlns:q=\"urn:opaque\"><q:Checked>1</q:Checked></q:ClientData>".len())
            .any(|window| {
                window
                    == b"<q:ClientData xmlns:q=\"urn:opaque\"><q:Checked>1</q:Checked></q:ClientData>"
            }),
        "foreign ClientData was not preserved"
    );
    assert!(
        output_vml
            .windows(b"<x:Checked>0</x:Checked>".len())
            .any(|window| window == b"<x:Checked>0</x:Checked>"),
        "real x:ClientData was not updated"
    );
    let reopened = Workbook::from_bytes(output).expect("reopen foreign ClientData source");
    assert_eq!(
        reopened
            .sheet("Sheet1")
            .expect("reopened sheet lookup")
            .expect("reopened worksheet")
            .form_control(ControlSelector::position(0))
            .expect("reopened control lookup")
            .expect("reopened control")
            .properties()
            .checked(),
        Some(&litchi_xlsx::form_control::KnownOrUnknown::Known(
            Checked::Unchecked
        ))
    );
}

#[test]
fn source_scalar_does_not_select_foreign_shape_with_duplicate_id() {
    let source = rewrite_member(&fixture_bytes("singlecontrol.xlsx"), VML_MEMBER, |xml| {
        let marker = "<v:shape id=\"_x0000_s1026\"";
        assert!(xml.contains(marker), "real v:shape marker missing");
        xml.replacen(
            marker,
            "<q:shape xmlns:q=\"urn:opaque\" id=\"_x0000_s1026\"><q:ClientData ObjectType=\"Checkbox\"><q:Checked>1</q:Checked></q:ClientData></q:shape>\n <v:shape id=\"_x0000_s1026\"",
            1,
        )
    });
    let editor =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(source.clone())))
            .expect("open duplicate-id namespace fixture");
    let mut edit = editor
        .edit("Sheet1")
        .expect("select duplicate-id worksheet");
    let staged = edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    );
    let mut output = Vec::new();
    let result = match staged {
        Err(error) => Err(error),
        Ok(_) => match edit.commit() {
            Err(error) => Err(error),
            Ok(commit) => editor
                .publish_commit_to_stream(&mut output, &commit)
                .map(|_| ()),
        },
    };
    match result {
        Ok(()) => {
            assert_only_members_changed(&source, &output, &[CONTROL_PROPERTIES_MEMBER, VML_MEMBER]);
            let output_vml = member(&output, VML_MEMBER);
            assert!(
                output_vml
                    .windows(b"<q:shape xmlns:q=\"urn:opaque\" id=\"_x0000_s1026\"><q:ClientData ObjectType=\"Checkbox\"><q:Checked>1</q:Checked></q:ClientData></q:shape>".len())
                    .any(|window| {
                        window
                            == b"<q:shape xmlns:q=\"urn:opaque\" id=\"_x0000_s1026\"><q:ClientData ObjectType=\"Checkbox\"><q:Checked>1</q:Checked></q:ClientData></q:shape>"
                    }),
                "foreign duplicate-id shape was not preserved"
            );
            assert!(
                output_vml
                    .windows(b"<x:Checked>0</x:Checked>".len())
                    .any(|window| window == b"<x:Checked>0</x:Checked>"),
                "real v:shape was not updated"
            );
        },
        Err(error) => {
            assert!(
                matches!(
                    error,
                    Error::Unsupported { .. } | Error::Invalid(_) | Error::FormControl(_)
                ),
                "foreign duplicate-id shape returned the wrong error: {error:?}"
            );
            assert!(
                output.is_empty(),
                "foreign duplicate-id refusal published output"
            );
        },
    }
}

#[test]
fn source_scalar_resolves_xml_escaped_vml_shape_id() {
    let source = rewrite_member(&fixture_bytes("singlecontrol.xlsx"), VML_MEMBER, |xml| {
        let marker = "id=\"_x0000_s1026\"";
        assert!(xml.contains(marker), "real shape id marker missing");
        xml.replacen(marker, "id=\"&#x5F;x0000_s1026\"", 1)
    });
    let (output, commit) = source_commit(
        &source,
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    );
    assert!(commit.changed());
    let output_vml = member(&output, VML_MEMBER);
    assert!(
        output_vml
            .windows(b"id=\"&#x5F;x0000_s1026\"".len())
            .any(|window| window == b"id=\"&#x5F;x0000_s1026\""),
        "escaped shape id lexical bytes were not preserved"
    );
    assert!(
        output_vml
            .windows(b"<x:Checked>0</x:Checked>".len())
            .any(|window| window == b"<x:Checked>0</x:Checked>"),
        "escaped shape id target was not updated"
    );
    let reopened = Workbook::from_bytes(output).expect("reopen escaped-shape-id source");
    assert_eq!(
        reopened
            .sheet("Sheet1")
            .expect("reopened escaped-id sheet lookup")
            .expect("reopened escaped-id worksheet")
            .form_control(ControlSelector::position(0))
            .expect("reopened escaped-id control lookup")
            .expect("reopened escaped-id control")
            .properties()
            .checked(),
        Some(&litchi_xlsx::form_control::KnownOrUnknown::Known(
            Checked::Unchecked
        ))
    );
}

#[test]
fn source_scalar_refuses_disagreement_without_publishing_one_sided_changes() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let disagreement = rewrite_member(&source, VML_MEMBER, |xml| {
        xml.replace("<x:Checked>1</x:Checked>", "<x:Checked>0</x:Checked>")
    });
    let editor = SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(
        disagreement.clone(),
    )))
    .expect("open disagreement source");
    let mut edit = editor.edit("Sheet1").expect("select source worksheet");
    edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::Checked,
        checked_scalar(Checked::Mixed),
    )
    .expect("stage disagreement candidate");
    assert!(
        edit.commit().is_err(),
        "disagreeing x14/VML pair was published"
    );

    let unknown = rewrite_member(&source, VML_MEMBER, |xml| {
        xml.replace("<x:Checked>1</x:Checked>", "<x:Checked>9</x:Checked>")
    });
    let editor =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(unknown)))
            .expect("open unknown-mirror source");
    let mut edit = editor
        .edit("Sheet1")
        .expect("select unknown-mirror worksheet");
    edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    )
    .expect("stage unknown-mirror candidate");
    assert!(edit.commit().is_err(), "unknown VML mirror was guessed");
}

#[test]
fn source_scalar_rejects_duplicate_name_selector() {
    let source = fixture_bytes("tdf161365.xlsx");
    let editor =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(source)))
            .expect("open duplicate-name source");
    let selected = editor.edit("Arkusz1");
    assert!(
        selected.is_ok(),
        "select duplicate-name sheet: {:?}",
        selected.as_ref().err()
    );
    let mut edit = selected.expect("select duplicate-name sheet");
    assert!(
        edit.set_scalar(
            ControlSelector::name("Check Box 4"),
            ScalarField::NoThreeD,
            Some(ScalarValue::Boolean(false)),
        )
        .is_err()
    );
}

#[test]
fn source_scalar_stale_version_and_cancellation_leave_output_empty() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let versioned = Arc::new(VersionedSource::new(source.clone()));
    let editor = SourceBackedFormControlEditor::from_read_at(versioned.clone())
        .expect("open versioned source");
    let mut edit = editor.edit("Sheet1").expect("select source worksheet");
    edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    )
    .expect("stage stale candidate");
    let commit = edit.commit().expect("prepare stale candidate");
    versioned.change();
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::PatchConflict { .. }) | Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty(), "stale refusal wrote output bytes");

    let canceled = SourceBackedFormControlEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(source)),
        ReadLimits::default(),
        canceled_context(),
    );
    assert!(canceled.is_err(), "canceled source editor was admitted");
}

#[test]
fn source_patch_rejects_changed_incoming_readset_and_signed_mutation() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let (_, commit) = source_commit(
        &source,
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    );

    let incoming_changed = rewrite_member(&source, WORKSHEET_RELS_MEMBER, |xml| {
        xml.replace(
            "</Relationships>",
            "<!-- incoming graph changed after candidate capture --></Relationships>",
        )
    });
    let mut package =
        OpcPackage::from_bytes(&incoming_changed).expect("incoming candidate package");
    let serialized_before =
        PackageWriter::to_bytes(&package).expect("serialize incoming package snapshot");
    assert!(matches!(
        commit.patch().apply(&mut package),
        Err(Error::PatchConflict { .. })
    ));
    let serialized_after =
        PackageWriter::to_bytes(&package).expect("serialize refused incoming package");
    assert_eq!(
        serialized_after, serialized_before,
        "incoming read-set refusal mutated the package"
    );

    let signed = signed_fixture(&source);
    let editor =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(signed)))
            .expect("open signed source");
    let mut edit = editor.edit("Sheet1").expect("select signed worksheet");
    edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    )
    .expect("stage signed mutation");
    let commit = edit.commit().expect("prepare signed mutation");
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Signed) | Err(Error::Package(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(output.is_empty(), "signed refusal wrote output bytes");
}

#[test]
fn source_scalar_limits_and_sink_refusals_do_not_publish_partial_pairs() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let limited =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(source.clone())))
            .expect("open source scalar editor")
            .with_limits(OwnerLimits::default().with_max_vml_bytes(1));
    assert!(limited.snapshot("Sheet1").is_err(), "VML limit was ignored");

    let editor =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(source.clone())))
            .expect("open source scalar editor");
    let mut edit = editor.edit("Sheet1").expect("select source worksheet");
    edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    )
    .expect("stage sink candidate");
    let commit = edit.commit().expect("prepare sink candidate");
    let mut sink = FailingSink {
        accepted: 0,
        limit: 0,
    };
    assert!(editor.publish_commit_to_stream(&mut sink, &commit).is_err());
    assert_eq!(
        sink.accepted, 0,
        "sink observed partial output after refusal"
    );
}

#[test]
fn source_scalar_multi_control_batches_refuse_mixed_targets_atomically() {
    let source = fixture_bytes("tdf120301_xmlSpaceParsing.xlsx");
    let editor =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(source.clone())))
            .expect("open multi-control source");
    let mut edit = editor.edit("Sheet1").expect("select multi-control sheet");
    edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::NoThreeD,
        Some(ScalarValue::Boolean(false)),
    )
    .expect("stage first scalar");
    edit.set_scalar(
        ControlSelector::position(1),
        ScalarField::NoThreeD,
        Some(ScalarValue::Boolean(false)),
    )
    .expect("stage second scalar");
    assert_eq!(edit.before().form_controls().len(), 2);
    assert!(edit.commit().is_err(), "mixed-control batch was published");
}

#[test]
fn ordinary_worksheet_edit_publishes_the_same_paired_scalar_closure() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let base = Workbook::from_bytes(source.clone()).expect("open ordinary workbook");
    let mut edit = base.edit().expect("begin ordinary workbook edit");
    {
        let mut sheet = edit
            .sheet("Sheet1")
            .expect("select ordinary worksheet")
            .expect("worksheet");
        sheet
            .set_form_control_scalar(
                ControlSelector::position(0),
                ScalarField::Checked,
                checked_scalar(Checked::Unchecked),
            )
            .expect("stage ordinary paired scalar");
        sheet
            .edit_form_control(ControlSelector::position(0))
            .expect("pin ordinary control editor")
            .set_scalar(ScalarField::NoThreeD, Some(ScalarValue::Boolean(false)))
            .expect("stage ordinary handle scalar");
    }
    let commit = edit.commit().expect("commit ordinary paired scalar");
    let output = commit
        .into_workbook()
        .to_plain_bytes()
        .expect("write ordinary workbook");
    assert_only_members_changed(&source, &output, &[CONTROL_PROPERTIES_MEMBER, VML_MEMBER]);
    assert!(
        member(&output, CONTROL_PROPERTIES_MEMBER)
            .windows(b"checked=\"Unchecked\"".len())
            .any(|window| window == b"checked=\"Unchecked\"")
    );
    assert!(
        member(&output, CONTROL_PROPERTIES_MEMBER)
            .windows(b"noThreeD=\"0\"".len())
            .any(|window| window == b"noThreeD=\"0\"")
    );
    assert!(
        member(&output, VML_MEMBER)
            .windows(b"<x:Checked>0</x:Checked>".len())
            .any(|window| window == b"<x:Checked>0</x:Checked>")
    );
    assert!(
        member(&output, VML_MEMBER)
            .windows(b"<x:NoThreeD>False</x:NoThreeD>".len())
            .any(|window| window == b"<x:NoThreeD>False</x:NoThreeD>")
    );
}

#[test]
fn source_scalar_composes_two_fields_on_one_control_without_losing_the_first_overlay() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let editor =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(source.clone())))
            .expect("open source scalar editor");
    let mut edit = editor.edit("Sheet1").expect("select source worksheet");
    edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    )
    .expect("stage checked scalar");
    edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::NoThreeD,
        Some(ScalarValue::Boolean(false)),
    )
    .expect("stage no-three-d scalar");
    let commit = edit.commit().expect("commit mixed scalar composition");
    let mut output = Vec::new();
    let published = editor
        .publish_commit_to_stream(&mut output, &commit)
        .expect("publish mixed scalar composition");
    let published_control = published
        .control(ControlSelector::position(0))
        .expect("returned scalar snapshot")
        .expect("returned scalar control");
    assert_eq!(
        published_control.properties().checked(),
        Some(&litchi_xlsx::form_control::KnownOrUnknown::Known(
            Checked::Unchecked
        ))
    );
    assert_eq!(published_control.properties().no_three_d(), Some(false));
    let properties = member(&output, CONTROL_PROPERTIES_MEMBER);
    let vml = member(&output, VML_MEMBER);
    assert!(
        properties
            .windows(b"checked=\"Unchecked\"".len())
            .any(|window| window == b"checked=\"Unchecked\"")
    );
    assert!(
        properties
            .windows(b"noThreeD=\"0\"".len())
            .any(|window| window == b"noThreeD=\"0\"")
    );
    assert!(
        vml.windows(b"<x:Checked>0</x:Checked>".len())
            .any(|window| window == b"<x:Checked>0</x:Checked>")
    );
    assert!(
        vml.windows(b"<x:NoThreeD>False</x:NoThreeD>".len())
            .any(|window| window == b"<x:NoThreeD>False</x:NoThreeD>")
    );

    let reopened = Workbook::from_bytes(output).expect("reopen mixed scalar composition");
    let reopened_control = reopened
        .sheet(0)
        .expect("reopened sheet lookup")
        .expect("reopened worksheet")
        .form_control(ControlSelector::position(0))
        .expect("reopened scalar control lookup")
        .expect("reopened scalar control");
    assert_eq!(
        reopened_control.properties().checked(),
        Some(&litchi_xlsx::form_control::KnownOrUnknown::Known(
            Checked::Unchecked
        ))
    );
    assert_eq!(reopened_control.properties().no_three_d(), Some(false));
}

#[test]
fn source_commit_rechecks_cancellation_after_snapshot() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let (context, cancellation) = live_context();
    let editor = SourceBackedFormControlEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(source)),
        ReadLimits::default(),
        context,
    )
    .expect("open managed source scalar editor");
    let mut edit = editor
        .edit("Sheet1")
        .expect("capture managed edit snapshot");
    assert!(edit.before().has_execution_budget());
    edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    )
    .expect("stage managed scalar before cancellation");
    cancellation.cancel();

    let error = edit
        .commit()
        .expect_err("commit ignored cancellation after the snapshot");
    assert!(
        matches!(
            error,
            Error::Package(OpcError::Cancelled)
                | Error::Package(OpcError::Execution(litchi_core::ExecutionError::Cancelled))
                | Error::FormControl(litchi_xlsx::form_control::FormControlError::Execution(
                    litchi_core::ExecutionError::Cancelled
                ))
        ),
        "cancellation was not retained by commit: {error:?}"
    );
}

#[test]
fn managed_scalar_result_handles_retain_and_refund_memory_and_objects() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let (probe_budget, _probe_cancellation, probe_context) =
        managed_scalar_context(u64::MAX, u64::MAX);
    let probe_editor = SourceBackedFormControlEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(source.clone())),
        ReadLimits::default(),
        probe_context,
    )
    .expect("open probe managed scalar editor");
    let probe_warmup = probe_editor
        .snapshot("Sheet1")
        .expect("warm probe source cache");
    drop(probe_warmup);
    let probe_baseline = scalar_resource_usage(&probe_budget);
    let mut probe_edit = probe_editor.edit("Sheet1").expect("select probe worksheet");
    probe_edit
        .set_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Unchecked),
        )
        .expect("stage probe scalar");
    let before_probe_commit = scalar_resource_usage(&probe_budget);
    let probe_commit = probe_edit.commit().expect("commit probe scalar");
    let after_probe_commit = scalar_resource_usage(&probe_budget);
    for (resource, delta) in [
        (
            Resource::Memory,
            after_probe_commit[0].saturating_sub(before_probe_commit[0]),
        ),
        (
            Resource::Objects,
            after_probe_commit[1].saturating_sub(before_probe_commit[1]),
        ),
    ] {
        assert!(
            delta > 0,
            "managed scalar commit did not retain {resource:?} output charge"
        );
    }

    let mut probe_second_edit = probe_editor
        .edit("Sheet1")
        .expect("select second probe worksheet");
    probe_second_edit
        .set_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Unchecked),
        )
        .expect("stage second probe scalar");
    let probe_limit = scalar_resource_usage(&probe_budget);
    drop(probe_second_edit);
    drop(probe_commit);
    drop(probe_editor);
    assert_eq!(scalar_resource_usage(&probe_budget), [0, 0]);

    let (budget, _cancellation, context) = managed_scalar_context(probe_limit[0], probe_limit[1]);
    let editor = SourceBackedFormControlEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(source)),
        ReadLimits::default(),
        context,
    )
    .expect("open finite managed scalar editor");
    let warmup = editor.snapshot("Sheet1").expect("warm finite source cache");
    drop(warmup);
    let baseline = scalar_resource_usage(&budget);
    assert_eq!(
        baseline, probe_baseline,
        "probe and finite managed warm-cache baselines diverged"
    );
    assert!(
        baseline[0] <= probe_limit[0] && baseline[1] <= probe_limit[1],
        "finite managed baseline exceeded measured probe limit: {baseline:?} > {probe_limit:?}"
    );

    let mut edit = editor.edit("Sheet1").expect("select finite worksheet");
    edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    )
    .expect("stage finite scalar");
    let commit = edit.commit().expect("commit finite scalar");
    let retained_patch = commit.patch().clone();
    let retained_snapshot = commit.snapshot().clone();
    let retained_collection = retained_snapshot.form_controls().clone();
    // This fixture has no public OpaqueXml extension child; the clone below
    // covers the source-backed lexical payload lifetime exposed by Properties.
    let retained_properties = retained_snapshot
        .control(ControlSelector::position(0))
        .expect("retained control lookup")
        .expect("retained control")
        // This public clone retains the source-backed lexical properties.
        .properties()
        .clone();
    drop(commit);
    let retained_usage = scalar_resource_usage(&budget);
    assert!(
        retained_usage[0] > baseline[0] && retained_usage[1] > baseline[1],
        "cloned scalar results did not retain both managed charges: baseline {baseline:?}, retained {retained_usage:?}"
    );

    let blocked = (|| {
        let mut blocked_edit = editor.edit("Sheet1")?;
        blocked_edit.set_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Unchecked),
        )?;
        blocked_edit.commit()
    })();
    match blocked {
        Err(Error::ResourceLimit(limit))
        | Err(Error::Package(OpcError::Execution(litchi_core::ExecutionError::ResourceLimit(
            limit,
        ))))
        | Err(Error::FormControl(litchi_xlsx::form_control::FormControlError::Execution(
            litchi_core::ExecutionError::ResourceLimit(limit),
        ))) => assert!(
            matches!(limit.resource, Resource::Memory | Resource::Objects),
            "retained scalar result exhausted the wrong resource: {limit:?}"
        ),
        Err(error) => panic!("retained scalar result returned the wrong error: {error:?}"),
        Ok(_) => panic!("retained scalar result did not exhaust the finite budget"),
    }

    drop(retained_patch);
    drop(retained_snapshot);
    let collection_usage = scalar_resource_usage(&budget);
    assert!(
        collection_usage[0] > baseline[0] || collection_usage[1] > baseline[1],
        "cloned FormControlCollection did not retain managed charges after its snapshot was dropped"
    );
    drop(retained_collection);
    let properties_usage = scalar_resource_usage(&budget);
    assert!(
        collection_usage[0] > properties_usage[0] && collection_usage[1] > properties_usage[1],
        "cloned FormControlCollection did not add both managed charges beyond its source-bearing Properties clone: collection {collection_usage:?}, properties {properties_usage:?}"
    );
    assert!(
        properties_usage[0] > baseline[0] && properties_usage[1] > baseline[1],
        "standalone cloned source-bearing Properties did not retain managed charges"
    );
    drop(retained_properties);
    assert_eq!(
        scalar_resource_usage(&budget),
        baseline,
        "dropping all retained scalar result handles did not refund charges"
    );

    let mut retry = editor.edit("Sheet1").expect("select retry worksheet");
    retry
        .set_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Unchecked),
        )
        .expect("stage retry scalar");
    let retry_commit = retry.commit().expect("refunded budget admitted retry");
    drop(retry_commit);
    assert_eq!(
        scalar_resource_usage(&budget),
        baseline,
        "retry result did not release after commit drop"
    );
    drop(editor);
    assert_eq!(scalar_resource_usage(&budget), [0, 0]);
}

#[test]
fn source_commit_rechecks_exact_custom_vml_limit_after_snapshot() {
    let source = rewrite_member(
        &fixture_bytes("tdf134769.xlsx"),
        CONTROL_PROPERTIES_MEMBER,
        |xml| xml.replace("#REF!", "A1"),
    );
    let source = rewrite_member(&source, VML_MEMBER, |xml| xml.replace("#REF!", "A1"));
    let exact_vml_limit = member(&source, VML_MEMBER).len();
    let editor =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(source)))
            .expect("open source scalar editor")
            .with_limits(OwnerLimits::default().with_max_vml_bytes(exact_vml_limit));
    let snapshot = editor
        .snapshot("Sheet1")
        .expect("exact source VML bound admits snapshot");
    assert_eq!(
        snapshot
            .read_set()
            .expect("source read set")
            .vml()
            .expect("retained VML")
            .bytes()
            .len(),
        exact_vml_limit
    );

    let mut edit = editor.edit("Sheet1").expect("select source worksheet");
    edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::FmlaLink,
        Some(ScalarValue::Formula(
            FormControlFormula::new("Sheet1!A1").expect("authored formula"),
        )),
    )
    .expect("stage candidate within source lexical policy");
    let error = edit
        .commit()
        .expect_err("commit ignored the exact custom VML output limit");
    assert!(
        matches!(
            error,
            Error::ResourceLimit(_)
                | Error::FormControl(litchi_xlsx::form_control::FormControlError::Limit { .. })
                | Error::Package(OpcError::ReadLimit { .. })
        ),
        "custom VML limit returned the wrong error: {error:?}"
    );
}

#[test]
fn source_publication_reopens_through_the_same_source_owner() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let (output, commit) = source_commit(
        &source,
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    );
    let reopened =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(output.clone())))
            .expect("reopen published source editor");
    let snapshot = reopened
        .snapshot("Sheet1")
        .expect("reopen published source snapshot");
    let control = snapshot
        .control(ControlSelector::position(0))
        .expect("reopen published control lookup")
        .expect("reopen published control");
    assert_eq!(
        control.properties().checked(),
        Some(&litchi_xlsx::form_control::KnownOrUnknown::Known(
            Checked::Unchecked
        ))
    );
    let read_set = snapshot.read_set().expect("published read set");
    assert_eq!(read_set.properties().len(), 1);
    assert_eq!(
        read_set.properties()[0].bytes(),
        member(&output, CONTROL_PROPERTIES_MEMBER)
    );
    assert_eq!(
        read_set.vml().expect("published VML read").bytes(),
        member(&output, VML_MEMBER)
    );
    assert_eq!(snapshot.profile(), commit.snapshot().profile());
    assert_eq!(
        snapshot.source_version(),
        commit.snapshot().source_version()
    );
}

#[test]
fn paired_patch_rejects_each_stale_member_without_mutating_the_package() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let (_, commit) = source_commit(
        &source,
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    );

    for stale_member in [CONTROL_PROPERTIES_MEMBER, VML_MEMBER] {
        let stale = rewrite_member(&source, stale_member, |xml| format!("{xml} "));
        let mut package = OpcPackage::from_bytes(&stale).expect("open stale paired package");
        let before = PackageWriter::to_bytes(&package).expect("serialize stale package");
        assert!(
            matches!(
                commit.patch().apply(&mut package),
                Err(Error::PatchConflict { .. })
            ),
            "stale {stale_member} member was accepted"
        );
        assert_eq!(
            PackageWriter::to_bytes(&package).expect("serialize refused stale package"),
            before,
            "stale {stale_member} refusal mutated the package"
        );
    }
}

#[test]
fn empty_patch_still_checks_the_complete_source_read_set() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let editor =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(source.clone())))
            .expect("open source scalar editor");
    let mut edit = editor.edit("Sheet1").expect("select source worksheet");
    edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::Checked,
        checked_scalar(Checked::Checked),
    )
    .expect("stage exact no-op");
    let commit = edit.commit().expect("freeze exact no-op");
    assert!(commit.patch().is_empty());

    let changed_read_set = rewrite_member(&source, WORKSHEET_RELS_MEMBER, |xml| {
        xml.replace(
            "</Relationships>",
            "<!-- changed after no-op capture --></Relationships>",
        )
    });
    let mut package = OpcPackage::from_bytes(&changed_read_set).expect("open changed read set");
    let before = PackageWriter::to_bytes(&package).expect("serialize changed read set");
    assert!(
        matches!(
            commit.patch().apply(&mut package),
            Err(Error::PatchConflict { .. })
        ),
        "empty patch ignored a changed source read set"
    );
    assert_eq!(
        PackageWriter::to_bytes(&package).expect("serialize refused no-op package"),
        before,
        "empty patch refusal mutated the package"
    );
}

#[test]
fn ordinary_form_control_scalar_noops_have_empty_patches() {
    let base = Workbook::from_bytes(fixture_bytes("singlecontrol.xlsx"))
        .expect("open ordinary form-control workbook");
    let base_bytes = base.to_plain_bytes().expect("serialize ordinary base");

    let mut same = base.edit().expect("begin same-value form-control edit");
    same.sheet("Sheet1")
        .expect("same-value sheet lookup")
        .expect("same-value worksheet")
        .set_form_control_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Checked),
        )
        .expect("stage authored same-value scalar");
    let same_commit = same.commit().expect("commit authored same-value scalar");
    assert!(same_commit.patch().is_empty());
    assert_eq!(same_commit.patch().len(), 0);
    assert_eq!(
        same_commit
            .into_workbook()
            .to_plain_bytes()
            .expect("serialize same-value result"),
        base_bytes
    );

    let mut restored = base.edit().expect("begin restored form-control edit");
    restored
        .sheet("Sheet1")
        .expect("restored sheet lookup")
        .expect("restored worksheet")
        .set_form_control_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Unchecked),
        )
        .expect("stage changed scalar");
    restored
        .sheet("Sheet1")
        .expect("restored sheet lookup after change")
        .expect("restored worksheet after change")
        .set_form_control_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Checked),
        )
        .expect("restore original scalar");
    let restored_commit = restored.commit().expect("commit restored scalar");
    assert!(restored_commit.patch().is_empty());
    assert_eq!(restored_commit.patch().len(), 0);
    assert_eq!(
        restored_commit
            .into_workbook()
            .to_plain_bytes()
            .expect("serialize restored result"),
        base_bytes
    );
}

#[test]
fn inverse_patch_reopens_to_the_original_scalar_state() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let (forward, commit) = source_commit(
        &source,
        ScalarField::Checked,
        checked_scalar(Checked::Unchecked),
    );
    let forward_control = Workbook::from_bytes(forward)
        .expect("reopen forward pair")
        .sheet(0)
        .expect("forward sheet")
        .expect("forward worksheet")
        .form_control(ControlSelector::position(0))
        .expect("forward control lookup")
        .expect("forward control");
    assert_eq!(
        forward_control.properties().checked(),
        Some(&litchi_xlsx::form_control::KnownOrUnknown::Known(
            Checked::Unchecked
        ))
    );

    let mut package = OpcPackage::from_bytes(&source).expect("open source OPC package");
    commit
        .patch()
        .apply(&mut package)
        .expect("apply forward pair");
    commit
        .patch()
        .inverse()
        .apply(&mut package)
        .expect("apply inverse pair");
    let restored_properties = package
        .get_part(&litchi_opc::PackURI::new(format!("/{CONTROL_PROPERTIES_MEMBER}")).unwrap())
        .expect("restored properties part")
        .blob();
    let restored_vml = package
        .get_part(&litchi_opc::PackURI::new(format!("/{VML_MEMBER}")).unwrap())
        .expect("restored VML part")
        .blob();
    assert_eq!(
        restored_properties,
        member(&source, CONTROL_PROPERTIES_MEMBER).as_slice()
    );
    assert_eq!(restored_vml, member(&source, VML_MEMBER).as_slice());
    let restored_control =
        litchi_xlsx::form_control::parse(restored_properties).expect("parse inverse properties");
    assert_eq!(
        restored_control.checked(),
        Some(&litchi_xlsx::form_control::KnownOrUnknown::Known(
            Checked::Checked
        ))
    );
}

#[test]
fn signed_exact_noop_is_byte_preserving_but_authored_clear_is_refused() {
    let source = fixture_bytes("singlecontrol.xlsx");
    let signed = signed_fixture(&source);
    let editor =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(signed.clone())))
            .expect("open signed source");
    let mut noop_edit = editor.edit("Sheet1").expect("select signed worksheet");
    noop_edit
        .set_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Checked),
        )
        .expect("stage signed exact no-op");
    let noop = noop_edit.commit().expect("freeze signed exact no-op");
    assert!(!noop.changed());
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &noop)
        .expect("publish signed exact no-op");
    assert_eq!(output, signed);

    let editor =
        SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(source)))
            .expect("reopen unsigned source");
    let mut clear_edit = editor.edit("Sheet1").expect("select unsigned worksheet");
    clear_edit
        .set_scalar(ControlSelector::position(0), ScalarField::Checked, None)
        .expect("stage authored clear");
    let error = clear_edit
        .commit()
        .expect_err("authored scalar clear was silently accepted");
    assert!(
        matches!(error, Error::Unsupported { .. } | Error::FormControl(_)),
        "authored scalar clear returned the wrong error: {error:?}"
    );
}

#[test]
fn scalar_resource_limits_fail_before_any_pair_is_published() {
    let source = fixture_bytes("singlecontrol.xlsx");
    for limits in [
        OwnerLimits::default().with_max_scalar_operations(0),
        OwnerLimits::default().with_max_changed_parts(1),
        OwnerLimits::default().with_max_output_bytes(1),
        OwnerLimits::default().with_max_staging_bytes(1),
    ] {
        let editor = SourceBackedFormControlEditor::from_read_at(Arc::new(VersionedSource::new(
            source.clone(),
        )))
        .expect("open limited source editor")
        .with_limits(limits);
        let mut edit = editor.edit("Sheet1").expect("select limited worksheet");
        let staged = edit.set_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Unchecked),
        );
        if let Err(error) = staged {
            assert!(
                matches!(error, Error::ResourceLimit(_)),
                "scalar-operation limit returned the wrong error: {error:?}"
            );
            continue;
        }
        let error = edit
            .commit()
            .expect_err("resource-limited scalar pair was committed");
        assert!(
            matches!(error, Error::ResourceLimit(_)),
            "scalar resource limit returned the wrong error: {error:?}"
        );
    }
}

#[test]
fn ordinary_form_control_join_rejects_name_and_position_aliases_without_losing_either_edit() {
    let base = Workbook::from_bytes(fixture_bytes("singlecontrol.xlsx"))
        .expect("open ordinary form-control workbook");
    let mut by_position = base.edit().expect("begin position edit");
    by_position
        .sheet("Sheet1")
        .expect("position sheet lookup")
        .expect("position worksheet")
        .set_form_control_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Unchecked),
        )
        .expect("stage position scalar");
    let mut by_name = base.edit().expect("begin name edit");
    by_name
        .sheet("Sheet1")
        .expect("name sheet lookup")
        .expect("name worksheet")
        .set_form_control_scalar(
            ControlSelector::name("Check Box 2"),
            ScalarField::NoThreeD,
            Some(ScalarValue::Boolean(false)),
        )
        .expect("stage name scalar");

    let error = by_position
        .join(by_name)
        .expect_err("two aliases of one paired form-control owner joined");
    match error.failure() {
        JoinFailure::Overlap(conflicts) => {
            assert_eq!(conflicts.len(), 1);
            assert!(matches!(
                conflicts.conflicts(),
                [Conflict::FormControls { sheet, position }]
                    if sheet.as_ref() == "Sheet1" && *position == 0
            ));
        },
        other => panic!("wrong join failure for aliased form control: {other:?}"),
    }
    assert!(
        !by_position.is_empty(),
        "accepted branch lost its scalar edit"
    );
    assert!(
        !error.rejected().is_empty(),
        "rejected branch lost its scalar edit"
    );

    let left_commit = by_position.commit().expect("commit accepted branch");
    let right_commit = error
        .into_rejected()
        .commit()
        .expect("commit recovered rejected branch");
    for commit in [&left_commit, &right_commit] {
        assert!(!commit.patch().is_empty());
        assert_eq!(
            commit.patch().changes().len(),
            0,
            "form-control-only edit unexpectedly fabricated a semantic change"
        );
        assert!(matches!(
            commit.patch().package_changes(),
            [PackageChange::FormControl {
                sheet,
                control: 0,
                ..
            }] if sheet.as_ref() == "Sheet1"
        ));
        assert_eq!(
            commit.patch().len(),
            1,
            "form-control package change was omitted from Patch::len"
        );
    }
}

#[test]
fn ordinary_form_control_three_way_plan_conflicts_across_selector_aliases_and_resolves() {
    let base = Workbook::from_bytes(fixture_bytes("singlecontrol.xlsx"))
        .expect("open ordinary form-control workbook");
    let mut left = base.edit().expect("begin left form-control edit");
    left.sheet("Sheet1")
        .expect("left sheet lookup")
        .expect("left worksheet")
        .set_form_control_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Unchecked),
        )
        .expect("stage left scalar by position");
    let mut right = base.edit().expect("begin right form-control edit");
    right
        .sheet("Sheet1")
        .expect("right sheet lookup")
        .expect("right worksheet")
        .set_form_control_scalar(
            ControlSelector::name("Check Box 2"),
            ScalarField::NoThreeD,
            Some(ScalarValue::Boolean(false)),
        )
        .expect("stage right scalar by name");

    let mut plan = left
        .plan_three_way(right, MergeLimits::new(2, 64, 128, 64))
        .expect("plan aliased form-control branches");
    assert_eq!(plan.automatic_len(), 0);
    assert_eq!(plan.resolution(), None);
    assert!(matches!(
        plan.conflicts().conflicts(),
        [SubEditConflict::Effect(effect)]
            if effect.effect() == "sheet/0/form-controls"
                && effect.left_id() == "left"
                && effect.right_id() == "right"
                && effect.left_access() == EffectAccess::Write
                && effect.right_access() == EffectAccess::Write
    ));

    plan.resolve(MergeChoice::Left);
    assert_eq!(plan.resolution(), Some(MergeChoice::Left));
    let merged = plan
        .finish()
        .expect("resolve aliased form-control plan")
        .commit()
        .expect("commit resolved form-control plan")
        .into_workbook();
    assert_eq!(
        merged
            .sheet("Sheet1")
            .expect("merged sheet lookup")
            .expect("merged worksheet")
            .form_control(ControlSelector::position(0))
            .expect("merged control lookup")
            .expect("merged control")
            .properties()
            .checked(),
        Some(&litchi_xlsx::form_control::KnownOrUnknown::Known(
            Checked::Unchecked
        ))
    );
}

#[test]
fn ordinary_form_control_and_cell_edits_join_on_the_same_sheet() {
    let base = Workbook::from_bytes(fixture_bytes("singlecontrol.xlsx"))
        .expect("open ordinary form-control workbook");
    let mut control_edit = base.edit().expect("begin form-control edit");
    control_edit
        .sheet("Sheet1")
        .expect("control sheet lookup")
        .expect("control worksheet")
        .set_form_control_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Unchecked),
        )
        .expect("stage control scalar");
    let mut cell_edit = base.edit().expect("begin cell edit");
    cell_edit
        .sheet("Sheet1")
        .expect("cell sheet lookup")
        .expect("cell worksheet")
        .set("A1", "joined cell")
        .expect("stage cell edit");

    control_edit
        .join(cell_edit)
        .expect("disjoint control and cell effects were rejected");
    let commit = control_edit
        .commit()
        .expect("commit joined control and cell edits");
    assert_eq!(commit.patch().changes().len(), 1);
    assert!(matches!(commit.patch().changes(), [Change::Cell { .. }]));
    assert!(!commit.patch().is_empty());
    let output = commit
        .into_workbook()
        .to_plain_bytes()
        .expect("write joined workbook");
    let reopened = Workbook::from_bytes(output).expect("reopen joined workbook");
    assert!(matches!(
        reopened
            .sheet("Sheet1")
            .expect("reopened sheet lookup")
            .expect("reopened worksheet")
            .cell("A1")
            .expect("reopened cell")
            .stored()
            .expect("stored joined cell"),
        litchi_xlsx::Cell::Value(litchi_xlsx::Value::Text(value))
            if value.as_str() == "joined cell"
    ));
    assert_eq!(
        reopened
            .sheet("Sheet1")
            .expect("reopened control sheet lookup")
            .expect("reopened control worksheet")
            .form_control(ControlSelector::position(0))
            .expect("reopened control lookup")
            .expect("reopened control")
            .properties()
            .checked(),
        Some(&litchi_xlsx::form_control::KnownOrUnknown::Known(
            Checked::Unchecked
        ))
    );
}

#[test]
fn ordinary_form_control_patch_replays_forward_and_inverse_with_lineage_guards() {
    let base = Workbook::from_bytes(fixture_bytes("singlecontrol.xlsx"))
        .expect("open ordinary form-control workbook");
    let mut edit = base.edit().expect("begin ordinary form-control edit");
    edit.sheet("Sheet1")
        .expect("ordinary sheet lookup")
        .expect("ordinary worksheet")
        .set_form_control_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Unchecked),
        )
        .expect("stage ordinary form-control scalar");
    let commit = edit.commit().expect("commit ordinary form-control scalar");
    assert_eq!(commit.patch().package_changes().len(), 1);
    assert!(!commit.patch().is_empty());

    let applied = base
        .apply(commit.patch())
        .expect("replay ordinary forward patch")
        .into_workbook();
    assert_eq!(
        applied
            .sheet("Sheet1")
            .expect("applied sheet lookup")
            .expect("applied worksheet")
            .form_control(ControlSelector::position(0))
            .expect("applied control lookup")
            .expect("applied control")
            .properties()
            .checked(),
        Some(&litchi_xlsx::form_control::KnownOrUnknown::Known(
            Checked::Unchecked
        ))
    );

    let restored = commit
        .workbook()
        .apply(&commit.patch().inverse())
        .expect("replay ordinary inverse patch from committed target")
        .into_workbook();
    assert_eq!(
        restored
            .to_plain_bytes()
            .expect("serialize inverse workbook"),
        base.to_plain_bytes().expect("serialize base workbook")
    );

    let reopened_target = Workbook::from_bytes(
        commit
            .workbook()
            .to_plain_bytes()
            .expect("serialize committed target"),
    )
    .expect("reopen committed target independently");
    assert!(
        matches!(
            reopened_target.apply(&commit.patch().inverse()),
            Err(Error::PatchConflict { part }) if part == "workbook lineage"
        ),
        "inverse ordinary patch crossed an independently reopened lineage"
    );
}

#[test]
fn ordinary_form_control_durable_patch_replays_forward_and_inverse() {
    let base = Workbook::from_bytes(fixture_bytes("singlecontrol.xlsx"))
        .expect("open ordinary form-control workbook");
    let mut edit = base.edit().expect("begin ordinary form-control edit");
    edit.sheet("Sheet1")
        .expect("durable sheet lookup")
        .expect("durable worksheet")
        .set_form_control_scalar(
            ControlSelector::position(0),
            ScalarField::Checked,
            checked_scalar(Checked::Unchecked),
        )
        .expect("stage durable form-control scalar");
    let commit = edit.commit().expect("commit durable form-control scalar");
    let durable = commit.patch().durable().expect("encode durable patch");
    let applied = durable.apply(&base).expect("apply durable forward patch");
    let restored = durable
        .inverse()
        .apply(&applied)
        .expect("apply durable inverse patch");
    assert_eq!(
        restored
            .to_plain_bytes()
            .expect("serialize durable inverse"),
        base.to_plain_bytes().expect("serialize durable base")
    );
}
