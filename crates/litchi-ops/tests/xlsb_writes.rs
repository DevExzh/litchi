//! Unified XLSB worksheet-name writes through the public [`File`] API.
#![cfg(feature = "xlsb")]

use std::io::Cursor;

use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{
    BatchItem, BatchMode, Code, EditOptions, File, ItemResult, OpenOptions, Overwrite, Props,
    SaveOptions, SaveTo, SourceMode, props,
};

const FIRST: &str = "/sheet[1]";
const SECOND: &str = "/sheet[2]";

fn fixture() -> Vec<u8> {
    use litchi_xlsb::builder::{self, Builder};

    let mut writer = Builder::new();
    let mut data = builder::Worksheet::new("Data");
    data.set_cell(0, 0, "source");
    writer.add_worksheet(data);
    let mut archive = builder::Worksheet::new("Archive");
    archive.set_cell(0, 0, "untouched");
    writer.add_worksheet(archive);

    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build XLSB fixture")
        .write_to(&mut output)
        .expect("serialize XLSB fixture");
    output.into_inner()
}

fn signed_fixture() -> Vec<u8> {
    use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part as _, TargetMode};

    let source = OpcPackage::from_bytes(&fixture(), litchi_opc::SourceLimits::default())
        .expect("parse XLSB fixture as OPC");
    let mut package = OpcPackage::new();
    for relationship in source.rels().iter() {
        package
            .rels_mut()
            .try_add_relationship(
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.r_id().to_owned(),
                relationship.target_mode(),
            )
            .expect("copy package relationship");
    }
    let names: Vec<_> = source
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    for name in names {
        let part = source.get_part(&name).expect("find OPC part");
        let mut copy = BlobPart::new(
            name.clone(),
            part.content_type().to_owned(),
            part.blob().to_vec(),
        );
        for relationship in part.rels().iter() {
            copy.rels_mut()
                .try_add_relationship(
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.r_id().to_owned(),
                    relationship.target_mode(),
                )
                .expect("copy part relationship");
        }
        package.try_add_part(Box::new(copy)).expect("copy OPC part");
    }

    let origin = PackURI::new("/_xmlsignatures/origin.sigs").expect("signature origin URI");
    let signature = PackURI::new("/_xmlsignatures/sig1.xml").expect("signature URI");
    let mut origin_part = BlobPart::new(
        origin.clone(),
        litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        Vec::new(),
    );
    origin_part
        .rels_mut()
        .try_add_relationship(
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
                .to_owned(),
            signature.relative_ref(origin.base_uri()),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature relationship");
    package
        .try_add_part(Box::new(origin_part))
        .expect("signature origin part");
    package
        .try_add_part(Box::new(BlobPart::new(
            signature,
            litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
            b"<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>".to_vec(),
        )))
        .expect("signature XML part");
    package
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignatureOrigin".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature origin relationship");
    PackageWriter::to_bytes(&package).expect("serialize signed XLSB fixture")
}

fn has_signature_graph(bytes: &[u8]) -> bool {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open OPC archive");
    (0..archive.len()).any(|index| {
        archive
            .by_index(index)
            .expect("read OPC archive member")
            .name()
            .expect("valid ZIP member name")
            .to_ascii_lowercase()
            .starts_with("_xmlsignatures/")
    })
}

fn open_owner(bytes: &[u8], eager: bool) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .expect("open XLSB owner")
}

fn open_without_source(bytes: &[u8], eager: bool) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new()
            .source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            })
            .retain_source(false),
    )
    .expect("open unretained XLSB owner")
}

fn sheet_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn set_batch(path: &'static str, name: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(Props::new().with("name", name));
    item
}

fn scoped_sheet_query(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .expect("parse sheet selector")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/").expect("root selector scope"),
        )))
}

#[test]
fn sheet_name_set_query_and_publication_round_trip_for_source_and_eager() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original XLSB");
        assert_eq!(
            sheet_property(&original, FIRST, "name").as_deref(),
            Some("Data")
        );

        let renamed = original
            .set(FIRST, props! { "name" => "Input" })
            .expect("rename first worksheet");
        assert_eq!(
            sheet_property(&renamed, FIRST, "name").as_deref(),
            Some("Input")
        );
        assert_eq!(
            sheet_property(&original, FIRST, "name").as_deref(),
            Some("Data")
        );

        let matches = renamed
            .query("sheet[name=Input]")
            .expect("query renamed worksheet")
            .collect::<litchi_ops::Result<Vec<_>>>()
            .expect("collect renamed worksheet query");
        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].path().to_string(), FIRST);

        let direct = renamed.to_bytes().expect("publish renamed XLSB");
        let mut streamed = Vec::new();
        let count = renamed
            .write_to(&mut streamed)
            .expect("stream renamed XLSB");
        assert_eq!(count, streamed.len() as u64);
        assert_eq!(streamed, direct);

        let reopened = open_owner(&direct, eager);
        assert_eq!(
            sheet_property(&reopened, FIRST, "name").as_deref(),
            Some("Input")
        );
        assert_eq!(
            sheet_property(&reopened, SECOND, "name").as_deref(),
            Some("Archive")
        );

        let directory = tempfile::tempdir().expect("save directory");
        let origin = directory.path().join("origin.xlsb");
        std::fs::write(&origin, &bytes).expect("write XLSB origin");
        let path_owner = litchi_ops::open_with(
            &origin,
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .expect("open XLSB origin");
        path_owner
            .set(FIRST, props! { "name" => "Input" })
            .expect("rename path-backed worksheet")
            .save()
            .expect("save path-backed XLSB");
        let saved_origin = litchi_ops::open(&origin).expect("reopen saved XLSB origin");
        assert_eq!(
            sheet_property(&saved_origin, FIRST, "name").as_deref(),
            Some("Input")
        );

        let destination = directory.path().join("renamed.xlsb");
        renamed.save_as(&destination).expect("save renamed XLSB");
        assert_eq!(std::fs::read(destination).expect("read saved XLSB"), direct);

        assert_eq!(
            original.to_bytes().expect("original remains unchanged"),
            before
        );
    }
}

#[test]
fn signed_xlsb_routes_preserve_exact_bytes_until_explicit_unsign() {
    let bytes = signed_fixture();
    assert!(has_signature_graph(&bytes));

    for eager in [false, true] {
        let signed = open_owner(&bytes, eager);
        assert!(signed.is_signed());
        assert_eq!(signed.to_bytes().expect("publish signed XLSB"), bytes);

        let mut streamed = Vec::new();
        let count = signed.write_to(&mut streamed).expect("stream signed XLSB");
        assert_eq!(count, streamed.len() as u64);
        assert_eq!(streamed, bytes);

        let unchanged = signed
            .set(FIRST, props! { "name" => "Data" })
            .expect("signed exact no-op preserves the signature");
        assert!(unchanged.is_signed());
        assert_eq!(unchanged.to_bytes().expect("publish signed no-op"), bytes);
        assert_eq!(signed.to_bytes().unwrap(), bytes);

        let File::Signed(wrapper) = signed else {
            panic!("the synthetic OPC signature must open as File::Signed");
        };
        let unsigned = wrapper.unsign();
        assert!(!unsigned.is_signed());
        assert_eq!(
            sheet_property(&unsigned, FIRST, "name").as_deref(),
            Some("Data")
        );

        let no_edit = unsigned
            .to_bytes()
            .expect("publish explicitly unsigned XLSB");
        assert!(!has_signature_graph(&no_edit));
        let mut no_edit_stream = Vec::new();
        unsigned
            .write_to(&mut no_edit_stream)
            .expect("stream explicitly unsigned XLSB");
        assert_eq!(no_edit_stream, no_edit);

        let changed = unsigned
            .set(FIRST, props! { "name" => "Input" })
            .expect("edit explicitly unsigned XLSB");
        let changed_bytes = changed.to_bytes().expect("publish changed XLSB");
        assert!(!has_signature_graph(&changed_bytes));
        let reopened = open_owner(&changed_bytes, eager);
        assert!(!reopened.is_signed());
        assert_eq!(
            sheet_property(&reopened, FIRST, "name").as_deref(),
            Some("Input")
        );
    }
}

#[test]
fn signed_coalesced_cell_cancellation_keeps_a_positive_snapshot_peak() {
    let bytes = signed_fixture();

    for eager in [false, true] {
        let signed = open_owner(&bytes, eager);
        let canonical = signed
            .edit(|edit| edit.set("/sheet[1]/A1", props! { "value" => "source" }))
            .expect("commit the signed canonical identity patch");
        let canonical_charge = canonical.diagnostics().snapshot_bytes();
        assert!(
            canonical_charge > 0,
            "signed identity patches retain native bytes"
        );

        let cancelled = signed
            .edit(|edit| {
                edit.set(
                    "/sheet[1]/A1",
                    props! { "value" => "transient native value" },
                )?;
                edit.set("/sheet[1]/A1", props! { "value" => "source" })?;
                Ok::<_, litchi_ops::Error>(())
            })
            .expect("commit the signed cancellation");
        assert!(!cancelled.diagnostics().changed());
        assert_eq!(cancelled.diagnostics().operations(), 2);
        assert!(cancelled.patch().is_noop());
        assert!(
            cancelled.diagnostics().snapshot_bytes() >= canonical_charge,
            "a coalesced cancellation must retain at least its canonical identity charge"
        );
        assert_eq!(cancelled.file().to_bytes().unwrap(), bytes);
    }
}

#[test]
fn unsigned_cell_snapshot_budget_is_exact_for_runs_and_ranges() {
    let bytes = fixture();

    for eager in [false, true] {
        let probe = open_owner(&bytes, eager)
            .edit(|edit| edit.set("/sheet[1]/A1", props! { "value" => "changed" }))
            .expect("probe the unsigned coalesced cell patch");
        let charge = probe.diagnostics().snapshot_bytes();
        assert!(charge > 0, "cell runs must retain their native proof");

        let exact = SemanticLimits::builder()
            .max_snapshot_bytes(charge)
            .expect("exact cell-run snapshot limit")
            .build()
            .expect("exact cell-run semantic limits");
        open_owner(&bytes, eager)
            .edit_with_options(EditOptions::new().semantic_limits(exact), |edit| {
                edit.set("/sheet[1]/A1", props! { "value" => "changed" })
            })
            .expect("the exact unsigned coalesced proof budget must fit");

        let under = SemanticLimits::builder()
            .max_snapshot_bytes(charge - 1)
            .expect("under-budget cell-run snapshot limit")
            .build()
            .expect("under-budget cell-run semantic limits");
        let error = open_owner(&bytes, eager)
            .edit_with_options(EditOptions::new().semantic_limits(under), |edit| {
                edit.set("/sheet[1]/A1", props! { "value" => "changed" })
            })
            .expect_err("an unsigned coalesced proof must not bypass the limit");
        assert_eq!(error.code(), Code::LimitExceeded);

        let probe = open_owner(&bytes, eager)
            .edit(|edit| edit.set("Data!A1:B1", props! { "value" => "range" }))
            .expect("probe the unsigned grouped cell patch");
        let charge = probe.diagnostics().snapshot_bytes();
        assert!(
            charge > 0,
            "grouped cell patches must retain their native proof"
        );
        let exact = SemanticLimits::builder()
            .max_snapshot_bytes(charge)
            .expect("exact grouped snapshot limit")
            .build()
            .expect("exact grouped semantic limits");
        open_owner(&bytes, eager)
            .edit_with_options(EditOptions::new().semantic_limits(exact), |edit| {
                edit.set("Data!A1:B1", props! { "value" => "range" })
            })
            .expect("the exact unsigned grouped proof budget must fit");
    }
}

#[test]
fn worksheet_name_transactions_replay_inverse_and_collapse_noops() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original XLSB");
        let commit = original
            .edit(|edit| {
                edit.set(FIRST, props! { "name" => "Input" })?;
                edit.set(SECOND, props! { "name" => "Archive 2026" })
            })
            .expect("stage worksheet renames");
        assert!(commit.diagnostics().changed());
        assert_eq!(
            sheet_property(commit.file(), FIRST, "name").as_deref(),
            Some("Input")
        );
        assert_eq!(
            sheet_property(commit.file(), SECOND, "name").as_deref(),
            Some("Archive 2026")
        );
        assert_eq!(
            commit
                .patch()
                .apply(&original)
                .expect("replay patch")
                .to_bytes()
                .unwrap(),
            commit.file().to_bytes().unwrap()
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .expect("apply inverse patch")
                .to_bytes()
                .unwrap(),
            before
        );

        let noop = original
            .edit(|edit| edit.set(FIRST, props! { "name" => "Data" }))
            .expect("stage same worksheet name");
        assert!(!noop.diagnostics().changed());
        assert!(noop.patch().is_noop());
        assert_eq!(noop.file().to_bytes().unwrap(), before);
    }
}

#[test]
fn worksheet_name_batches_keep_atomic_and_best_effort_boundaries() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original XLSB");
        let commands = || [set_batch(FIRST, "Input"), set_batch(SECOND, "Input")];

        let atomic = original
            .batch(commands(), BatchMode::Atomic)
            .expect_err("atomic duplicate name must roll back");
        assert_eq!(atomic.item(), Some(1));
        assert_eq!(atomic.code(), Code::DuplicateName);
        assert_eq!(original.to_bytes().unwrap(), before);

        let best_effort = original
            .batch(commands(), BatchMode::BestEffort)
            .expect("best-effort duplicate name");
        assert!(matches!(
            &best_effort.value()[0],
            ItemResult::Applied(applied) if applied.changed
        ));
        assert!(matches!(
            &best_effort.value()[1],
            ItemResult::Failed(error)
                if error.item() == Some(1) && error.code() == Code::DuplicateName
        ));
        assert_eq!(
            sheet_property(best_effort.file(), FIRST, "name").as_deref(),
            Some("Input")
        );
        assert_eq!(
            sheet_property(best_effort.file(), SECOND, "name").as_deref(),
            Some("Archive")
        );
        assert_eq!(original.to_bytes().unwrap(), before);
    }
}

#[test]
fn worksheet_name_validation_is_atomic_for_duplicates_keys_controls_and_length() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original XLSB");

        let duplicate = original
            .set(SECOND, props! { "name" => "Data" })
            .expect_err("duplicate worksheet name");
        assert_eq!(duplicate.code(), Code::DuplicateName);
        assert_eq!(duplicate.property(), Some("name"));

        let duplicate_key = original
            .set(FIRST, Props::new().with("name", "One").with("NAME", "Two"))
            .expect_err("duplicate worksheet property key");
        assert_eq!(duplicate_key.code(), Code::InvalidArgument);
        assert_eq!(duplicate_key.property(), Some("name"));

        let invalid_char = original
            .set(FIRST, props! { "name" => "bad/name" })
            .expect_err("slash in worksheet name");
        assert_eq!(invalid_char.code(), Code::InvalidValue);
        assert_eq!(invalid_char.property(), Some("name"));

        let control = original
            .set(FIRST, props! { "name" => "bad\u{3}" })
            .expect_err("control character in worksheet name");
        assert_eq!(control.code(), Code::InvalidValue);
        assert_eq!(control.property(), Some("name"));

        let too_long = "x".repeat(32);
        let length = original
            .set(FIRST, Props::new().with("name", too_long.as_str()))
            .expect_err("worksheet name over the XLSB limit");
        assert_eq!(length.code(), Code::InvalidValue);
        assert_eq!(length.property(), Some("name"));

        assert_eq!(original.to_bytes().unwrap(), before);
    }
}

#[test]
fn worksheet_names_use_unicode_caseless_duplicate_detection() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let composed = original
            .set(FIRST, props! { "name" => "Café" })
            .expect("set composed Unicode worksheet name");
        let duplicate = composed
            .set(SECOND, props! { "name" => "Cafe\u{0301}" })
            .expect_err("canonically equivalent worksheet name must be refused");
        assert_eq!(duplicate.code(), Code::DuplicateName);
        assert_eq!(duplicate.property(), Some("name"));
        assert_eq!(
            sheet_property(&composed, FIRST, "name").as_deref(),
            Some("Café")
        );
        assert_eq!(
            sheet_property(&composed, SECOND, "name").as_deref(),
            Some("Archive")
        );
    }
}

#[test]
fn multi_match_sheet_name_set_refuses_before_any_native_mutation() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original XLSB");
        let error = original
            .set(scoped_sheet_query("sheet"), props! { "name" => "Merged" })
            .expect_err("renaming multiple sheets is ambiguous");
        assert_eq!(error.code(), Code::InvalidArgument);
        assert_eq!(error.property(), Some("name"));
        assert_eq!(original.to_bytes().unwrap(), before);
        assert_eq!(
            sheet_property(&original, FIRST, "name").as_deref(),
            Some("Data")
        );
        assert_eq!(
            sheet_property(&original, SECOND, "name").as_deref(),
            Some("Archive")
        );
    }
}

#[test]
fn unsupported_sheet_properties_are_rejected_before_editor_admission() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original XLSB");
        let error = original
            .set(FIRST, props! { "zoom" => 125 })
            .expect_err("unsupported XLSB sheet property");
        assert_eq!(error.code(), Code::UnsupportedProperty);
        assert_eq!(error.property(), Some("zoom"));
        assert_eq!(original.to_bytes().unwrap(), before);
    }
}

#[test]
fn transaction_operation_and_snapshot_limits_leave_xlsb_input_unchanged() {
    let bytes = fixture();

    for eager in [false, true] {
        let original = open_owner(&bytes, eager);
        let before = original.to_bytes().expect("publish original XLSB");

        let operation_limits = SemanticLimits::builder()
            .max_operations(1)
            .expect("operation limit")
            .build()
            .expect("operation profile");
        let operation_error = original
            .edit_with_options(
                EditOptions::new().semantic_limits(operation_limits),
                |edit| {
                    edit.set(FIRST, props! { "name" => "Input" })?;
                    edit.set(SECOND, props! { "name" => "Archive 2026" })
                },
            )
            .expect_err("second worksheet mutation exceeds operation limit");
        assert_eq!(operation_error.code(), Code::LimitExceeded);
        assert_eq!(original.to_bytes().unwrap(), before);

        let snapshot_limits = SemanticLimits::builder()
            .max_snapshot_bytes(1)
            .expect("snapshot limit")
            .build()
            .expect("snapshot profile");
        let snapshot_error = original
            .edit_with_options(
                EditOptions::new().semantic_limits(snapshot_limits),
                |edit| edit.set(FIRST, props! { "name" => "Input" }),
            )
            .expect_err("worksheet patch exceeds snapshot limit");
        assert_eq!(snapshot_error.code(), Code::LimitExceeded);
        assert_eq!(original.to_bytes().unwrap(), before);
    }
}

#[test]
fn stale_source_is_refused_and_save_noclobber_preserves_destination() {
    let bytes = fixture();
    let directory = tempfile::tempdir().expect("temporary directory");
    let origin = directory.path().join("stale.xlsb");
    std::fs::write(&origin, &bytes).expect("write XLSB origin");
    let stale = litchi_ops::open(&origin).expect("open source-backed XLSB");

    std::fs::write(&origin, b"source changed after open").expect("change XLSB source");
    let error = stale
        .set(FIRST, props! { "name" => "Input" })
        .expect_err("stale source edit");
    assert_eq!(error.code(), Code::IoError);
    let destination = directory.path().join("stale-destination.xlsb");
    std::fs::write(&destination, b"destination before stale save").expect("write destination");
    let error = stale.save_as(&destination).expect_err("stale source save");
    assert_eq!(error.code(), Code::IoError);
    assert_eq!(
        std::fs::read(&destination).expect("read preserved destination"),
        b"destination before stale save"
    );

    for eager in [false, true] {
        let source = open_owner(&bytes, eager);
        let updated = source
            .set(FIRST, props! { "name" => "Input" })
            .expect("rename for no-clobber save");
        let target = directory.path().join(format!("noclobber-{eager}.xlsb"));
        std::fs::write(&target, b"keep this destination").expect("write existing target");
        let error = updated
            .save_with(
                SaveTo::Path(&target),
                &SaveOptions::new().overwrite(Overwrite::Refuse),
            )
            .expect_err("refuse must not clobber an existing XLSB target");
        assert_eq!(error.code(), Code::FileExists);
        assert_eq!(std::fs::read(&target).unwrap(), b"keep this destination");
    }
}

#[test]
fn retained_source_policy_refuses_xlsb_writes_without_changing_reads() {
    let bytes = fixture();

    for eager in [false, true] {
        let file = open_without_source(&bytes, eager);
        assert_eq!(
            sheet_property(&file, FIRST, "name").as_deref(),
            Some("Data")
        );
        let before = file.to_bytes().expect("publish unretained no-op");
        let error = file
            .set(FIRST, props! { "name" => "Input" })
            .expect_err("unretained XLSB cannot be edited");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert!(error.suggestion().is_some());
        assert_eq!(file.to_bytes().unwrap(), before);
    }
}

#[cfg(feature = "encryption")]
#[test]
fn unsupported_reencryption_does_not_publish_plaintext() {
    let file = open_owner(&fixture(), false);
    let directory = tempfile::tempdir().unwrap();
    let target = directory.path().join("encrypted.xlsb");
    std::fs::write(&target, b"existing destination").unwrap();
    let error = file
        .save_with(
            SaveTo::Path(&target),
            &SaveOptions::new().reencrypt(litchi_ops::Password::new("test password".to_owned())),
        )
        .expect_err("XLSB re-encryption remains unsupported");
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(std::fs::read(target).unwrap(), b"existing destination");
}

#[test]
fn worksheet_names_cannot_collide_with_chart_tabs() {
    use litchi_xlsb::builder::chart::{Anchor, Chart};
    use litchi_xlsb::builder::{self, Builder};
    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .unwrap();
    let mut writer = Builder::new();
    writer.add_worksheet(builder::Worksheet::new("Data"));
    writer
        .add_chart_sheet(builder::ChartSheet::new("Café", chart))
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    let bytes = output.into_inner();
    for eager in [false, true] {
        let file = open_owner(&bytes, eager);
        let before = file.to_bytes().unwrap();
        let error = file
            .set(FIRST, props! { "name" => "CAFE\u{0301}" })
            .expect_err("complete tab names are unique");
        assert_eq!(error.code(), Code::DuplicateName);
        assert_eq!(error.property(), Some("name"));
        assert_eq!(file.to_bytes().unwrap(), before);
    }
}
