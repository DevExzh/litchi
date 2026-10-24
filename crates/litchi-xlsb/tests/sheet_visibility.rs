#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "these integration fixtures deliberately fail at the assertion site"
)]

//! Native workbook-tab visibility transactions.
//!
//! The fixtures in this file exercise the complete tab catalog, including
//! chart and unbound tabs.  The low-level rewrites are intentionally small:
//! they create the kinds of source images that Excel can produce while
//! leaving all unrelated BIFF12 records opaque to the transaction.

use std::io::Cursor;

use litchi_core::sheet::WorkbookTrait as _;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part as _, TargetMode};
use litchi_xlsb::Package;
use litchi_xlsb::builder::chart::{Anchor, Chart};
use litchi_xlsb::builder::{self, Builder};
use litchi_xlsb::cell_values::TransferLimits;
use litchi_xlsb::cell_values::structure::Patch;
use litchi_xlsb::package::PackageError;
use litchi_xlsb::raw::{Encoder, Kind as RawKind, Records, kind};
use litchi_xlsb::sheet::State;

fn worksheets(names: &[&str]) -> Package {
    let mut builder = Builder::new();
    for name in names {
        builder.add_worksheet(builder::Worksheet::new(*name));
    }
    builder.build().expect("worksheet fixture")
}

fn chart_first_package() -> Package {
    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .expect("chart fixture");
    let mut builder = Builder::new();
    builder
        .add_chart_sheet(builder::ChartSheet::new("Chart", chart))
        .expect("chart sheet fixture");
    builder.add_worksheet(builder::Worksheet::new("Data"));
    builder.add_worksheet(builder::Worksheet::new("Tail"));
    builder.build().expect("chart-first fixture")
}

fn catalog_states(package: &Package) -> Vec<State> {
    package
        .workbook()
        .expect("workbook projection")
        .sheets()
        .map(|tab| tab.state())
        .collect()
}

fn operation_offset(wire: &[u8]) -> usize {
    let before = usize::try_from(u64::from_le_bytes(
        wire[8..16].try_into().expect("before length"),
    ))
    .expect("before length fits usize");
    let after = usize::try_from(u64::from_le_bytes(
        wire[16..24].try_into().expect("after length"),
    ))
    .expect("after length fits usize");
    32 + before + after
}

#[test]
fn selector_missing_is_distinct_from_a_resolved_noop() {
    let package = worksheets(&["Data", "Tail"]);
    let mut edit = package.edit().expect("edit");
    assert_eq!(
        edit.set_sheet_state("does-not-exist", State::Hidden)
            .expect("missing selector"),
        None
    );
    assert_eq!(
        edit.set_sheet_state(0usize, State::Visible)
            .expect("resolved no-op"),
        Some(false)
    );
}

#[test]
fn ambiguous_catalog_names_are_rejected_before_state_selection() {
    let mut builder = Builder::new();
    builder.add_worksheet(builder::Worksheet::new("DATA"));
    builder.add_worksheet(builder::Worksheet::new("Else"));
    let mut bytes = Cursor::new(Vec::new());
    builder
        .build()
        .expect("distinct source")
        .write_to(&mut bytes)
        .expect("write source");

    let mut opc = OpcPackage::from_bytes(&bytes.into_inner(), litchi_opc::SourceLimits::default())
        .expect("OPC source");
    let workbook_uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let original = opc
        .get_part(&workbook_uri)
        .expect("workbook part")
        .blob()
        .to_vec();
    let old: Vec<_> = "Else".encode_utf16().flat_map(u16::to_le_bytes).collect();
    let replacement: Vec<_> = "data".encode_utf16().flat_map(u16::to_le_bytes).collect();
    let mut rewritten = Vec::new();
    let mut writer = Encoder::new(&mut rewritten, litchi_xlsb::raw::Limits::DEFAULT);
    let mut replaced = false;
    for item in Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("workbook record");
        let mut payload = record.payload().to_vec();
        if record.kind() == kind::BUNDLE_SH && payload.ends_with(&old) {
            let start = payload.len() - old.len();
            payload[start..].copy_from_slice(&replacement);
            replaced = true;
        }
        writer
            .write_record(record.kind(), &payload)
            .expect("rewritten record");
    }
    assert!(replaced);
    opc.get_part_mut(&workbook_uri)
        .expect("workbook part")
        .set_blob(rewritten);

    // Catalog construction is eager.  A malformed duplicate must fail the
    // package boundary instead of allowing a name selector to pick a tab.
    assert!(Package::from_opc(opc).is_err());
}

#[test]
fn repeated_state_setters_last_wins_and_cancel_to_an_exact_source() {
    let package = worksheets(&["Data", "Tail"]);
    let source = package.to_bytes().expect("source bytes");
    let mut edit = package.edit().expect("edit");
    assert_eq!(
        edit.set_sheet_state(0usize, State::Hidden).expect("hide"),
        Some(true)
    );
    assert_eq!(
        edit.set_sheet_state(0usize, State::VeryHidden)
            .expect("very hide"),
        Some(true)
    );
    assert_eq!(
        edit.set_sheet_state(0usize, State::Visible)
            .expect("reveal"),
        Some(true)
    );

    let commit = edit.commit().expect("cancelled state commit");
    assert_eq!(commit.patch().operation_count(), 1);
    assert!(commit.patch().is_noop());
    assert_eq!(
        commit.package().to_bytes().expect("candidate bytes"),
        source
    );
}

#[test]
fn one_edit_can_hide_the_active_tab_and_reveal_the_next_visible_tab() {
    let package = worksheets(&["Data", "Tail"]);
    let mut edit = package.edit().expect("edit");
    assert_eq!(
        edit.set_sheet_state(0usize, State::Hidden)
            .expect("hide active"),
        Some(true)
    );
    assert_eq!(
        edit.set_sheet_state(1usize, State::Visible)
            .expect("reveal replacement"),
        Some(false)
    );
    let commit = edit.commit().expect("state commit");
    assert_eq!(
        catalog_states(commit.package()),
        vec![State::Hidden, State::Visible]
    );
}

#[test]
fn durable_state_wire_round_trips_and_rejects_unknown_tags_and_states() {
    let package = worksheets(&["Data", "Tail"]);
    let mut edit = package.edit().expect("edit");
    edit.set_sheet_state(0usize, State::Hidden)
        .expect("hide active");
    let commit = edit.commit().expect("state commit");
    let wire = commit
        .patch()
        .to_bytes(TransferLimits::DEFAULT)
        .expect("durable state patch");
    let offset = operation_offset(&wire);
    assert_eq!(wire[offset], 5, "visibility operation has its own wire tag");
    assert_eq!(
        Package::from_bytes(
            &package
                .apply(
                    &Patch::from_bytes(
                        &wire,
                        TransferLimits::DEFAULT,
                        litchi_xlsb::raw::Limits::DEFAULT
                    )
                    .expect("decode")
                )
                .map(litchi_xlsb::cell_values::structure::Commit::into_package)
                .expect("replay")
                .to_bytes()
                .expect("replayed bytes"),
        )
        .expect("reopen replayed package")
        .workbook()
        .expect("replayed workbook")
        .sheets()
        .next()
        .expect("first tab")
        .state(),
        State::Hidden
    );

    let mut invalid_tag = wire.clone();
    invalid_tag[offset] = 0xff;
    assert!(
        Patch::from_bytes(
            &invalid_tag,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT
        )
        .is_err()
    );

    let mut invalid_state = wire;
    invalid_state[offset + 1 + 8] = 3;
    assert!(
        Patch::from_bytes(
            &invalid_state,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT
        )
        .is_err()
    );
}

// Reconstruct the historical BundleSh-only after image. New operations also
// update XML dependencies, so simply changing their tag would forge an invalid
// old patch rather than test backward compatibility.
fn metadata_only_name_wire(package: &Package, patch: &Patch) -> Vec<u8> {
    let mut legacy = package.opc_package().clone();
    let renamed = package
        .apply(patch)
        .map(litchi_xlsb::cell_values::structure::Commit::into_package)
        .expect("new rename package");
    let uri = legacy.main_document_part().unwrap().partname().clone();
    let workbook = renamed
        .opc_package()
        .get_part(&uri)
        .unwrap()
        .blob()
        .to_vec();
    legacy.get_part_mut(&uri).unwrap().set_blob(workbook);
    let after = Package::from_opc(legacy).unwrap().to_bytes().unwrap();
    let original = patch.to_bytes(TransferLimits::DEFAULT).unwrap();
    let before_len =
        usize::try_from(u64::from_le_bytes(original[8..16].try_into().unwrap())).unwrap();
    let mut wire = original[..32 + before_len].to_vec();
    wire[16..24].copy_from_slice(&(after.len() as u64).to_le_bytes());
    wire.extend_from_slice(&after);
    wire.extend_from_slice(&original[operation_offset(&original)..]);
    let offset = operation_offset(&wire);
    assert_eq!(wire[offset], 7, "dependency-aware name operation");
    wire[offset] = 6;
    wire
}

#[test]
fn old_rename_wire_tag_and_worksheet_ordinal_remain_compatible() {
    // The new operation addresses the complete catalog, so the first
    // worksheet is catalog position 1 after the chart tab. Rewriting only the
    // tag and payload below exercises the legacy decoder's
    // worksheet-ordinal path while keeping the payload's checked integer
    // encoding unchanged.
    let package = chart_first_package();
    let mut edit = package.edit().expect("edit");
    edit.set_sheet_name(1, "Renamed".to_owned())
        .expect("rename")
        .expect("resolved worksheet selector");
    let commit = edit.commit().expect("rename commit");
    let mut wire = metadata_only_name_wire(&package, commit.patch());
    let offset = operation_offset(&wire);
    assert_eq!(
        wire[offset], 6,
        "historical catalog-position name operation"
    );
    wire[offset] = 1;
    wire[offset + 1..offset + 9].copy_from_slice(&0_u64.to_le_bytes());
    assert_eq!(
        u64::from_le_bytes(wire[offset + 1..offset + 9].try_into().unwrap()),
        0,
        "legacy rename payload addresses a worksheet ordinal"
    );

    let decoded = Patch::from_bytes(
        &wire,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .expect("decode rename");
    let replayed = package
        .apply(&decoded)
        .map(litchi_xlsb::cell_values::structure::Commit::into_package)
        .expect("replay rename");
    assert_eq!(
        replayed
            .workbook()
            .expect("renamed workbook")
            .worksheet_names(),
        &["Renamed".to_owned(), "Tail".to_owned()]
    );
}

#[test]
fn signed_exact_noop_is_preserved_but_visibility_change_is_refused() {
    let package = signed_package();
    let source = package.to_bytes().expect("signed source bytes");

    let mut noop = package.edit().expect("signed no-op edit");
    noop.set_sheet_state(0usize, State::Visible)
        .expect("same state");
    let commit = noop.commit().expect("signed no-op commit");
    assert!(commit.patch().is_noop());
    assert_eq!(
        commit.package().to_bytes().expect("signed no-op bytes"),
        source
    );

    let mut changed = package.edit().expect("signed changed edit");
    changed
        .set_sheet_state(1usize, State::Hidden)
        .expect("hide second tab");
    assert!(matches!(changed.commit(), Err(PackageError::Signed)));
    assert_eq!(package.to_bytes().expect("source remains signed"), source);
}

#[test]
fn all_hidden_source_accepts_an_equal_state_exact_noop() {
    let visible = worksheets(&["Data", "Tail"]);
    let bytes = with_states(
        &visible.to_bytes().expect("source bytes"),
        &[(0, 1), (1, 1)],
    );
    let package = Package::from_bytes(&bytes).expect("all-hidden source");
    let source = package.to_bytes().expect("all-hidden source bytes");
    let mut edit = package.edit().expect("all-hidden edit");
    assert_eq!(
        edit.set_sheet_state(0usize, State::Hidden)
            .expect("equal hidden state"),
        Some(false)
    );
    let commit = edit.commit().expect("all-hidden exact no-op");
    assert!(commit.patch().is_noop());
    assert_eq!(
        commit.package().to_bytes().expect("all-hidden bytes"),
        source
    );
}

#[test]
fn unknown_workbook_records_survive_a_native_state_patch() {
    let source = with_unknown_workbook_record(
        &worksheets(&["Data", "Tail"])
            .to_bytes()
            .expect("source bytes"),
    );
    let package = Package::from_bytes(&source).expect("unknown-record source");
    let mut edit = package.edit().expect("edit");
    edit.set_sheet_state(0usize, State::Hidden)
        .expect("hide active");
    let commit = edit.commit().expect("state commit");
    let uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let workbook = commit
        .package()
        .opc_package()
        .get_part(&uri)
        .expect("workbook part");
    let unknown = Records::new(workbook.blob(), litchi_xlsb::raw::Limits::DEFAULT)
        .map(|record| record.expect("record"))
        .find(|record| record.kind() == RawKind::new(0x3ffe).unwrap())
        .expect("unknown record");
    assert_eq!(unknown.payload(), &[0xde, 0xad, 0xbe, 0xef]);
}

#[test]
fn multi_sheet_visibility_preserves_opaque_records_and_exact_inverse() {
    let source = with_opaque_sheet_records(
        &worksheets(&["Active", "First", "Second"])
            .to_bytes()
            .expect("source bytes"),
        &[1, 2],
    );
    let package = Package::from_bytes(&source).expect("opaque-sheet source");
    let source_bytes = package.to_bytes().expect("canonical source bytes");
    let expected_opaque = [1usize, 2usize].map(|position| {
        let uri = sheet_partname(package.opc_package(), position);
        let blob = package
            .opc_package()
            .get_part(&uri)
            .expect("sheet part")
            .blob();
        record_bytes(blob, opaque_sheet_kind(position))
    });

    // The first tab is the active sentinel.  Both target sheets are hidden
    // in one native transaction, so the payload replacement path updates two
    // physical parts while leaving the selected active view in place.
    let mut edit = package.edit().expect("edit");
    assert_eq!(
        edit.set_sheet_state(1usize, State::Hidden)
            .expect("hide first opaque sheet"),
        Some(true)
    );
    assert_eq!(
        edit.set_sheet_state(2usize, State::Hidden)
            .expect("hide second opaque sheet"),
        Some(true)
    );
    let commit = edit.commit().expect("multi-sheet state commit");
    let hidden = commit.package();
    assert_eq!(
        catalog_states(hidden),
        vec![State::Visible, State::Hidden, State::Hidden]
    );
    assert_eq!(
        hidden
            .workbook()
            .expect("hidden workbook")
            .active_catalog_position(),
        Some(0)
    );
    assert_ne!(sheet_view_flags(hidden, 0) & 0x0040, 0);
    for (position, expected) in [(1usize, &expected_opaque[0]), (2, &expected_opaque[1])] {
        let uri = sheet_partname(hidden.opc_package(), position);
        let blob = hidden
            .opc_package()
            .get_part(&uri)
            .expect("hidden sheet part")
            .blob();
        assert_eq!(
            record_bytes(blob, opaque_sheet_kind(position)).as_slice(),
            expected.as_slice()
        );
        assert_eq!(sheet_view_flags(hidden, position) & 0x0040, 0);
    }

    let restored = hidden
        .apply(&commit.patch().inverse())
        .map(litchi_xlsb::cell_values::structure::Commit::into_package)
        .expect("apply exact inverse");
    assert_eq!(
        restored.to_bytes().expect("inverse bytes"),
        source_bytes,
        "inverse restores the complete package image"
    );
    assert_ne!(sheet_view_flags(&restored, 0) & 0x0040, 0);
    for (position, expected) in [(1usize, &expected_opaque[0]), (2, &expected_opaque[1])] {
        let uri = sheet_partname(restored.opc_package(), position);
        let blob = restored
            .opc_package()
            .get_part(&uri)
            .expect("restored sheet part")
            .blob();
        assert_eq!(
            record_bytes(blob, opaque_sheet_kind(position)).as_slice(),
            expected.as_slice()
        );
        assert_ne!(sheet_view_flags(&restored, position) & 0x0040, 0);
    }
}

#[test]
fn grouped_rename_metadata_rejects_duplicate_final_names() {
    let package = worksheets(&["Data", "Tail"]);
    let mut edit = package.edit().expect("edit");
    edit.set_sheet_name(0, "Collision".to_owned())
        .expect("first rename")
        .expect("resolved worksheet selector");
    assert!(edit.set_sheet_name(1, "Collision".to_owned()).is_err());
    let committed = edit
        .commit()
        .expect("failed setter leaves prior edit intact");
    assert_eq!(
        committed.package().workbook().unwrap().worksheet_names(),
        &["Collision", "Tail"]
    );
}

#[test]
fn chart_visibility_is_native_and_unmodeled_selected_kinds_fail_closed() {
    let bytes = every_kind_workbook();
    let package = Package::from_bytes(&bytes).expect("complete-kind source");

    // Dialog, macro, and international macro tabs have no selected-view
    // codec.  Hiding one must refuse before leaving a selected tab behind.
    for position in [2usize, 3usize, 4usize] {
        let mut edit = package.edit().expect("edit");
        let result = edit.set_sheet_state(position, State::VeryHidden);
        if result
            .expect("catalog selector")
            .is_some_and(|changed| changed)
        {
            assert!(edit.commit().is_err(), "unmodeled kind {position}");
        }
    }

    // The chart sheet is part of the complete catalog and can be selected by
    // a state transition when the active worksheet is hidden in the same edit.
    let mut edit = package.edit().expect("chart edit");
    edit.set_sheet_state(0usize, State::Hidden)
        .expect("hide worksheet");
    edit.set_sheet_state(1usize, State::Visible)
        .expect("reveal chart");
    let commit = edit.commit().expect("chart transition");
    assert_eq!(catalog_states(commit.package())[1], State::Visible);
}

#[test]
fn unbound_tabs_accept_only_very_hidden_state() {
    let bytes = with_null_relationship(&every_kind_workbook(), 3, 2);
    let package = Package::from_bytes(&bytes).expect("unbound source");
    let mut edit = package.edit().expect("edit");
    assert!(edit.set_sheet_state(3usize, State::Hidden).is_err());
    assert_eq!(
        edit.set_sheet_state(3usize, State::VeryHidden)
            .expect("equal very-hidden state"),
        Some(false)
    );
}

#[test]
fn active_relocation_refuses_multiple_chart_views() {
    let bytes = with_duplicate_chart_view(&with_states(&every_kind_workbook(), &[(1, 2)]));
    let package = Package::from_bytes(&bytes).expect("multiple-view source");
    let mut edit = package.edit().expect("edit");
    edit.set_sheet_state(0usize, State::Hidden)
        .expect("hide active worksheet");
    edit.set_sheet_state(1usize, State::Visible)
        .expect("reveal chart");
    assert!(edit.commit().is_err());
}

#[test]
fn active_relocation_refuses_a_missing_chart_view() {
    let bytes = without_chart_views(&with_states(&every_kind_workbook(), &[(1, 2)]));
    let package = Package::from_bytes(&bytes).expect("missing-view source");
    let mut edit = package.edit().expect("edit");
    edit.set_sheet_state(0usize, State::Hidden)
        .expect("hide active worksheet");
    edit.set_sheet_state(1usize, State::Visible)
        .expect("reveal chart");
    assert!(edit.commit().is_err());
}

fn signed_package() -> Package {
    let mut package = worksheets(&["Data", "Tail"]).into_opc();
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
        .expect("signature origin");
    package
        .try_add_part(Box::new(BlobPart::new(
            signature,
            litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
            b"<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>".to_vec(),
        )))
        .expect("signature part");
    package.relate_to(
        "_xmlsignatures/origin.sigs",
        litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    let package = Package::from_opc(package).expect("signed package");
    assert!(package.opc_package().is_signed());
    package
}

fn with_states(bytes: &[u8], states: &[(usize, u32)]) -> Vec<u8> {
    let mut package =
        OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).expect("OPC source");
    let workbook_uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let original = package
        .get_part(&workbook_uri)
        .expect("workbook part")
        .blob()
        .to_vec();
    let mut rewritten = Vec::new();
    let mut position = 0usize;
    let mut writer = Encoder::new(&mut rewritten, litchi_xlsb::raw::Limits::DEFAULT);
    for item in Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("workbook record");
        let mut payload = record.payload().to_vec();
        if record.kind() == kind::BUNDLE_SH {
            if let Some((_, state)) = states.iter().find(|(at, _)| *at == position) {
                payload[..4].copy_from_slice(&state.to_le_bytes());
            }
            position += 1;
        }
        writer
            .write_record(record.kind(), &payload)
            .expect("rewritten record");
    }
    package
        .get_part_mut(&workbook_uri)
        .expect("workbook part")
        .set_blob(rewritten);
    let mut output = Vec::new();
    package.write_to(&mut output).expect("serialize OPC");
    output
}

fn with_unknown_workbook_record(bytes: &[u8]) -> Vec<u8> {
    let mut package =
        OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).expect("OPC source");
    let workbook_uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let original = package
        .get_part(&workbook_uri)
        .expect("workbook part")
        .blob()
        .to_vec();
    let unknown = RawKind::new(0x3ffe).expect("unknown record kind");
    let mut rewritten = Vec::new();
    let mut inserted = false;
    let mut writer = Encoder::new(&mut rewritten, litchi_xlsb::raw::Limits::DEFAULT);
    for item in Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("workbook record");
        if !inserted && record.kind() == kind::END_BOOK {
            writer
                .write_record(unknown, &[0xde, 0xad, 0xbe, 0xef])
                .expect("unknown record");
            inserted = true;
        }
        writer
            .write_record(record.kind(), record.payload())
            .expect("rewritten record");
    }
    assert!(inserted);
    package
        .get_part_mut(&workbook_uri)
        .expect("workbook part")
        .set_blob(rewritten);
    let mut output = Vec::new();
    package.write_to(&mut output).expect("serialize OPC");
    output
}

fn bundle_entries(package: &OpcPackage) -> Vec<(String, String)> {
    let workbook_uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    Records::new(
        package
            .get_part(&workbook_uri)
            .expect("workbook part")
            .blob(),
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .map(|record| record.expect("workbook record"))
    .filter(|record| record.kind() == kind::BUNDLE_SH)
    .map(|record| {
        let payload = record.payload();
        let current_id = u32::from_le_bytes(payload[4..8].try_into().unwrap());
        let strings_offset = if (1..=0xffff).contains(&current_id) {
            8
        } else {
            12
        };
        let rel_units = u32::from_le_bytes(
            payload[strings_offset..strings_offset + 4]
                .try_into()
                .unwrap(),
        ) as usize;
        let rel_start = strings_offset + 4;
        let rel_end = rel_start + rel_units * 2;
        let rel_id = String::from_utf16(
            &payload[rel_start..rel_end]
                .chunks_exact(2)
                .map(|unit| u16::from_le_bytes([unit[0], unit[1]]))
                .collect::<Vec<_>>(),
        )
        .unwrap();
        let name_units =
            u32::from_le_bytes(payload[rel_end..rel_end + 4].try_into().unwrap()) as usize;
        let name_start = rel_end + 4;
        let name_end = name_start + name_units * 2;
        let name = String::from_utf16(
            &payload[name_start..name_end]
                .chunks_exact(2)
                .map(|unit| u16::from_le_bytes([unit[0], unit[1]]))
                .collect::<Vec<_>>(),
        )
        .unwrap();
        (name, rel_id)
    })
    .collect()
}

fn sheet_partname(package: &OpcPackage, position: usize) -> PackURI {
    let workbook_uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let workbook_part = package.get_part(&workbook_uri).expect("workbook part");
    let entries = bundle_entries(package);
    workbook_part
        .rels()
        .get(&entries[position].1)
        .expect("sheet relationship")
        .target_partname()
        .expect("sheet part")
}

fn opaque_sheet_kind(position: usize) -> RawKind {
    RawKind::new(match position {
        1 => 0x0ffe,
        2 => 0x0ffd,
        _ => panic!("unexpected opaque sheet position {position}"),
    })
    .expect("opaque record kind")
}

fn record_bytes(source: &[u8], target: RawKind) -> Vec<u8> {
    let mut records = Records::new(source, litchi_xlsb::raw::Limits::DEFAULT);
    while let Some(item) = records.next() {
        let record = item.expect("record");
        if record.kind() == target {
            return source[record.offset()..records.offset()].to_vec();
        }
    }
    panic!("record {target} not found");
}

fn sheet_view_flags(package: &Package, position: usize) -> u16 {
    let uri = sheet_partname(package.opc_package(), position);
    let blob = package
        .opc_package()
        .get_part(&uri)
        .expect("sheet part")
        .blob();
    Records::new(blob, litchi_xlsb::raw::Limits::DEFAULT)
        .map(|item| item.expect("sheet record"))
        .find(|record| record.kind() == kind::BEGIN_WS_VIEW)
        .map(|record| {
            u16::from_le_bytes(
                record
                    .payload()
                    .get(..2)
                    .expect("worksheet view flags")
                    .try_into()
                    .expect("worksheet view flags length"),
            )
        })
        .expect("worksheet view")
}

fn with_opaque_sheet_records(bytes: &[u8], positions: &[usize]) -> Vec<u8> {
    let mut package =
        OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).expect("OPC source");
    for &position in positions {
        let uri = sheet_partname(&package, position);
        let original = package.get_part(&uri).expect("sheet part").blob().to_vec();
        let mut opaque = Vec::new();
        let payload: &[u8] = match position {
            1 => b"opaque-first-sheet-record",
            2 => b"opaque-second-sheet-record",
            _ => panic!("unexpected opaque sheet position {position}"),
        };
        Encoder::new(&mut opaque, litchi_xlsb::raw::Limits::DEFAULT)
            .write_record(opaque_sheet_kind(position), payload)
            .expect("opaque record");
        let end_sheet = Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT)
            .map(|item| item.expect("sheet record"))
            .find(|record| record.kind() == kind::END_SHEET)
            .expect("end sheet record")
            .offset();
        let mut rewritten = Vec::with_capacity(original.len() + opaque.len());
        rewritten.extend_from_slice(&original[..end_sheet]);
        rewritten.extend_from_slice(&opaque);
        rewritten.extend_from_slice(&original[end_sheet..]);
        package
            .get_part_mut(&uri)
            .expect("sheet part")
            .set_blob(rewritten);
    }
    let mut output = Vec::new();
    package.write_to(&mut output).expect("serialize OPC");
    output
}

fn every_kind_workbook() -> Vec<u8> {
    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .expect("chart fixture");
    let mut builder = Builder::new();
    builder.add_worksheet(builder::Worksheet::new("Data"));
    builder
        .add_chart_sheet(builder::ChartSheet::new("Sales Chart", chart))
        .expect("chart sheet");
    for name in ["Dialog", "Macro", "Intl"] {
        builder.add_worksheet(builder::Worksheet::new(name));
    }
    builder.add_worksheet(builder::Worksheet::new("Tail"));
    let mut bytes = Cursor::new(Vec::new());
    builder
        .build()
        .expect("all-kind package")
        .write_to(&mut bytes)
        .expect("all-kind bytes");
    let mut bytes = bytes.into_inner();
    for (name, relationship_type, content_type) in [
        (
            "Dialog",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
            "application/vnd.ms-excel.dialogsheet",
        ),
        (
            "Macro",
            "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
            "application/vnd.ms-excel.macrosheet",
        ),
        (
            "Intl",
            "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet",
            "application/vnd.ms-excel.intlmacrosheet",
        ),
    ] {
        bytes = retype_sheet(&bytes, name, relationship_type, content_type);
    }
    with_states(&bytes, &[(1, 2), (3, 1), (5, 1)])
}

fn retype_sheet(bytes: &[u8], name: &str, relationship_type: &str, content_type: &str) -> Vec<u8> {
    let mut package =
        OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).expect("OPC source");
    let workbook_uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let entries = bundle_entries(&package);
    let position = entries
        .iter()
        .position(|(sheet_name, _)| sheet_name == name)
        .expect("sheet name");
    let sheet_uri = sheet_partname(&package, position);
    let (relationship_id, target_ref) = package
        .get_part(&workbook_uri)
        .expect("workbook part")
        .rels()
        .iter()
        .find_map(|relationship| {
            (relationship.target_partname().ok()? == sheet_uri).then(|| {
                (
                    relationship.r_id().to_owned(),
                    relationship.target_ref().to_owned(),
                )
            })
        })
        .expect("sheet relationship");
    let workbook = package.get_part_mut(&workbook_uri).expect("workbook part");
    workbook
        .rels_mut()
        .remove(&relationship_id)
        .expect("remove sheet relationship");
    workbook.rels_mut().add_relationship(
        relationship_type.to_owned(),
        target_ref,
        relationship_id,
        false,
    );
    package
        .get_part_mut(&sheet_uri)
        .expect("sheet part")
        .set_content_type(content_type.to_owned())
        .expect("content type");
    let mut output = Vec::new();
    package.write_to(&mut output).expect("serialize OPC");
    output
}

fn with_null_relationship(bytes: &[u8], position: usize, state: u32) -> Vec<u8> {
    let mut package =
        OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).expect("OPC source");
    let workbook_uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let original = package
        .get_part(&workbook_uri)
        .expect("workbook part")
        .blob()
        .to_vec();
    let mut rewritten = Vec::new();
    let mut current_position = 0usize;
    let mut writer = Encoder::new(&mut rewritten, litchi_xlsb::raw::Limits::DEFAULT);
    for item in Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("workbook record");
        let mut payload = record.payload().to_vec();
        if record.kind() == kind::BUNDLE_SH {
            if current_position == position {
                payload[..4].copy_from_slice(&state.to_le_bytes());
                let current_id = u32::from_le_bytes(payload[4..8].try_into().unwrap());
                let strings_offset = if (1..=0xffff).contains(&current_id) {
                    8
                } else {
                    12
                };
                let rel_units = u32::from_le_bytes(
                    payload[strings_offset..strings_offset + 4]
                        .try_into()
                        .unwrap(),
                ) as usize;
                let rel_end = strings_offset + 4 + rel_units * 2;
                let mut null_payload = Vec::with_capacity(payload.len() - rel_units * 2);
                null_payload.extend_from_slice(&payload[..strings_offset]);
                null_payload.extend_from_slice(&u32::MAX.to_le_bytes());
                null_payload.extend_from_slice(&payload[rel_end..]);
                payload = null_payload;
            }
            current_position += 1;
        }
        writer
            .write_record(record.kind(), &payload)
            .expect("rewritten record");
    }
    package
        .get_part_mut(&workbook_uri)
        .expect("workbook part")
        .set_blob(rewritten);
    let mut output = Vec::new();
    package.write_to(&mut output).expect("serialize OPC");
    output
}

fn with_duplicate_chart_view(bytes: &[u8]) -> Vec<u8> {
    let mut package =
        OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).expect("OPC source");
    let uri = sheet_partname(&package, 1);
    let original = package.get_part(&uri).expect("chart part").blob().to_vec();
    let mut rewritten = Vec::new();
    let mut view_payload = None;
    let mut inserted = false;
    let mut writer = Encoder::new(&mut rewritten, litchi_xlsb::raw::Limits::DEFAULT);
    for item in Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("chart record");
        if record.kind() == kind::BEGIN_CS_VIEW {
            view_payload = Some(record.payload().to_vec());
        }
        writer
            .write_record(record.kind(), record.payload())
            .expect("chart record");
        if record.kind() == kind::END_CS_VIEW && !inserted {
            writer
                .write_record(kind::BEGIN_CS_VIEW, view_payload.as_deref().unwrap())
                .expect("duplicate chart view");
            writer
                .write_record(kind::END_CS_VIEW, &[])
                .expect("duplicate chart view end");
            inserted = true;
        }
    }
    assert!(inserted);
    package
        .get_part_mut(&uri)
        .expect("chart part")
        .set_blob(rewritten);
    let mut output = Vec::new();
    package.write_to(&mut output).expect("serialize OPC");
    output
}

fn without_chart_views(bytes: &[u8]) -> Vec<u8> {
    let mut package =
        OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).expect("OPC source");
    let uri = sheet_partname(&package, 1);
    let original = package.get_part(&uri).expect("chart part").blob().to_vec();
    let mut rewritten = Vec::new();
    let mut in_views = false;
    let mut removed = false;
    let mut writer = Encoder::new(&mut rewritten, litchi_xlsb::raw::Limits::DEFAULT);
    for item in Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("chart record");
        if record.kind() == kind::BEGIN_CS_VIEWS {
            in_views = true;
            removed = true;
            continue;
        }
        if in_views {
            if record.kind() == kind::END_CS_VIEWS {
                in_views = false;
            }
            continue;
        }
        writer
            .write_record(record.kind(), record.payload())
            .expect("chart record");
    }
    assert!(removed && !in_views);
    package
        .get_part_mut(&uri)
        .expect("chart part")
        .set_blob(rewritten);
    let mut output = Vec::new();
    package.write_to(&mut output).expect("serialize OPC");
    output
}

#[test]
fn canonical_name_selectors_track_staged_names_and_exact_cancellation() {
    let package = chart_first_package();
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit().unwrap();
    assert_eq!(
        edit.set_sheet_name("absent", "Bad/Name".into()).unwrap(),
        None
    );
    assert_eq!(
        edit.set_sheet_name("chart", "Chart".into()).unwrap(),
        Some(false)
    );
    assert_eq!(
        edit.set_sheet_name("Chart", "Dashboard".into()).unwrap(),
        Some(true)
    );
    assert_eq!(
        edit.set_sheet_name("Dashboard", "Other".into()).unwrap(),
        None
    );
    assert_eq!(
        edit.set_sheet_name(0, "Dashboard".into()).unwrap(),
        Some(false)
    );
    assert!(edit.set_sheet_name(1, "dashboard".into()).is_err());
    assert_eq!(
        edit.set_sheet_name("Chart", "Chart".into()).unwrap(),
        Some(true)
    );
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), before);
}

#[test]
fn every_catalog_kind_supports_name_patch_replay_and_inverse() {
    let package =
        Package::from_bytes(&with_null_relationship(&every_kind_workbook(), 3, 2)).unwrap();
    let before = package.to_bytes().unwrap();
    for position in 0..6 {
        let mut edit = package.edit().unwrap();
        assert_eq!(
            edit.set_sheet_name(position, "Renamed".into()).unwrap(),
            Some(true)
        );
        let commit = edit.commit().unwrap();
        let wire = commit.patch().to_bytes(TransferLimits::DEFAULT).unwrap();
        let patch = Patch::from_bytes(
            &wire,
            TransferLimits::DEFAULT,
            litchi_xlsb::raw::Limits::DEFAULT,
        )
        .unwrap();
        let replay = package
            .apply(&patch)
            .map(litchi_xlsb::cell_values::structure::Commit::into_package)
            .unwrap();
        assert_eq!(
            replay
                .workbook()
                .unwrap()
                .sheets()
                .nth(position)
                .unwrap()
                .name(),
            "Renamed"
        );
        assert_eq!(
            replay
                .apply(&patch.inverse())
                .map(litchi_xlsb::cell_values::structure::Commit::into_package)
                .unwrap()
                .to_bytes()
                .unwrap(),
            before
        );
    }
}

#[test]
fn legacy_and_canonical_name_patches_share_merge_identity() {
    let package = chart_first_package();
    let mut left = package.edit().unwrap();
    left.set_sheet_name(1, "Left".into()).unwrap();
    let left = left.commit().unwrap();
    let mut wire = metadata_only_name_wire(&package, left.patch());
    let offset = operation_offset(&wire);
    wire[offset] = 1;
    wire[offset + 1..offset + 9].copy_from_slice(&0_u64.to_le_bytes());
    let legacy = Patch::from_bytes(
        &wire,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .unwrap();
    let canonical = legacy.to_bytes(TransferLimits::DEFAULT).unwrap();
    assert_eq!(canonical[operation_offset(&canonical)], 6);
    let mut right = package.edit().unwrap();
    right.set_sheet_name(1, "Right".into()).unwrap();
    let right = right.commit().unwrap();
    let merge = legacy.merge_three_way(right.patch()).unwrap();
    assert_eq!(merge.conflicts().len(), 1);
    assert_eq!(merge.conflicts()[0].sheet_index(), 1);
}

#[test]
fn historical_catalog_name_patch_retains_metadata_only_after_image() {
    let package = chart_first_package();
    let mut edit = package.edit().unwrap();
    edit.set_sheet_name(1, "Renamed".into()).unwrap();
    let commit = edit.commit().unwrap();
    let wire = metadata_only_name_wire(&package, commit.patch());
    let decoded = Patch::from_bytes(
        &wire,
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .unwrap();
    let replay = package
        .apply(&decoded)
        .map(litchi_xlsb::cell_values::structure::Commit::into_package)
        .unwrap();
    assert_eq!(replay.workbook().unwrap().worksheet_names()[0], "Renamed");
    let app = PackURI::new("/docProps/app.xml").unwrap();
    assert_eq!(
        replay.opc_package().get_part(&app).unwrap().blob(),
        package.opc_package().get_part(&app).unwrap().blob()
    );
    assert_ne!(
        commit
            .package()
            .opc_package()
            .get_part(&app)
            .unwrap()
            .blob(),
        package.opc_package().get_part(&app).unwrap().blob()
    );
    assert_eq!(
        replay
            .apply(&decoded.inverse())
            .map(litchi_xlsb::cell_values::structure::Commit::into_package)
            .unwrap()
            .to_bytes()
            .unwrap(),
        package.to_bytes().unwrap()
    );
}

#[test]
fn dependency_name_patch_quotes_formulas_and_preserves_unrelated_payloads() {
    let package = chart_first_package();
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit().unwrap();
    edit.set_sheet_name("Data", "O'Brien & Co".into()).unwrap();
    let commit = edit.commit().unwrap();
    let patch = Patch::from_bytes(
        &commit.patch().to_bytes(TransferLimits::DEFAULT).unwrap(),
        TransferLimits::DEFAULT,
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .unwrap();
    let replay = package
        .apply(&patch)
        .map(litchi_xlsb::cell_values::structure::Commit::into_package)
        .unwrap();
    for metadata in package.opc_package().iter_parts() {
        let part = package.opc_package().get_part(metadata.partname()).unwrap();
        let changed = replay.opc_package().get_part(part.partname()).unwrap();
        if part.content_type()
            == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
        {
            let source = std::str::from_utf8(part.blob()).unwrap();
            let expected = source.replace("Data!", "&apos;O&apos;&apos;Brien &amp; Co&apos;!");
            assert_eq!(changed.blob(), expected.as_bytes());
        } else if part.partname().as_str() == "/docProps/app.xml" {
            let source = std::str::from_utf8(part.blob()).unwrap();
            assert_eq!(
                changed.blob(),
                source
                    .replace(">Data<", ">O&apos;Brien &amp; Co<")
                    .as_bytes()
            );
        } else if part.partname().as_str() != "/xl/workbook.bin" {
            assert_eq!(part.blob(), changed.blob(), "{}", part.partname());
        }
    }
    assert_eq!(
        replay
            .apply(&patch.inverse())
            .map(litchi_xlsb::cell_values::structure::Commit::into_package)
            .unwrap()
            .to_bytes()
            .unwrap(),
        before
    );
}
