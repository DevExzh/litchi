//! XLSB P0 writes through the unified API (record 0261): cell formatting,
//! formulas and merged ranges, with read-back, exact no-ops, signature and
//! encryption policy, and transaction replay.
#![cfg(feature = "xlsb")]

use std::io::Cursor;

use litchi_core::edit::RefusalKind;
use litchi_ops::error::Source;
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{Code, EditOptions, File, OpenOptions, Props, SourceMode, props};
use litchi_xlsb::package::PackageError;

fn fixture() -> Vec<u8> {
    use litchi_xlsb::builder::{self, Builder};

    let mut builder = Builder::new();
    let mut data = builder::Worksheet::new("Data");
    data.set_cell(0, 0, 1.5_f64);
    data.set_cell(1, 0, 2.5_f64);
    data.set_cell(0, 1, "label");
    builder.add_worksheet(data);
    builder.add_worksheet(builder::Worksheet::new("Other"));
    let mut output = Cursor::new(Vec::new());
    builder
        .build()
        .expect("build XLSB fixture")
        .write_to(&mut output)
        .expect("serialize XLSB fixture");
    output.into_inner()
}

fn owners(bytes: &[u8]) -> [File; 2] {
    [false, true].map(|eager| {
        litchi_ops::open_bytes_with(
            bytes.to_vec(),
            &OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            }),
        )
        .expect("open XLSB owner")
    })
}

fn property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn reopen(file: &File) -> File {
    litchi_ops::open_bytes(file.to_bytes().expect("publish")).expect("reopen")
}

#[test]
fn cell_formats_set_and_read_back_on_both_owners() {
    let bytes = fixture();
    for file in owners(&bytes) {
        let edited = file
            .set(
                "/Data/A1",
                props! {
                    "font.bold" => true,
                    "font.italic" => true,
                    "underline" => "single",
                    "font.size" => "14pt",
                    "font.name" => "Arial",
                    "font.color" => "1F4E79",
                    "fill" => "FFEE00",
                    "border.bottom" => "thick",
                    "border.bottom.color" => "000000",
                    "numberformat" => "0.000",
                    "alignment.horizontal" => "center",
                    "alignment.wrapText" => true,
                    "alignment.indent" => 2,
                },
            )
            .expect("format A1");
        assert_eq!(file.to_bytes().unwrap(), bytes, "the source is unchanged");
        for file in [&edited, &reopen(&edited)] {
            let get = |key: &str| property(file, "/Data/A1", key);
            assert_eq!(get("font.bold").as_deref(), Some("true"));
            assert_eq!(get("font.italic").as_deref(), Some("true"));
            assert_eq!(get("underline").as_deref(), Some("single"));
            assert_eq!(get("font.size").as_deref(), Some("14pt"));
            assert_eq!(get("font.name").as_deref(), Some("Arial"));
            assert_eq!(get("font.color").as_deref(), Some("#1F4E79"));
            assert_eq!(get("fill").as_deref(), Some("#FFEE00"));
            assert_eq!(get("border.bottom").as_deref(), Some("thick"));
            assert_eq!(get("border.bottom.color").as_deref(), Some("#000000"));
            assert_eq!(get("border.top"), None);
            assert_eq!(get("numberformat").as_deref(), Some("0.000"));
            assert_eq!(get("alignment.horizontal").as_deref(), Some("center"));
            assert_eq!(get("alignment.wrapText").as_deref(), Some("true"));
            assert_eq!(get("alignment.indent").as_deref(), Some("2"));
            // The value is untouched; other cells keep their format.
            assert_eq!(
                file.get("/Data/A1")
                    .unwrap()
                    .view()
                    .text()
                    .unwrap()
                    .as_deref(),
                Some("1.5")
            );
            assert_eq!(property(file, "/Data/A2", "font.bold"), None);
        }

        // A partial change keeps the other attributes.
        let partial = edited
            .set(
                "/Data/A1",
                props! { "font.bold" => false, "border.bottom" => "none" },
            )
            .expect("partial format");
        assert_eq!(property(&partial, "/Data/A1", "font.bold"), None);
        assert_eq!(
            property(&partial, "/Data/A1", "font.italic").as_deref(),
            Some("true")
        );
        assert_eq!(property(&partial, "/Data/A1", "border.bottom"), None);
        assert_eq!(
            property(&partial, "/Data/A1", "fill").as_deref(),
            Some("#FFEE00")
        );
    }
}

#[test]
fn equal_formats_are_exact_no_ops() {
    let bytes = fixture();
    for file in owners(&bytes) {
        let once = file
            .set(
                "/Data/A1",
                props! { "font.bold" => true, "fill" => "FFEE00" },
            )
            .expect("format");
        let published = once.to_bytes().unwrap();
        let again = once
            .set("/Data/A1", props! { "bold" => true, "bgcolor" => "FFEE00" })
            .expect("equal format");
        assert_eq!(again.to_bytes().unwrap(), published);
    }
}

#[test]
fn range_and_query_formats_apply_to_every_selected_cell() {
    let bytes = fixture();
    for file in owners(&bytes) {
        let edited = file
            .set(
                "/Data/A1:B2",
                props! { "border.all" => "thin", "font.bold" => true },
            )
            .expect("format range");
        for cell in ["A1", "A2", "B1", "B2"] {
            let path = format!("/Data/{cell}");
            assert_eq!(
                property(&edited, &path, "font.bold").as_deref(),
                Some("true")
            );
            for side in ["left", "right", "top", "bottom"] {
                assert_eq!(
                    property(&edited, &path, &format!("border.{side}")).as_deref(),
                    Some("thin"),
                    "{cell} {side}"
                );
            }
        }
        // B2 was a blank coordinate; it is now a formatted blank cell.
        assert_eq!(
            property(&edited, "/Data/B2", "empty").as_deref(),
            Some("true")
        );
    }
}

#[test]
fn format_refusals_are_typed_and_change_nothing() {
    let bytes = fixture();
    let file = litchi_ops::open_bytes(bytes.clone()).unwrap();
    for (props, code) in [
        (
            Props::new().with("font.color", "accent1"),
            Code::UnsupportedCapability,
        ),
        (
            Props::new().with("alignment.textRotation", 200),
            Code::InvalidValue,
        ),
        (
            Props::new().with("border.left.color", "FF0000"),
            Code::InvalidArgument,
        ),
        (Props::new().with("underline", "double"), Code::InvalidValue),
        (
            Props::new().with("value", 1).with("formula", "A2"),
            Code::InvalidArgument,
        ),
    ] {
        let error = file.set("/Data/A1", props).expect_err("refused");
        assert_eq!(error.code(), code, "{error:?}");
    }
    assert_eq!(file.to_bytes().unwrap(), bytes);
}

#[test]
fn formulas_write_and_read_back() {
    let bytes = fixture();
    for file in owners(&bytes) {
        let edited = file
            .set("/Data/A3", props! { "formula" => "=SUM(A1:A2)" })
            .expect("formula on a new cell")
            .set("/Data/A2", props! { "formula" => "A1*2" })
            .expect("formula on a value cell");
        for file in [&edited, &reopen(&edited)] {
            assert_eq!(
                property(file, "/Data/A3", "formula").as_deref(),
                Some("SUM(A1:A2)")
            );
            assert_eq!(
                property(file, "/Data/A2", "formula").as_deref(),
                Some("(A1*2)")
            );
            // The previous value is the cached result until Excel recalculates.
            assert_eq!(
                property(file, "/Data/A2", "cachedValue").as_deref(),
                Some("2.5")
            );
        }
        let same = edited
            .set("/Data/A3", props! { "formula" => "SUM(A1:A2)" })
            .expect("equal formula");
        assert_eq!(same.to_bytes().unwrap(), edited.to_bytes().unwrap());

        let error = file
            .set("/Data/A4", props! { "formula" => "Other!A1" })
            .expect_err("cross-sheet formula");
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error:?}");
        let error = file
            .set("/Data/A4", props! { "formula" => "=(" })
            .expect_err("invalid formula");
        assert_eq!(error.code(), Code::InvalidValue, "{error:?}");
    }
}

#[test]
fn merges_set_read_back_and_unmerge_exactly() {
    let bytes = fixture();
    for file in owners(&bytes) {
        let merged = file
            .set("/Data/A1:B2", props! { "merge" => true })
            .expect("merge");
        for file in [&merged, &reopen(&merged)] {
            assert_eq!(
                property(file, "/Data/A1:B2", "merge").as_deref(),
                Some("true")
            );
            assert_eq!(
                property(file, "/Data/A1", "merge").as_deref(),
                Some("A1:B2")
            );
            assert_eq!(
                property(file, "/Data/A1", "mergeAnchor").as_deref(),
                Some("true")
            );
            assert_eq!(property(file, "/Data/A1:B3", "merge"), None);
        }
        let again = merged
            .set("/Data/A1:B2", props! { "merge" => true })
            .expect("equal merge");
        assert_eq!(again.to_bytes().unwrap(), merged.to_bytes().unwrap());

        let before_refusal = merged.to_bytes().unwrap();
        let error = merged
            .set("/Data/B2:C3", props! { "merge" => true })
            .expect_err("overlap");
        assert_eq!(error.code(), Code::InvalidValue, "{error:?}");
        assert!(
            matches!(error.leaf(), Source::Xlsb(PackageError::Refused(refusal))
            if refusal.kind() == RefusalKind::InvalidValue)
        );
        assert_eq!(error.property(), Some("merge"));
        assert!(error.suggestion().is_some());
        assert_eq!(merged.to_bytes().unwrap(), before_refusal);
        let error = merged
            .set("/Data/A1:A3", props! { "merge" => false })
            .expect_err("partial unmerge");
        assert_eq!(error.code(), Code::UnsupportedType, "{error:?}");
        assert!(
            matches!(error.leaf(), Source::Xlsb(PackageError::Refused(refusal))
            if refusal.kind() == RefusalKind::UnsupportedTarget)
        );
        assert_eq!(error.property(), Some("merge"));
        assert!(error.suggestion().is_some());
        assert_eq!(merged.to_bytes().unwrap(), before_refusal);
        let error = merged
            .set("/Data/A1", props! { "merge" => true })
            .expect_err("a cell is not a range");
        assert_eq!(error.code(), Code::UnsupportedElement, "{error:?}");

        let unmerged = merged
            .set("/Data/A1:B2", props! { "merge" => false })
            .expect("unmerge");
        assert_eq!(property(&unmerged, "/Data/A1:B2", "merge"), None);
        assert_eq!(property(&unmerged, "/Data/A1", "merge"), None);
    }
}

#[test]
fn p0_writes_replay_and_invert_in_transactions() {
    let bytes = fixture();
    for file in owners(&bytes) {
        let commit = file
            .edit(|tx| {
                tx.set("/Data/A1", props! { "font.bold" => true })?;
                tx.set("/Data/C1", props! { "formula" => "A1+1" })?;
                tx.set("/Data/A2:B3", props! { "merge" => true })?;
                tx.set("/metadata", props! { "title" => "P0" })?;
                tx.add(
                    "/",
                    litchi_ops::NewNode::new("namedrange")
                        .prop("name", "Input")
                        .prop("ref", "Data!$A$1"),
                )?;
                Ok(())
            })
            .expect("transaction");
        assert!(commit.diagnostics().changed());
        let after = commit.file();
        assert_eq!(
            property(after, "/Data/A1", "font.bold").as_deref(),
            Some("true")
        );
        assert_eq!(
            property(after, "/Data/C1", "formula").as_deref(),
            Some("(A1+1)")
        );
        assert_eq!(
            property(after, "/Data/A2:B3", "merge").as_deref(),
            Some("true")
        );
        assert_eq!(
            property(after, "/namedrange[1]", "ref").as_deref(),
            Some("Data!$A$1")
        );
        let restored = commit.patch().inverse().apply(after).expect("inverse");
        assert_eq!(restored.to_bytes().unwrap(), bytes);
        let replayed = commit.patch().apply(&file).expect("replay");
        assert_eq!(replayed.to_bytes().unwrap(), after.to_bytes().unwrap());
    }
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
    package
        .try_add_part(Box::new(BlobPart::new(
            origin,
            litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            Vec::new(),
        )))
        .expect("signature origin part");
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

#[test]
fn signed_workbooks_refuse_every_p0_write_until_unsigned() {
    let bytes = signed_fixture();
    for file in owners(&bytes) {
        assert!(file.is_signed());
        for (path, props) in [
            ("/Data/A1", Props::new().with("font.bold", true)),
            ("/Data/A1", Props::new().with("formula", "1+1")),
            ("/Data/A1:B2", Props::new().with("merge", true)),
            ("/metadata", Props::new().with("title", "Signed")),
        ] {
            let error = file.set(path, props).expect_err("signed refusal");
            assert_eq!(error.code(), Code::DocumentProtected, "{path}: {error:?}");
        }
        let error = file
            .add(
                "/",
                litchi_ops::NewNode::new("namedrange")
                    .prop("name", "Signed")
                    .prop("ref", "Data!A1"),
            )
            .expect_err("signed refusal");
        assert_eq!(error.code(), Code::DocumentProtected, "{error:?}");
        assert_eq!(file.to_bytes().unwrap(), bytes);
        let File::Signed(signed) = file else {
            panic!("a signed fixture opens as File::Signed");
        };
        let unsigned = signed.unsign();
        let edited = unsigned
            .set("/Data/A1", props! { "font.bold" => true })
            .expect("explicitly unsigned edit");
        assert_eq!(
            property(&edited, "/Data/A1", "font.bold").as_deref(),
            Some("true")
        );
    }
}

#[test]
fn unsigned_metadata_snapshot_charge_is_enforced() {
    let limits = SemanticLimits::builder()
        .max_snapshot_bytes(1)
        .expect("snapshot limit")
        .build()
        .expect("semantic limits");
    for file in owners(&fixture()) {
        let probe = file
            .edit(|edit| edit.set("/metadata", props! { "title" => "measured" }))
            .expect("metadata probe");
        assert!(
            probe.diagnostics().snapshot_bytes() > 0,
            "metadata commits must report their retained property payloads"
        );
        let error = file
            .edit_with_options(EditOptions::new().semantic_limits(limits), |edit| {
                edit.set("/metadata", props! { "title" => "bounded" })
            })
            .expect_err("metadata publication must charge its retained patch");
        assert_eq!(error.code(), Code::LimitExceeded, "{error:?}");

        let exact = SemanticLimits::builder()
            .max_snapshot_bytes(probe.diagnostics().snapshot_bytes())
            .expect("exact snapshot limit")
            .build()
            .expect("exact semantic limits");
        file.edit_with_options(EditOptions::new().semantic_limits(exact), |edit| {
            edit.set("/metadata", props! { "title" => "measured" })
        })
        .expect("source-backed metadata must fit its exact retained-byte budget");
    }
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_xlsb_stays_locked_and_is_never_republished_as_plaintext() {
    let plain = fixture();
    let encrypted = litchi_crypto::ooxml::encrypt(
        plain.clone(),
        "xlsb",
        litchi_crypto::ooxml::Mode::Agile,
        &litchi_crypto::ooxml::Limits::default(),
    )
    .expect("encrypt");
    let locked = litchi_ops::open_bytes(encrypted.clone()).expect("open encrypted");
    let error = locked
        .set("/Data/A1", props! { "font.bold" => true })
        .expect_err("a locked workbook refuses writes");
    assert_eq!(error.code(), Code::DocumentProtected, "{error:?}");

    // Formatting an ordinary workbook never yields encrypted output: XLSB
    // re-encryption remains a typed refusal that publishes nothing.
    let edited = litchi_ops::open_bytes(plain)
        .unwrap()
        .set("/Data/A1", props! { "font.bold" => true })
        .unwrap();
    let directory = tempfile::tempdir().unwrap();
    let target = directory.path().join("encrypted.xlsb");
    let error = edited
        .save_with(
            litchi_ops::SaveTo::Path(&target),
            &litchi_ops::SaveOptions::new().reencrypt(litchi_ops::Password::new("pw".to_owned())),
        )
        .expect_err("XLSB re-encryption");
    assert_eq!(error.code(), Code::UnsupportedCapability, "{error:?}");
    assert!(!target.exists());
}

#[test]
fn named_ranges_add_set_read_back_and_remove() {
    use litchi_ops::NewNode;

    let bytes = fixture();
    for file in owners(&bytes) {
        let added = file
            .add(
                "/",
                NewNode::new("namedrange")
                    .prop("name", "Totals")
                    .prop("ref", "Data!$A$1:$A$2"),
            )
            .expect("add a defined name");
        assert_eq!(added.path.to_string(), "/namedrange[1]");
        let scoped = added
            .file
            .add(
                "/",
                NewNode::new("namedrange")
                    .prop("name", "Local")
                    .prop("ref", "'Other'!B2:C3")
                    .prop("scope", "Data"),
            )
            .expect("add a scoped name");
        assert_eq!(scoped.path.to_string(), "/namedrange[2]");
        for file in [&scoped.file, &reopen(&scoped.file)] {
            assert_eq!(
                property(file, "/namedrange[1]", "name").as_deref(),
                Some("Totals")
            );
            assert_eq!(
                property(file, "/namedrange[1]", "ref").as_deref(),
                Some("Data!$A$1:$A$2")
            );
            assert_eq!(
                property(file, "/namedrange[1]", "scope").as_deref(),
                Some("workbook")
            );
            assert_eq!(
                property(file, "/namedrange[2]", "ref").as_deref(),
                Some("Other!$B$2:$C$3")
            );
            assert_eq!(
                property(file, "/namedrange[2]", "scope").as_deref(),
                Some("Data")
            );
        }

        let moved = scoped
            .file
            .set("/namedrange[1]", props! { "ref" => "Data!B1:B2" })
            .expect("set ref");
        assert_eq!(
            property(&moved, "/namedrange[1]", "ref").as_deref(),
            Some("Data!$B$1:$B$2")
        );
        let same = moved
            .set("/namedrange[1]", props! { "ref" => "Data!$B$1:$B$2" })
            .expect("equal ref");
        assert_eq!(same.to_bytes().unwrap(), moved.to_bytes().unwrap());

        // Only the last name can be removed.
        let error = moved.remove("/namedrange[1]").expect_err("earlier name");
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error:?}");
        let removed = moved.remove("/namedrange[2]").expect("remove last name");
        assert!(removed.get("/namedrange[2]").is_err());
        assert_eq!(
            property(&removed, "/namedrange[1]", "name").as_deref(),
            Some("Totals")
        );
    }
}

#[test]
fn named_range_refusals_are_typed_and_change_nothing() {
    use litchi_ops::NewNode;

    let bytes = fixture();
    let file = litchi_ops::open_bytes(bytes.clone()).unwrap();
    for (node, code) in [
        (
            NewNode::new("namedrange")
                .prop("name", "A1")
                .prop("ref", "Data!A1"),
            Code::InvalidValue,
        ),
        (
            NewNode::new("namedrange")
                .prop("name", "Valid")
                .prop("ref", "A1:B2"),
            Code::InvalidValue,
        ),
        (
            NewNode::new("namedrange")
                .prop("name", "Valid")
                .prop("ref", "Missing!A1"),
            Code::InvalidValue,
        ),
        (
            NewNode::new("namedrange").prop("name", "Valid"),
            Code::InvalidArgument,
        ),
    ] {
        let error = file.add("/", node).expect_err("refused");
        assert_eq!(error.code(), code, "{error:?}");
    }
    assert_eq!(file.to_bytes().unwrap(), bytes);
}
