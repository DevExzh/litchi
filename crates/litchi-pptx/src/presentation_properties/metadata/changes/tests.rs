#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::{BlobPart, OpcPackage, PackURI};

use super::package::*;
use super::*;

const MAIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const PC: &str = "http://schemas.microsoft.com/office/powerpoint/2013/main/command";
const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn package() -> OpcPackage {
    let mut package = OpcPackage::new();
    package.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument".into(),
        "ppt/presentation.xml".into(),
        "rId1".into(),
        false,
    );
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/presentation.xml").unwrap(),
        MAIN_CONTENT_TYPE.into(),
        b"<p:presentation/>".to_vec(),
    )));
    package
}

fn value() -> Part {
    let xml = format!(
        r#"<pc:chgInfo xmlns:pc="{PC}" xmlns:p="{P}" xmlns:v="urn:vendor" xmlns:r="{R}">
  <pc:docChgLst>
    <pc:chgData name="Ada" userId="ada@example.test" email="ada@example.test"><a:extLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ext uri="urn:review"><v:opaque r:id="rIdNeverFetched" href="https://example.invalid/not-opened"/></a:ext></a:extLst></pc:chgData>
    <pc:docChg chg="addSld"><pc:docMkLst><pc:docMk/></pc:docMkLst><p:extLst><p:ext uri="urn:future"><v:payload r:id="rIdAlsoNeverFetched"/></p:ext></p:extLst></pc:docChg>
    <p:extLst><p:ext uri="urn:list"><v:list><v:unknown>keep</v:unknown></v:list></p:ext></p:extLst>
  </pc:docChgLst>
</pc:chgInfo>"#
    );
    Part {
        relationship_id: "rIdChanges".into(),
        part_name: "/ppt/changesInfo.xml".into(),
        changes_information: Info::parse(xml.as_bytes(), Limits::DEFAULT).unwrap(),
    }
}

fn stored_package() -> OpcPackage {
    let mut package = package();
    store(&mut package, &value(), Limits::DEFAULT).unwrap();
    package
}

#[test]
fn no_op_commit_and_patch_preserve_exact_source() {
    let mut package = stored_package();
    let before = load_snapshot(&package, Limits::DEFAULT).unwrap().unwrap();
    let source = before.source_xml().to_vec();
    let commit = before.edit().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.snapshot().source_xml(), source.as_slice());

    let result = apply_commit(&mut package, commit).unwrap();
    assert_eq!(result.source_xml(), source.as_slice());
    assert_eq!(
        package
            .get_part(&PackURI::new("/ppt/changesInfo.xml").unwrap())
            .unwrap()
            .blob(),
        source.as_slice()
    );
}

#[test]
fn typed_edits_preserve_opaque_commands_and_relationship_looking_xml() {
    let mut package = stored_package();
    let snapshot = load_snapshot(&package, Limits::DEFAULT).unwrap().unwrap();
    let mut edit = snapshot.edit();
    edit.edit_author(0, |author| {
        author.name = Some("Grace".into());
        Ok(())
    })
    .unwrap();
    assert!(
        edit.set_change_kinds(0, 0, vec![Kind::DeleteSlide])
            .unwrap()
    );
    let commit = edit.commit().unwrap();
    let output = String::from_utf8_lossy(commit.snapshot().source_xml());
    assert!(output.contains(r#"name="Grace""#));
    assert!(output.contains(r#"chg="delSld""#));
    assert!(output.contains(r#"r:id="rIdNeverFetched""#));
    assert!(output.contains(r#"r:id="rIdAlsoNeverFetched""#));
    assert!(output.contains("https://example.invalid/not-opened"));
    assert!(output.contains("<v:unknown>keep</v:unknown>"));

    let relationship_before = {
        let presentation = package.main_document_part().unwrap();
        let relationship = presentation.rels().get("rIdChanges").unwrap();
        (
            relationship.reltype().to_owned(),
            relationship.target_ref().to_owned(),
            relationship.is_external(),
        )
    };
    let after = apply_commit(&mut package, commit).unwrap();
    assert_eq!(
        after.info().change_lists[0].author.as_ref().unwrap().name,
        Some("Grace".into())
    );
    let presentation = package.main_document_part().unwrap();
    let relationship = presentation.rels().get("rIdChanges").unwrap();
    assert_eq!(relationship.reltype(), relationship_before.0.as_str());
    assert_eq!(relationship.target_ref(), relationship_before.1.as_str());
    assert_eq!(relationship.is_external(), relationship_before.2);
}

#[test]
fn stale_patch_rejection_is_atomic_and_inverse_restores_original_bytes() {
    let mut package = stored_package();
    let original = load_snapshot(&package, Limits::DEFAULT).unwrap().unwrap();
    let mut edit = original.edit();
    edit.edit_author(0, |author| {
        author.name = Some("Grace".into());
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.patch().clone();

    let part = package
        .get_part_mut(&PackURI::new("/ppt/changesInfo.xml").unwrap())
        .unwrap();
    let mut stale = part.blob().to_vec();
    stale.extend_from_slice(b" ");
    part.set_blob(stale);
    let stale_source = package
        .get_part(&PackURI::new("/ppt/changesInfo.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    assert!(patch.apply(&mut package).is_err());
    assert_eq!(
        package
            .get_part(&PackURI::new("/ppt/changesInfo.xml").unwrap())
            .unwrap()
            .blob(),
        stale_source.as_slice()
    );

    let mut package = stored_package();
    let committed = patch.apply(&mut package).unwrap();
    assert_eq!(
        committed.info().change_lists[0]
            .author
            .as_ref()
            .unwrap()
            .name,
        Some("Grace".into())
    );
    let inverse = patch.inverse();
    let restored = inverse.apply(&mut package).unwrap();
    assert_eq!(restored.source_xml(), original.source_xml());
}

#[test]
fn failed_typed_edits_do_not_mutate_the_staged_snapshot() {
    let snapshot = load_snapshot(&stored_package(), Limits::DEFAULT)
        .unwrap()
        .unwrap();
    let mut edit = snapshot.edit();
    let before = edit.snapshot().clone();
    assert!(
        edit.set_change_kinds(0, 0, vec![Kind::AddSlide, Kind::AddSlide])
            .is_err()
    );
    assert_eq!(edit.snapshot(), &before);
    assert!(
        edit.edit_author(0, |author| {
            author.change_id = Some("not-a-guid".into());
            Ok(())
        })
        .is_err()
    );
    assert_eq!(edit.snapshot(), &before);
}

#[test]
fn create_read_update_delete_is_atomic_and_rejects_unexpected_inbound_edges() {
    let mut package = package();
    assert!(load_snapshot(&package, Limits::DEFAULT).unwrap().is_none());
    store(&mut package, &value(), Limits::DEFAULT).unwrap();
    assert!(load_snapshot(&package, Limits::DEFAULT).unwrap().is_some());
    let removed = remove(&mut package, Limits::DEFAULT).unwrap().unwrap();
    assert_eq!(removed.part_name, "/ppt/changesInfo.xml");
    assert!(load_snapshot(&package, Limits::DEFAULT).unwrap().is_none());

    let mut package = stored_package();
    let target = PackURI::new("/ppt/changesInfo.xml").unwrap();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/other.xml").unwrap(),
        "application/xml".into(),
        Vec::new(),
    )));
    package
        .get_part_mut(&PackURI::new("/ppt/other.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            RELATIONSHIP_TYPE.into(),
            target.relative_ref("/ppt/"),
            "rIdOther".into(),
            false,
        );
    assert!(load_snapshot(&package, Limits::DEFAULT).is_err());
    let part_count = package.part_count();
    assert!(remove(&mut package, Limits::DEFAULT).is_err());
    assert_eq!(package.part_count(), part_count);
}

const LIST_DATA_ORDER: &str = r#"Invalid("docChgLst chgData is duplicated or out of order")"#;
const LIST_CHANGE_ORDER: &str = r#"Invalid("docChg is out of order")"#;
const LIST_EXTENSION_ORDER: &str = r#"Invalid("docChgLst extLst is duplicated or out of order")"#;
const CHANGE_ORDER: &str = r#"Invalid("docChg children are duplicated or out of order")"#;
const CHANGE_DUPLICATE_MONIKER: &str = r#"Invalid("docChg has duplicate docMkLst")"#;
const CHANGE_WITHOUT_MONIKER: &str = r#"Invalid("docChg requires docMkLst")"#;

/// A Changes Information part around `lists`.
fn changes_information(lists: &str) -> String {
    format!(
        r#"<pc:chgInfo xmlns:pc="{PC}" xmlns:p="{P}" xmlns:v="urn:vendor">{lists}</pc:chgInfo>"#
    )
}

/// The parsed part, or its refusal with the error variant and exact message.
fn parsed(xml: &str) -> Result<Info, String> {
    Info::parse(xml.as_bytes(), Limits::DEFAULT).map_err(|error| format!("{error:?}"))
}

/// `"ok"`, or the refusal with the error variant and exact message.
fn outcome(xml: &str) -> String {
    match parsed(xml) {
        Ok(_) => "ok".into(),
        Err(error) => error,
    }
}

/// Every sequence of one to `longest` letters of `alphabet`, repeats included.
fn sequences(alphabet: &str, longest: usize) -> Vec<String> {
    let mut all = Vec::new();
    let mut last = vec![String::new()];
    for _ in 0..longest {
        let mut next = Vec::new();
        for prefix in &last {
            for letter in alphabet.chars() {
                next.push(format!("{prefix}{letter}"));
            }
        }
        all.extend(next.iter().cloned());
        last = next;
    }
    all
}

/// Direct `pc:docChgLst` children by letter: `A` chgData, `G` a complete
/// docChg, `X` extLst and `U` a child the list does not know.
fn list_children(sequence: &str) -> String {
    sequence
        .chars()
        .map(|letter| match letter {
            'A' => r#"<pc:chgData name="Ada"/>"#,
            'G' => r#"<pc:docChg chg="addSld"><pc:docMkLst><pc:docMk/></pc:docMkLst></pc:docChg>"#,
            'X' => r#"<p:extLst><p:ext uri="urn:list"><v:data/></p:ext></p:extLst>"#,
            'U' => "<pc:other/>",
            other => panic!("unknown docChgLst child letter {other}"),
        })
        .collect()
}

fn change_list(sequence: &str) -> String {
    format!("<pc:docChgLst>{}</pc:docChgLst>", list_children(sequence))
}

#[test]
fn change_list_children_follow_author_then_changes_then_extensions() {
    // Every sequence of up to three children that is accepted, or refused at
    // a docChg or at an extLst. Every other sequence is refused at a chgData.
    let accepted = [
        "A", "G", "X", "AG", "AX", "GG", "GX", "AGG", "AGX", "GGG", "GGX",
    ];
    let refused_at_change = ["XG", "AXG", "GXG", "XGA", "XGG", "XGX"];
    let refused_at_extension = ["XX", "AXX", "GXX", "XXA", "XXG", "XXX"];
    let mut refused_at_data = 0;
    for sequence in sequences("AGX", 3) {
        let xml = changes_information(&change_list(&sequence));
        if accepted.contains(&sequence.as_str()) {
            let info = parsed(&xml).unwrap_or_else(|error| panic!("{sequence}: {error}"));
            assert_eq!(info.change_lists.len(), 1, "children {sequence}");
            let list = &info.change_lists[0];
            assert_eq!(
                list.author.is_some(),
                sequence.contains('A'),
                "children {sequence}"
            );
            assert_eq!(
                list.changes.len(),
                sequence.matches('G').count(),
                "children {sequence}"
            );
            assert_eq!(
                list.extension_xml.is_some(),
                sequence.contains('X'),
                "children {sequence}"
            );
            continue;
        }
        let expected = if refused_at_change.contains(&sequence.as_str()) {
            LIST_CHANGE_ORDER
        } else if refused_at_extension.contains(&sequence.as_str()) {
            LIST_EXTENSION_ORDER
        } else {
            refused_at_data += 1;
            LIST_DATA_ORDER
        };
        assert_eq!(outcome(&xml), expected, "children {sequence}");
    }
    assert_eq!(refused_at_data, 16);
}

#[test]
fn an_empty_change_list_is_accepted_and_an_unknown_child_is_refused() {
    for lists in ["<pc:docChgLst/>", "<pc:docChgLst></pc:docChgLst>"] {
        let info = parsed(&changes_information(lists)).unwrap();
        assert_eq!(info.change_lists, [List::default()], "{lists}");
    }
    for sequence in ["U", "UA", "AU", "GU", "XU", "AGXU"] {
        assert_eq!(
            outcome(&changes_information(&change_list(sequence))),
            r#"Invalid("unexpected docChgLst child")"#,
            "children {sequence}"
        );
    }
}

#[test]
fn each_change_list_starts_its_own_child_order() {
    for (lists, expected) in [
        (["X", "A"], "ok"),
        (["AGX", "AGX"], "ok"),
        (["G", "AG"], "ok"),
        (["X", "XX"], LIST_EXTENSION_ORDER),
        (["A", "GA"], LIST_DATA_ORDER),
        (["AX", "XG"], LIST_CHANGE_ORDER),
    ] {
        let xml = changes_information(&lists.map(change_list).concat());
        assert_eq!(outcome(&xml), expected, "lists {lists:?}");
    }
}

/// Direct `pc:docChg` children by letter: `D` chgData, `M` docMkLst, `S`
/// sldChg, `T` sldMasterChg, `C` cmAuthorChg, `E` extLst and `U` a child the
/// descriptor does not know.
fn change_children(sequence: &str) -> String {
    sequence
        .chars()
        .map(|letter| match letter {
            'D' => r#"<pc:chgData name="Ada"/>"#,
            'M' => "<pc:docMkLst><pc:docMk/></pc:docMkLst>",
            'S' => r#"<pc:sldChg chg="mod"/>"#,
            'T' => r#"<pc:sldMasterChg chg="mod"/>"#,
            'C' => r#"<pc:cmAuthorChg chg="mod"/>"#,
            'E' => "<p:extLst/>",
            'U' => "<pc:other/>",
            other => panic!("unknown docChg child letter {other}"),
        })
        .collect()
}

fn change(sequence: &str) -> String {
    changes_information(&format!(
        r#"<pc:docChgLst><pc:docChg chg="addSld">{}</pc:docChg></pc:docChgLst>"#,
        change_children(sequence)
    ))
}

#[test]
fn change_children_keep_schema_order_and_only_the_moniker_list_is_single() {
    // Every sequence of up to three children that is not refused for its
    // order. A descriptor without a docMkLst is refused when it closes.
    let accepted = [
        "M", "DM", "MS", "MT", "MC", "ME", "DDM", "DMS", "DMT", "DMC", "DME", "MSS", "MST", "MSC",
        "MSE", "MTT", "MTC", "MTE", "MCC", "MCE", "MEE",
    ];
    let duplicate_moniker = ["MM", "DMM", "MMD", "MMM", "MMS", "MMT", "MMC", "MME"];
    let without_moniker = [
        "D", "S", "T", "C", "E", "DD", "DS", "DT", "DC", "DE", "SS", "ST", "SC", "SE", "TT", "TC",
        "TE", "CC", "CE", "EE", "DDD", "DDS", "DDT", "DDC", "DDE", "DSS", "DST", "DSC", "DSE",
        "DTT", "DTC", "DTE", "DCC", "DCE", "DEE", "SSS", "SST", "SSC", "SSE", "STT", "STC", "STE",
        "SCC", "SCE", "SEE", "TTT", "TTC", "TTE", "TCC", "TCE", "TEE", "CCC", "CCE", "CEE", "EEE",
    ];
    let mut out_of_order = 0;
    for sequence in sequences("DMSTCE", 3) {
        let expected = if accepted.contains(&sequence.as_str()) {
            "ok"
        } else if duplicate_moniker.contains(&sequence.as_str()) {
            CHANGE_DUPLICATE_MONIKER
        } else if without_moniker.contains(&sequence.as_str()) {
            CHANGE_WITHOUT_MONIKER
        } else {
            out_of_order += 1;
            CHANGE_ORDER
        };
        assert_eq!(outcome(&change(&sequence)), expected, "children {sequence}");
    }
    assert_eq!(out_of_order, 174);
}

#[test]
fn change_children_longer_runs_keep_the_same_rules() {
    for (sequence, expected) in [
        ("DDDDM", "ok"),
        ("DMSTCE", "ok"),
        ("DDMSSTTCCEE", "ok"),
        ("MEEEE", "ok"),
        ("DMSTCEM", CHANGE_ORDER),
        ("DMSTCED", CHANGE_ORDER),
        ("DMD", CHANGE_ORDER),
        ("MD", CHANGE_ORDER),
        ("DMSTCEC", CHANGE_ORDER),
        ("DDMM", CHANGE_DUPLICATE_MONIKER),
        ("DDSSTTCCEE", CHANGE_WITHOUT_MONIKER),
        ("SM", CHANGE_ORDER),
    ] {
        assert_eq!(outcome(&change(sequence)), expected, "children {sequence}");
    }
}

#[test]
fn an_empty_change_and_an_unknown_change_child_are_refused() {
    assert_eq!(
        outcome(&changes_information(
            r#"<pc:docChgLst><pc:docChg chg="addSld"/></pc:docChgLst>"#
        )),
        r#"Invalid("docChg cannot be empty")"#
    );
    assert_eq!(outcome(&change("")), CHANGE_WITHOUT_MONIKER);
    // An unknown child is named before its place is judged.
    for sequence in ["U", "UM", "MU", "DU", "MEU", "MMU"] {
        let expected = if sequence == "MMU" {
            CHANGE_DUPLICATE_MONIKER
        } else {
            r#"Invalid("unexpected docChg child")"#
        };
        assert_eq!(outcome(&change(sequence)), expected, "children {sequence}");
    }
}

#[test]
fn a_change_child_is_checked_for_its_place_before_its_attributes() {
    for (children, expected) in [
        (
            r#"<pc:docMkLst><pc:docMk/></pc:docMkLst><pc:sldChg/>"#,
            r#"Invalid("nested change descriptor requires chg")"#,
        ),
        (
            r#"<pc:docMkLst><pc:docMk/></pc:docMkLst><pc:cmAuthorChg chg="mod"/><pc:sldChg/>"#,
            CHANGE_ORDER,
        ),
        (
            r#"<pc:docMkLst><pc:docMk/></pc:docMkLst><pc:docMkLst bad="1"/>"#,
            CHANGE_DUPLICATE_MONIKER,
        ),
        (
            r#"<pc:docMkLst><pc:docMk/></pc:docMkLst><pc:chgData actId="x"/>"#,
            CHANGE_ORDER,
        ),
        (
            r#"<pc:chgData actId="x"/><pc:docMkLst><pc:docMk/></pc:docMkLst>"#,
            r#"Invalid("invalid change action ID")"#,
        ),
    ] {
        let xml = changes_information(&format!(
            r#"<pc:docChgLst><pc:docChg chg="addSld">{children}</pc:docChg></pc:docChgLst>"#
        ));
        assert_eq!(outcome(&xml), expected, "children {children}");
    }
}
