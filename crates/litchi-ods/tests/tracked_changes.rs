use litchi_odf_common::core::{OwnedPackage, PackageWriter, Profile};
use litchi_ods::Spreadsheet;
use litchi_ods::tracked_changes::{
    Acceptance, CellValue, Change, Dimension, Integer, Limits, Patch, Snapshot,
};
use std::sync::Arc;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const DC: &str = "http://purl.org/dc/elements/1.1/";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";

fn info(comment: &str) -> String {
    format!(
        "<office:change-info><dc:creator>Author &amp; Co</dc:creator><dc:date>2026-08-08T00:00:00Z</dc:date><text:p>{comment}</text:p></office:change-info>"
    )
}

fn insertion(id: &str) -> String {
    format!(
        r#"<table:insertion table:id="{id}" table:type="row" table:position="1" table:count="2" table:table="0">{}</table:insertion>"#,
        info("insert &lt;row&gt;")
    )
}

fn deletion(id: &str) -> String {
    format!(
        r#"<table:deletion table:id="{id}" table:type="column" table:position="3" table:table="0" table:multi-deletion-spanned="1">{}<table:dependencies><table:dependency table:id="i"/></table:dependencies><table:deletions><table:change-deletion table:id="i"/></table:deletions><table:cut-offs><table:insertion-cut-off table:id="i" table:position="1"/><table:movement-cut-off table:position="2"/><table:movement-cut-off table:start-position="3" table:end-position="5"/></table:cut-offs></table:deletion>"#,
        info("delete")
    )
}

fn movement(id: &str) -> String {
    format!(
        r#"<table:movement table:id="{id}"><table:source-range-address table:start-table="0" table:start-column="0" table:start-row="0" table:end-table="0" table:end-column="2" table:end-row="3"/><table:target-range-address table:table="0" table:column="4" table:row="5"/>{}</table:movement>"#,
        info("move")
    )
}

fn content_change(id: &str) -> String {
    format!(
        r#"<table:cell-content-change table:id="{id}"><table:cell-address table:table="0" table:column="6" table:row="7"/>{}<table:previous><table:change-track-table-cell office:value-type="currency" office:value="12.5" office:currency="CNY" table:formula="of:=SUM([.A1:.A2])"><text:p>old &lt;&amp;&gt;</text:p></table:change-track-table-cell></table:previous></table:cell-content-change>"#,
        info("content")
    )
}

fn independent_owner(track: Option<bool>) -> String {
    let attribute = match track {
        None => "",
        Some(false) => r#" table:track-changes="false""#,
        Some(true) => r#" table:track-changes="true""#,
    };
    format!(
        "<table:tracked-changes{attribute}>{}{}{}{}</table:tracked-changes>",
        insertion("i"),
        deletion_without_references("d"),
        movement("m"),
        content_change("c")
    )
}

fn deletion_without_references(id: &str) -> String {
    format!(
        r#"<table:deletion table:id="{id}" table:type="column" table:position="3" table:table="0">{}</table:deletion>"#,
        info("delete")
    )
}

fn content(owner: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" xmlns:dc="{DC}" xmlns:xlink="{XLINK}" xmlns:ext="urn:example:tracked" office:version="1.4">
  <office:body><office:spreadsheet>{owner}
    <table:table table:name="Data"><table:table-row><table:table-cell table:formula="of:=CONCAT(&quot;a&amp;b&quot;;[.$A$1])"><text:p><text:a xlink:href="https://example.invalid/?a=1&amp;b=%3Cx%3E">literal &lt;&amp;&gt;</text:a></text:p></table:table-cell></table:table-row></table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#
    )
}

fn package(content_xml: &str, signed: bool) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIMETYPE).unwrap();
    writer
        .add_file("content.xml", content_xml.as_bytes())
        .unwrap();
    writer
        .add_file("Thumbnails/thumbnail.png", b"inert auxiliary bytes")
        .unwrap();
    if signed {
        writer
            .add_file(
                "META-INF/documentsignatures.xml",
                format!(
                    r#"<ds:document-signatures xmlns:ds="{}"/>"#,
                    "urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"
                )
                .as_bytes(),
            )
            .unwrap();
        writer
            .add_file(
                "META-INF/macrosignatures.xml",
                format!(
                    r#"<ds:document-signatures xmlns:ds="{}"/>"#,
                    "urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"
                )
                .as_bytes(),
            )
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn assert_in_order(haystack: &str, needles: &[&str]) {
    let mut cursor = 0;
    for needle in needles {
        let offset = haystack[cursor..]
            .find(needle)
            .unwrap_or_else(|| panic!("missing ordered XML token {needle:?} in {haystack}"));
        cursor += offset + needle.len();
    }
}

fn ids(snapshot: &Snapshot) -> Vec<&str> {
    snapshot
        .changes()
        .unwrap()
        .changes
        .iter()
        .map(|change| change.metadata().id.as_str())
        .collect()
}

#[test]
fn distinguishes_absent_empty_and_all_tracking_attribute_states() {
    let absent = Snapshot::parse(Arc::<str>::from(content(""))).unwrap();
    assert!(absent.changes().is_none());
    assert_eq!(absent.tracking(), None);

    for (owner, expected) in [
        ("<table:tracked-changes/>", None),
        (
            r#"<table:tracked-changes table:track-changes="false"/>"#,
            Some(false),
        ),
        (
            r#"<table:tracked-changes table:track-changes="true"/>"#,
            Some(true),
        ),
    ] {
        let source = content(owner);
        let snapshot = Snapshot::parse(Arc::<str>::from(source.clone())).unwrap();
        assert!(snapshot.changes().is_some());
        assert!(snapshot.changes().unwrap().changes.is_empty());
        assert_eq!(snapshot.tracking(), expected);
        let commit = snapshot.transaction().unwrap().commit().unwrap();
        assert!(!commit.changed());
        assert_eq!(commit.content_xml(), source);
    }

    let empty = Snapshot::parse(content("<table:tracked-changes/>")).unwrap();
    let mut transaction = empty.transaction().unwrap();
    transaction.remove_owner();
    let removed = transaction.commit().unwrap();
    assert!(removed.changed());
    assert!(removed.snapshot().changes().is_none());
}

#[test]
fn reads_all_four_families_and_authors_odf_14_order_with_schema_prefixes() {
    let original = Snapshot::parse(content(&format!(
        r#"<table:tracked-changes table:track-changes="true">{}{}{}{}</table:tracked-changes>"#,
        insertion("i"),
        deletion("d"),
        movement("m"),
        content_change("c")
    )))
    .unwrap();
    let families = original.changes().unwrap();
    assert!(matches!(families.changes[0], Change::Insertion(_)));
    assert!(matches!(families.changes[1], Change::Deletion(_)));
    assert!(matches!(families.changes[2], Change::Movement(_)));
    assert!(matches!(families.changes[3], Change::CellContent(_)));
    let Change::Deletion(deletion) = &families.changes[1] else {
        unreachable!()
    };
    assert!(deletion.cut_offs.iter().any(|cut_off| matches!(
        cut_off,
        litchi_ods::tracked_changes::CutOff::MovementPoint { position }
            if position.as_str() == "2"
    )));

    let semantic = families.clone();
    let canonical_owner = semantic.to_xml_fragment().unwrap();
    let canonical = Snapshot::parse(content(&canonical_owner)).unwrap();
    assert_eq!(
        canonical.changes(),
        Some(&semantic),
        "canonical semantic drift; owner={canonical_owner}"
    );
    let absent = Snapshot::parse(content("")).unwrap();
    let mut transaction = absent.transaction().unwrap();
    transaction.replace_all(Some(semantic)).unwrap();
    transaction.set_tracking(Some(true)).unwrap();
    let authored = transaction.commit().unwrap();
    let xml = authored.content_xml();
    assert!(xml.contains(&format!(r#"xmlns:table="{TABLE}""#)));
    assert!(xml.contains(&format!(r#"xmlns:office="{OFFICE}""#)));
    assert!(xml.contains(&format!(r#"xmlns:text="{TEXT}""#)));
    assert!(xml.contains(&format!(r#"xmlns:dc="{DC}""#)));
    assert_in_order(
        xml,
        &[
            "<table:tracked-changes",
            "<table:insertion",
            "<office:change-info>",
            "<table:deletion",
            "<office:change-info>",
            "<table:dependencies>",
            "<table:deletions>",
            "<table:cut-offs>",
            "<table:movement",
            "<table:source-range-address",
            "<table:target-range-address",
            "<office:change-info>",
            "<table:cell-content-change",
            "<table:cell-address",
            "<office:change-info>",
            "<table:previous>",
        ],
    );
    assert_eq!(
        ids(&Snapshot::parse(authored.source_arc()).unwrap()),
        ["i", "d", "m", "c"]
    );
}

#[test]
fn rejects_wrong_child_order_for_each_family() {
    let change_info = info("order");
    let invalid = [
        format!(
            r#"<table:insertion table:id="i" table:type="row" table:position="0"><table:dependencies><table:dependency table:id="x"/></table:dependencies>{change_info}</table:insertion>"#
        ),
        format!(
            r#"<table:deletion table:id="d" table:type="row" table:position="0">{change_info}<table:cut-offs><table:movement-cut-off table:position="1"/></table:cut-offs><table:deletions><table:change-deletion/></table:deletions></table:deletion>"#
        ),
        format!(
            r#"<table:movement table:id="m"><table:source-range-address table:table="0" table:column="0" table:row="0"/>{change_info}<table:target-range-address table:table="0" table:column="1" table:row="1"/></table:movement>"#
        ),
        format!(
            r#"<table:cell-content-change table:id="c"><table:cell-address table:table="0" table:column="0" table:row="0"/><table:previous><table:change-track-table-cell/></table:previous>{change_info}</table:cell-content-change>"#
        ),
    ];
    for record in invalid {
        let owner = format!("<table:tracked-changes>{record}</table:tracked-changes>");
        assert!(
            Snapshot::parse(content(&owner)).is_err(),
            "accepted {record}"
        );
    }
}

#[test]
fn transaction_crud_reorder_acceptance_reopen_and_rollback_are_atomic() {
    let snapshot = Snapshot::parse(content(&independent_owner(Some(true)))).unwrap();
    let donor = Snapshot::parse(content(&format!(
        "<table:tracked-changes>{}</table:tracked-changes>",
        insertion("x")
    )))
    .unwrap();
    let added = donor.changes().unwrap().changes[0].clone();
    let mut replacement = snapshot.changes().unwrap().changes[2].clone();
    let Change::Movement(movement) = &mut replacement else {
        unreachable!()
    };
    movement.target =
        litchi_ods::tracked_changes::RangeAddress::Cell(litchi_ods::tracked_changes::CellAddress {
            table: 0.into(),
            column: 9.into(),
            row: 9.into(),
        });

    let mut transaction = snapshot.transaction().unwrap();
    transaction.append(added).unwrap();
    transaction.replace("m", replacement).unwrap();
    transaction.remove("i").unwrap();
    transaction.move_to("c", 0).unwrap();
    transaction
        .reorder(&[
            "c".to_string(),
            "d".to_string(),
            "m".to_string(),
            "x".to_string(),
        ])
        .unwrap();
    transaction
        .set_acceptance("d", Some(Acceptance::Accepted))
        .unwrap();
    transaction.set_tracking(Some(false)).unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(ids(commit.snapshot()), ["c", "d", "m", "x"]);
    assert_eq!(commit.snapshot().tracking(), Some(false));
    assert_eq!(
        commit.snapshot().acceptance("d").unwrap(),
        Some(Acceptance::Accepted)
    );
    let reopened = Snapshot::parse(commit.source_arc()).unwrap();
    assert_eq!(ids(&reopened), ["c", "d", "m", "x"]);

    let mut rollback = reopened.transaction().unwrap();
    rollback.remove("x").unwrap();
    rollback.set_tracking(Some(true)).unwrap();
    rollback.rollback().unwrap();
    let rolled_back = rollback.commit().unwrap();
    assert!(!rolled_back.changed());
    assert_eq!(rolled_back.content_xml(), reopened.source_xml());

    let mut reverted = reopened.transaction().unwrap();
    reverted
        .set_acceptance("d", Some(Acceptance::Rejected))
        .unwrap();
    reverted
        .set_acceptance("d", Some(Acceptance::Accepted))
        .unwrap();
    assert!(!reverted.commit().unwrap().changed());
}

#[test]
fn invalid_reorder_preserves_the_complete_draft_and_commits_an_exact_noop() {
    let source = content(&format!(
        r#"<table:tracked-changes table:track-changes="true">{}{}</table:tracked-changes>"#,
        insertion("i"),
        deletion("d")
    ));
    let snapshot = Snapshot::parse(source.clone()).unwrap();
    let mut transaction = snapshot.transaction().unwrap();
    let before = transaction.changes().cloned();
    assert!(!transaction.is_changed());

    assert!(
        transaction
            .reorder(&["d".to_string(), "i".to_string()])
            .is_err()
    );
    assert_eq!(transaction.changes(), before.as_ref());
    assert!(!transaction.is_changed());

    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.content_xml(), source);
    assert_eq!(commit.snapshot().source_xml(), snapshot.source_xml());
}

#[test]
fn direct_insert_succeeds_in_bounds_and_is_atomic_out_of_bounds() {
    let source = content(&format!(
        "<table:tracked-changes>{}</table:tracked-changes>",
        insertion("i")
    ));
    let snapshot = Snapshot::parse(source.clone()).unwrap();
    let deletion = Snapshot::parse(content(&format!(
        r#"<table:tracked-changes>{}{}</table:tracked-changes>"#,
        insertion("i"),
        deletion("d")
    )))
    .unwrap()
    .changes()
    .unwrap()
    .changes[1]
        .clone();

    let mut inserted = snapshot.transaction().unwrap();
    inserted.insert(1, deletion).unwrap();
    let commit = inserted.commit().unwrap();
    assert_eq!(ids(commit.snapshot()), ["i", "d"]);

    let donor = Snapshot::parse(content(&format!(
        "<table:tracked-changes>{}</table:tracked-changes>",
        insertion("x")
    )))
    .unwrap()
    .changes()
    .unwrap()
    .changes[0]
        .clone();
    let mut out_of_bounds = snapshot.transaction().unwrap();
    let before = out_of_bounds.changes().cloned();
    assert!(out_of_bounds.insert(2, donor).is_err());
    assert_eq!(out_of_bounds.changes(), before.as_ref());
    assert!(!out_of_bounds.is_changed());
    let noop = out_of_bounds.commit().unwrap();
    assert!(!noop.changed());
    assert_eq!(noop.content_xml(), source);
}

#[test]
fn patches_are_exact_source_bound_reversible_and_reject_stale_sources() {
    let before = Snapshot::parse(content(&independent_owner(None))).unwrap();
    let mut transaction = before.transaction().unwrap();
    transaction.set_tracking(Some(true)).unwrap();
    let commit = transaction.commit().unwrap();
    let patch = commit.patch().clone();
    assert!(!patch.is_empty());

    let replay = patch.apply(&before).unwrap();
    assert_eq!(replay.content_xml(), commit.content_xml());
    let inverse = patch.inverse();
    let restored = inverse.apply(commit.snapshot()).unwrap();
    assert_eq!(restored.content_xml(), before.source_xml());

    let stale = Snapshot::parse(
        before
            .source_xml()
            .replace("  <office:body>", "<office:body>"),
    )
    .unwrap();
    assert!(patch.apply(&stale).is_err());
    assert!(inverse.apply(&before).is_err());
}

#[test]
fn duplicate_unknown_forward_wrong_family_and_cyclic_references_are_rejected() {
    let duplicate = format!("{}{}", insertion("same"), insertion("same"));
    let unknown = format!(
        r#"<table:insertion table:id="a" table:type="row" table:position="0">{}<table:dependencies><table:dependency table:id="missing"/></table:dependencies></table:insertion>"#,
        info("unknown")
    );
    let forward = format!(
        r#"<table:insertion table:id="a" table:rejecting-change-id="b" table:type="row" table:position="0">{}</table:insertion><table:insertion table:id="b" table:acceptance-state="rejected" table:type="row" table:position="1">{}</table:insertion>"#,
        info("forward rejecting change"),
        info("later rejected change")
    );
    let wrong_family = format!(
        r#"{}<table:cell-content-change table:id="c"><table:cell-address table:table="0" table:column="0" table:row="0"/>{}<table:previous table:id="i"><table:change-track-table-cell/></table:previous></table:cell-content-change>"#,
        insertion("i"),
        info("wrong family")
    );
    let cycle = format!(
        r#"<table:insertion table:id="a" table:type="row" table:position="0">{}<table:dependencies><table:dependency table:id="b"/></table:dependencies></table:insertion><table:insertion table:id="b" table:type="row" table:position="1">{}<table:dependencies><table:dependency table:id="a"/></table:dependencies></table:insertion>"#,
        info("cycle a"),
        info("cycle b")
    );
    for records in [duplicate, unknown, forward, wrong_family, cycle] {
        let owner = format!("<table:tracked-changes>{records}</table:tracked-changes>");
        assert!(
            Snapshot::parse(content(&owner)).is_err(),
            "accepted invalid references: {records}"
        );
    }
}

#[test]
fn dtd_is_rejected_but_pi_foreign_records_and_unrelated_rebinding_are_preserved() {
    let owner = format!(
        "<table:tracked-changes>{}</table:tracked-changes>",
        insertion("i")
    );
    let with_dtd = content(&owner).replacen(
        "<office:document-content",
        "<!DOCTYPE office:document-content [<!ENTITY hostile 'expanded'>]><office:document-content",
        1,
    );
    assert!(Snapshot::parse(with_dtd).is_err());

    let extensions = r#"<?producer keep?><ext:future-record ext:flag='keep' xmlns:table="urn:unrelated:rebound"><table:lookalike table:value="opaque"/></ext:future-record>"#;
    let extended = content(&format!(
        "<table:tracked-changes>{}{extensions}</table:tracked-changes>",
        insertion("i")
    ));
    let snapshot = Snapshot::parse(extended).unwrap();
    assert_eq!(ids(&snapshot), ["i"]);
    let mut transaction = snapshot.transaction().unwrap();
    transaction
        .set_acceptance("i", Some(Acceptance::Accepted))
        .unwrap();
    let committed = transaction.commit().unwrap();
    assert!(committed.content_xml().contains(extensions));
}

#[test]
fn reserved_odf_markup_rejects_while_no_namespace_extensions_remain_opaque() {
    let draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    let style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    let reserved_child = format!(
        r#"<table:tracked-changes>{}<draw:frame xmlns:draw="{draw}"/></table:tracked-changes>"#,
        insertion("i")
    );
    assert!(Snapshot::parse(content(&reserved_child)).is_err());

    let reserved_attribute = format!(
        r#"<table:tracked-changes><table:insertion xmlns:style="{style}" style:name="reserved" table:id="i" table:type="row" table:position="0">{}</table:insertion></table:tracked-changes>"#,
        info("reserved attribute")
    );
    assert!(Snapshot::parse(content(&reserved_attribute)).is_err());

    let unqualified_required = format!(
        r#"<table:tracked-changes><table:insertion id="i" type="row" position="0">{}</table:insertion></table:tracked-changes>"#,
        info("unqualified required attributes")
    );
    assert!(Snapshot::parse(content(&unqualified_required)).is_err());

    let opaque = r#"<future insertion='plain' id='opaque' acceptance-state='accepted'><insertion id='fake' type='row' position='0'><dependency id='i'/></insertion></future>"#;
    let extended_owner = format!(
        "<table:tracked-changes>{}{opaque}</table:tracked-changes>",
        insertion("i")
    );
    let extended = Snapshot::parse(content(&extended_owner)).unwrap();
    assert_eq!(ids(&extended), ["i"]);

    let mut surgical = extended.transaction().unwrap();
    surgical
        .set_acceptance("i", Some(Acceptance::Accepted))
        .unwrap();
    assert!(surgical.commit().unwrap().content_xml().contains(opaque));

    let donor = Snapshot::parse(content(&format!(
        "<table:tracked-changes>{}</table:tracked-changes>",
        insertion("x")
    )))
    .unwrap()
    .changes()
    .unwrap()
    .changes[0]
        .clone();
    let mut replacement = extended.changes().unwrap().clone();
    replacement.changes.push(donor);
    let mut regenerate = extended.transaction().unwrap();
    regenerate.replace_all(Some(replacement)).unwrap();
    assert!(regenerate.commit().is_err());
}

#[test]
fn multi_deletion_span_requires_an_immediate_matching_sequence() {
    fn deletion_record(id: &str, dimension: &str, position: i64, span: &str) -> String {
        format!(
            r#"<table:deletion table:id="{id}" table:type="{dimension}" table:position="{position}" table:table="0"{span}>{}</table:deletion>"#,
            info("multi")
        )
    }

    let first = deletion_record("d1", "row", 3, r#" table:multi-deletion-spanned="2""#);
    let follower = deletion_record("d2", "row", 3, "");
    let valid_owner = format!("<table:tracked-changes>{first}{follower}</table:tracked-changes>");
    let valid = Snapshot::parse(content(&valid_owner)).unwrap();
    let Change::Deletion(deletion) = &valid.changes().unwrap().changes[0] else {
        unreachable!()
    };
    assert_eq!(
        deletion
            .multi_deletion_spanned
            .as_ref()
            .map(|value| value.as_str()),
        Some("2")
    );

    let invalid = [
        deletion_record("table", "table", 3, r#" table:multi-deletion-spanned="2""#),
        first.clone(),
        format!("{first}{}", deletion_record("mismatch", "column", 3, "")),
        format!(
            "{first}{}",
            deletion_record(
                "nested-span",
                "row",
                3,
                r#" table:multi-deletion-spanned="1""#
            )
        ),
    ];
    for records in invalid {
        assert!(
            Snapshot::parse(content(&format!(
                "<table:tracked-changes>{records}</table:tracked-changes>"
            )))
            .is_err()
        );
    }
}

#[test]
fn incompatible_value_attributes_and_invalid_cell_lexicals_are_rejected() {
    fn cell_change(id: &str, cell_attributes: &str) -> String {
        format!(
            r#"<table:cell-content-change table:id="{id}"><table:cell-address table:table="0" table:column="0" table:row="0"/>{}<table:previous><table:change-track-table-cell {cell_attributes}/></table:previous></table:cell-content-change>"#,
            info("invalid cell")
        )
    }

    let invalid = [
        cell_change(
            "boolean-choice",
            r#"office:value-type="boolean" office:value="1""#,
        ),
        cell_change(
            "float-choice",
            r#"office:value-type="float" office:boolean-value="true""#,
        ),
        cell_change(
            "error-choice",
            r#"office:value-type="error" office:value="1""#,
        ),
        cell_change(
            "address",
            r#"office:value-type="string" office:string-value="x" table:cell-address="not-a-cell""#,
        ),
        cell_change(
            "date",
            r#"office:value-type="date" office:date-value="2026-99-99""#,
        ),
        cell_change(
            "duration",
            r#"office:value-type="time" office:time-value="P1X""#,
        ),
    ];
    for record in invalid {
        assert!(
            Snapshot::parse(content(&format!(
                "<table:tracked-changes>{record}</table:tracked-changes>"
            )))
            .is_err()
        );
    }
}

#[test]
fn fixed_depth_limit_accepts_the_boundary_and_rejects_one_more_level() {
    fn deeply_nested(levels: usize) -> String {
        let open = "<ext:n>".repeat(levels);
        let close = "</ext:n>".repeat(levels);
        content(&format!(
            r#"<table:tracked-changes><table:insertion table:id="deep" table:type="row" table:position="0"><office:change-info><dc:creator>Depth</dc:creator><dc:date>2026-08-08T00:00:00Z</dc:date><text:p>{open}x{close}</text:p></office:change-info></table:insertion></table:tracked-changes>"#
        ))
    }

    assert!(Snapshot::parse(deeply_nested(249)).is_ok());
    assert!(Snapshot::parse(deeply_nested(250)).is_err());
}

#[test]
fn rich_neighbor_bytes_are_preserved_while_unknown_records_and_regeneration_are_refused() {
    let rich = format!(
        r#"<table:insertion table:id="rich" table:type="row" table:position="0" ext:flag='keep'><office:change-info><dc:creator>Rich</dc:creator><dc:date>2026-08-08T00:00:00Z</dc:date><text:p>before <ext:inline ext:mode="exact"><text:span text:style-name="Emphasis">rich &amp; exact</text:span></ext:inline> after</text:p></office:change-info></table:insertion>"#
    );
    let opaque = r#"<ext:future-record ext:flag='keep'><ext:payload><![CDATA[<opaque>&bytes]]></ext:payload></ext:future-record>"#;
    let unknown_owner = format!(
        "<table:tracked-changes>{}{opaque}</table:tracked-changes>",
        insertion("plain")
    );
    let unknown = Snapshot::parse(content(&unknown_owner)).unwrap();
    let mut unknown_neighbor = unknown.transaction().unwrap();
    unknown_neighbor
        .set_acceptance("plain", Some(Acceptance::Accepted))
        .unwrap();
    assert!(
        unknown_neighbor
            .commit()
            .unwrap()
            .content_xml()
            .contains(opaque)
    );

    let owner = format!(
        "<table:tracked-changes>{rich}{}</table:tracked-changes>",
        insertion("plain")
    );
    let snapshot = Snapshot::parse(content(&owner)).unwrap();

    let mut neighbor = snapshot.transaction().unwrap();
    neighbor
        .set_acceptance("plain", Some(Acceptance::Accepted))
        .unwrap();
    let edited = neighbor.commit().unwrap();
    assert!(edited.content_xml().contains(&rich));

    let mut rich_edit = snapshot.transaction().unwrap();
    let mut replacement = snapshot.changes().unwrap().changes[0].clone();
    let Change::Insertion(insertion) = &mut replacement else {
        unreachable!()
    };
    insertion.position = 9.into();
    rich_edit.replace("rich", replacement).unwrap();
    assert!(rich_edit.commit().is_err());
}

#[test]
fn package_facade_reopens_preserves_auxiliary_and_keeps_formulas_links_inert() {
    let source_xml = content(&independent_owner(None));
    let bytes = package(&source_xml, false);
    let mut mutable = Spreadsheet::from_bytes(bytes).unwrap();
    mutable
        .update_tracked_changes(|transaction| {
            transaction.set_tracking(Some(true))?;
            transaction.set_acceptance("c", Some(Acceptance::Rejected))?;
            Ok(())
        })
        .unwrap();
    let output = mutable.into_bytes();
    let archive = OwnedPackage::from_bytes(output.clone()).unwrap();
    assert_eq!(
        archive.get_file("Thumbnails/thumbnail.png").unwrap(),
        b"inert auxiliary bytes"
    );
    let reopened = Spreadsheet::from_bytes(output).unwrap();
    assert!(
        reopened
            .tracked_changes_with(Limits::new().with_max_changes(4))
            .is_ok()
    );
    assert!(
        reopened
            .tracked_changes_with(Limits::new().with_max_changes(3))
            .is_err()
    );
    let tracked = reopened.tracked_changes().unwrap();
    assert_eq!(tracked.tracking(), Some(true));
    assert_eq!(tracked.acceptance("c").unwrap(), Some(Acceptance::Rejected));
    assert!(
        reopened
            .content_xml()
            .contains(r#"table:formula="of:=CONCAT(&quot;a&amp;b&quot;;[.$A$1])""#)
    );
    assert!(
        reopened
            .content_xml()
            .contains(r#"xlink:href="https://example.invalid/?a=1&amp;b=%3Cx%3E""#)
    );
    assert!(reopened.content_xml().contains("literal &lt;&amp;&gt;"));
}

#[test]
fn signed_exact_noop_is_byte_exact_but_changed_publication_drops_signatures() {
    let bytes = package(&content(&independent_owner(None)), true);

    let mut noop = Spreadsheet::from_bytes(bytes.clone()).unwrap();
    noop.update_tracked_changes(|_| Ok(())).unwrap();
    assert_eq!(noop.into_bytes(), bytes);

    let mut changed = Spreadsheet::from_bytes(bytes).unwrap();
    changed
        .update_tracked_changes(|transaction| {
            transaction.set_tracking(Some(true))?;
            Ok(())
        })
        .unwrap();
    let archive = OwnedPackage::from_bytes(changed.into_bytes()).unwrap();
    assert!(!archive.has_file("META-INF/documentsignatures.xml").unwrap());
    assert!(!archive.has_file("META-INF/macrosignatures.xml").unwrap());
    assert_eq!(
        archive.get_file("Thumbnails/thumbnail.png").unwrap(),
        b"inert auxiliary bytes"
    );
}

#[test]
fn encrypted_content_is_refused_before_a_tracked_change_transaction_can_begin() {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIMETYPE).unwrap();
    writer
        .set_encryption("tracked-change-password", Profile::compatible())
        .unwrap();
    writer
        .add_file("content.xml", content(&independent_owner(None)).as_bytes())
        .unwrap();
    let encrypted = writer.finish_to_bytes().unwrap();

    assert!(Spreadsheet::from_bytes(encrypted).is_err());
}

#[test]
fn authors_table_dimension_and_all_remaining_historical_scalar_families() {
    let records = format!(
        r##"<table:insertion table:id="table" table:type="table" table:position="0">{}</table:insertion>
<table:cell-content-change table:id="boolean"><table:cell-address table:table="0" table:column="0" table:row="0"/>{}<table:previous><table:change-track-table-cell office:value-type="boolean" office:boolean-value="true"><text:p>true</text:p></table:change-track-table-cell></table:previous></table:cell-content-change>
<table:cell-content-change table:id="percentage"><table:cell-address table:table="0" table:column="1" table:row="0"/>{}<table:previous><table:change-track-table-cell office:value-type="percentage" office:value="0.5"><text:p>50%</text:p></table:change-track-table-cell></table:previous></table:cell-content-change>
<table:cell-content-change table:id="date"><table:cell-address table:table="0" table:column="2" table:row="0"/>{}<table:previous><table:change-track-table-cell office:value-type="date" office:date-value="2026-08-08"><text:p>2026-08-08</text:p></table:change-track-table-cell></table:previous></table:cell-content-change>
<table:cell-content-change table:id="time"><table:cell-address table:table="0" table:column="3" table:row="0"/>{}<table:previous><table:change-track-table-cell office:value-type="time" office:time-value="PT1H2M3S" table:cell-address="Sheet1.$D$1"><text:p>01:02:03</text:p></table:change-track-table-cell></table:previous></table:cell-content-change>
<table:cell-content-change table:id="error"><table:cell-address table:table="0" table:column="4" table:row="0"/>{}<table:previous><table:change-track-table-cell office:value-type="error" office:string-value="#DIV/0!"><text:p>#DIV/0!</text:p></table:change-track-table-cell></table:previous></table:cell-content-change>
<table:cell-content-change table:id="inf"><table:cell-address table:table="0" table:column="5" table:row="0"/>{}<table:previous><table:change-track-table-cell office:value-type="float" office:value="INF"><text:p>INF</text:p></table:change-track-table-cell></table:previous></table:cell-content-change>
<table:cell-content-change table:id="negative-inf"><table:cell-address table:table="0" table:column="6" table:row="0"/>{}<table:previous><table:change-track-table-cell office:value-type="percentage" office:value="-INF"><text:p>-INF%</text:p></table:change-track-table-cell></table:previous></table:cell-content-change>
<table:cell-content-change table:id="nan"><table:cell-address table:table="0" table:column="7" table:row="0"/>{}<table:previous><table:change-track-table-cell office:value-type="currency" office:value="NaN" office:currency="USD"><text:p>NaN USD</text:p></table:change-track-table-cell></table:previous></table:cell-content-change>"##,
        info("table"),
        info("boolean"),
        info("percentage"),
        info("date"),
        info("time"),
        info("error"),
        info("inf"),
        info("negative inf"),
        info("nan")
    );
    let owner = format!("<table:tracked-changes>{records}</table:tracked-changes>");
    let snapshot = Snapshot::parse(content(&owner)).unwrap();
    let changes = snapshot.changes().unwrap();
    let Change::Insertion(table) = &changes.changes[0] else {
        unreachable!()
    };
    assert_eq!(table.dimension, Dimension::Table);
    let values: Vec<&CellValue> = changes.changes[1..]
        .iter()
        .map(|change| {
            let Change::CellContent(change) = change else {
                unreachable!()
            };
            &change.previous.value
        })
        .collect();
    assert!(matches!(values[0], CellValue::Boolean(true)));
    assert!(matches!(values[1], CellValue::Percentage(value) if *value == 0.5));
    assert!(matches!(values[2], CellValue::Date(value) if value == "2026-08-08"));
    assert!(matches!(values[3], CellValue::Time(value) if value == "PT1H2M3S"));
    assert!(matches!(values[4], CellValue::Error(Some(value)) if value == "#DIV/0!"));
    assert!(matches!(values[5], CellValue::Number(value) if value == &f64::INFINITY));
    assert!(matches!(values[6], CellValue::Percentage(value) if value == &f64::NEG_INFINITY));
    assert!(
        matches!(values[7], CellValue::Currency { value, code } if value.is_nan() && code == "USD")
    );

    let semantic = changes.clone();
    let absent = Snapshot::parse(content("")).unwrap();
    let mut transaction = absent.transaction().unwrap();
    transaction.replace_all(Some(semantic.clone())).unwrap();
    let commit = transaction.commit().unwrap();
    let reopened = Snapshot::parse(commit.source_arc()).unwrap();
    assert_eq!(reopened.changes(), Some(&semantic));
}

#[test]
fn self_closing_spreadsheet_creates_owner_and_repeated_end_inserts_keep_indices() {
    let source = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" xmlns:dc="{DC}"><office:body><office:spreadsheet/></office:body></office:document-content>"#
    );
    let snapshot = Snapshot::parse(source).unwrap();
    assert!(snapshot.changes().is_none());
    let donor = Snapshot::parse(content(&format!(
        "<table:tracked-changes>{}</table:tracked-changes>",
        insertion("donor")
    )))
    .unwrap()
    .changes()
    .unwrap()
    .changes[0]
        .clone();
    let mut transaction = snapshot.transaction().unwrap();
    for (index, id) in ["a", "b", "c"].into_iter().enumerate() {
        let mut change = donor.clone();
        let Change::Insertion(insertion) = &mut change else {
            unreachable!()
        };
        insertion.metadata.id = id.to_string();
        transaction.insert(index, change).unwrap();
        assert_eq!(
            transaction
                .changes()
                .unwrap()
                .changes
                .get(index)
                .unwrap()
                .metadata()
                .id,
            id
        );
    }
    let commit = transaction.commit().unwrap();
    assert_eq!(ids(commit.snapshot()), ["a", "b", "c"]);
    assert!(commit.content_xml().contains("<office:spreadsheet>"));
    assert!(commit.content_xml().contains("<table:tracked-changes"));
}

#[test]
fn xsd_whitespace_collapses_and_writers_emit_canonical_atomic_lexicals() {
    let collapsed_info = r#"<office:change-info><dc:creator>Whitespace</dc:creator><dc:date>
 2026-08-08T00:00:00Z 	</dc:date></office:change-info>"#;
    let owner = format!(
        r#"<table:tracked-changes table:track-changes="
 true 	">
<table:insertion table:id="i" table:type=" row " table:position=" +0001 " table:count=" 0002 " table:table=" -0000 ">{collapsed_info}</table:insertion>
<table:deletion table:id="d" table:acceptance-state=" accepted " table:type=" column " table:position=" 0003 ">{collapsed_info}</table:deletion>
<table:cell-content-change table:id="boolean"><table:cell-address table:table=" 0 " table:column=" 0 " table:row=" 0 "/>{collapsed_info}<table:previous><table:change-track-table-cell office:value-type=" boolean " office:boolean-value=" 1 "/></table:previous></table:cell-content-change>
<table:cell-content-change table:id="double"><table:cell-address table:table="0" table:column="1" table:row="0"/>{collapsed_info}<table:previous><table:change-track-table-cell office:value-type=" percentage " office:value=" 001.2500 "/></table:previous></table:cell-content-change>
<table:cell-content-change table:id="duration"><table:cell-address table:table="0" table:column="2" table:row="0"/>{collapsed_info}<table:previous><table:change-track-table-cell office:value-type=" time " office:time-value=" PT1H2M3S "/></table:previous></table:cell-content-change>
</table:tracked-changes>"#
    );
    let parsed = Snapshot::parse(content(&owner)).unwrap();
    assert_eq!(parsed.tracking(), Some(true));
    let semantic = parsed.changes().unwrap().clone();
    let absent = Snapshot::parse(content("")).unwrap();
    let mut transaction = absent.transaction().unwrap();
    transaction.replace_all(Some(semantic)).unwrap();
    transaction.set_tracking(Some(true)).unwrap();
    let xml = transaction.commit().unwrap().into_source();
    for canonical in [
        r#"table:track-changes="true""#,
        r#"table:type="row""#,
        r#"table:position="1""#,
        r#"table:count="2""#,
        r#"table:table="0""#,
        r#"table:acceptance-state="accepted""#,
        r#"office:boolean-value="true""#,
        r#"office:value="1.25""#,
        r#"office:time-value="PT1H2M3S""#,
        "<dc:date>2026-08-08T00:00:00Z</dc:date>",
    ] {
        assert!(
            xml.contains(canonical),
            "missing canonical lexical {canonical}"
        );
    }
}

#[test]
fn arbitrary_precision_integers_round_trip_and_canonicalize_above_machine_sizes() {
    let huge = "9".repeat(200);
    let padded = format!(" +000{huge} ");
    let owner = format!(
        r#"<table:tracked-changes>
<table:insertion table:id="i" table:type="row" table:position="{padded}" table:count=" 000{huge} " table:table="{padded}">{}</table:insertion>
<table:deletion table:id="d" table:type="column" table:position="{padded}" table:table="{padded}">{}<table:cut-offs><table:movement-cut-off table:position="{padded}"/></table:cut-offs></table:deletion>
<table:movement table:id="m"><table:source-range-address table:start-table="{padded}" table:start-column="{padded}" table:start-row="{padded}" table:end-table="{padded}" table:end-column="{padded}" table:end-row="{padded}"/><table:target-range-address table:table="{padded}" table:column="{padded}" table:row="{padded}"/>{}</table:movement>
<table:cell-content-change table:id="c"><table:cell-address table:table="{padded}" table:column="{padded}" table:row="{padded}"/>{}<table:previous><table:change-track-table-cell office:value-type="boolean" office:boolean-value="true" table:number-matrix-columns-spanned=" 000{huge} " table:number-matrix-rows-spanned=" 000{huge} "/></table:previous></table:cell-content-change>
</table:tracked-changes>"#,
        info("big insertion"),
        info("big deletion"),
        info("big movement"),
        info("big cell")
    );
    let parsed = Snapshot::parse(content(&owner)).unwrap();
    let Change::Insertion(insertion) = &parsed.changes().unwrap().changes[0] else {
        unreachable!()
    };
    assert_eq!(insertion.position.as_str(), huge);
    assert_eq!(insertion.position.digit_count(), 200);
    assert_eq!(insertion.count.as_str(), huge);
    assert_eq!(
        Integer::parse(&format!("+000{huge}")).unwrap().as_str(),
        huge
    );

    let semantic = parsed.changes().unwrap().clone();
    let absent = Snapshot::parse(content("")).unwrap();
    let mut transaction = absent.transaction().unwrap();
    transaction.replace_all(Some(semantic.clone())).unwrap();
    let commit = transaction.commit().unwrap();
    assert!(
        commit
            .content_xml()
            .contains(&format!(r#"table:position="{huge}""#))
    );
    assert!(
        commit
            .content_xml()
            .contains(&format!(r#"table:count="{huge}""#))
    );
    assert!(
        commit
            .content_xml()
            .contains(&format!(r#"table:number-matrix-columns-spanned="{huge}""#))
    );
    assert!(!commit.content_xml().contains("+000"));
    assert_eq!(
        Snapshot::parse(commit.source_arc()).unwrap().changes(),
        Some(&semantic)
    );
}

#[test]
fn alternate_odf_prefixes_ignore_unused_vendor_canonical_bindings_and_preserve_spoofs() {
    let opaque = r#"<v:future v:insertion='attr' insertion='plain'><insertion id='spoof'/><v:deletion v:id='opaque'/></v:future>"#;
    let source = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}" xmlns:x="{TEXT}" xmlns:d="{DC}" xmlns:v="urn:vendor" xmlns:office="urn:vendor:office" xmlns:table="urn:vendor:table" xmlns:text="urn:vendor:text" xmlns:dc="urn:vendor:dc"><o:body><o:spreadsheet><t:tracked-changes><t:insertion t:id="i" t:type="row" t:position="0"><o:change-info><d:creator>Alias</d:creator><d:date>2026-08-08T00:00:00Z</d:date><x:p>alias</x:p></o:change-info></t:insertion>{opaque}</t:tracked-changes></o:spreadsheet></o:body></o:document-content>"#
    );
    let snapshot = Snapshot::parse(source).unwrap();
    assert_eq!(ids(&snapshot), ["i"]);
    let mut transaction = snapshot.transaction().unwrap();
    transaction
        .set_acceptance("i", Some(Acceptance::Accepted))
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.content_xml().contains(opaque));
    assert!(
        commit
            .content_xml()
            .contains(r#"xmlns:office="urn:vendor:office""#)
    );
    assert!(
        commit
            .content_xml()
            .contains(r#"xmlns:table="urn:vendor:table""#)
    );
    assert_eq!(
        commit.snapshot().acceptance("i").unwrap(),
        Some(Acceptance::Accepted)
    );
}

#[test]
fn invalid_cross_record_operations_fail_immediately_without_draft_mutation() {
    let one = Snapshot::parse(content(&format!(
        "<table:tracked-changes>{}</table:tracked-changes>",
        insertion("i")
    )))
    .unwrap();
    let mut invalid = one.changes().unwrap().changes[0].clone();
    let Change::Insertion(invalid_insertion) = &mut invalid else {
        unreachable!()
    };
    invalid_insertion.metadata.id = "x".to_string();
    invalid_insertion
        .metadata
        .dependencies
        .push("missing".to_string());
    let mut transaction = one.transaction().unwrap();
    let before = transaction.changes().cloned();
    assert!(transaction.insert(1, invalid).is_err());
    assert_eq!(transaction.changes(), before.as_ref());
    assert!(!transaction.is_changed());
    assert!(!transaction.commit().unwrap().changed());

    let cutoff = Snapshot::parse(content(&format!(
        "<table:tracked-changes>{}{}</table:tracked-changes>",
        insertion("i"),
        deletion("d")
    )))
    .unwrap();
    let movement_source = Snapshot::parse(content(&format!(
        "<table:tracked-changes>{}</table:tracked-changes>",
        movement("m")
    )))
    .unwrap();
    let mut movement_replacement = movement_source.changes().unwrap().changes[0].clone();
    let Change::Movement(movement) = &mut movement_replacement else {
        unreachable!()
    };
    movement.metadata.id = "i".to_string();
    let mut transaction = cutoff.transaction().unwrap();
    let before = transaction.changes().cloned();
    assert!(transaction.replace("i", movement_replacement).is_err());
    assert_eq!(transaction.changes(), before.as_ref());
    assert!(!transaction.is_changed());
    assert!(!transaction.commit().unwrap().changed());

    let rejecting = format!(
        r#"<table:tracked-changes><table:insertion table:id="r" table:acceptance-state="rejected" table:type="row" table:position="0">{}</table:insertion><table:insertion table:id="a" table:rejecting-change-id="r" table:type="row" table:position="1">{}</table:insertion></table:tracked-changes>"#,
        info("rejected"),
        info("rejecting")
    );
    let rejecting = Snapshot::parse(content(&rejecting)).unwrap();
    let mut transaction = rejecting.transaction().unwrap();
    let before = transaction.changes().cloned();
    assert!(
        transaction
            .set_acceptance("r", Some(Acceptance::Pending))
            .is_err()
    );
    assert_eq!(transaction.changes(), before.as_ref());
    assert!(!transaction.is_changed());
    assert!(!transaction.commit().unwrap().changed());

    let forward = format!(
        r#"<table:tracked-changes><table:insertion table:id="a" table:type="row" table:position="0">{}<table:dependencies><table:dependency table:id="b"/></table:dependencies></table:insertion>{}</table:tracked-changes>"#,
        info("forward"),
        insertion("b")
    );
    assert!(Snapshot::parse(content(&forward)).is_ok());
}

#[test]
fn asymmetric_input_output_limits_keep_commits_and_patch_results_reeditable() {
    let source = content("<table:tracked-changes/>");
    let limits = Limits::new()
        .with_max_input_bytes(source.len())
        .with_max_output_bytes(source.len() + 16_384);
    let initial = Snapshot::parse_with_limits(source.clone(), limits).unwrap();
    let donor = Snapshot::parse(content(&format!(
        "<table:tracked-changes>{}</table:tracked-changes>",
        insertion("i")
    )))
    .unwrap()
    .changes()
    .unwrap()
    .changes[0]
        .clone();

    let mut transaction = initial.transaction().unwrap();
    transaction.insert(0, donor.clone()).unwrap();
    let committed = transaction.commit().unwrap();
    assert!(committed.content_xml().len() > limits.max_input_bytes());
    let patch = committed.patch().clone();

    let mut second = committed.snapshot().transaction().unwrap();
    second
        .set_acceptance("i", Some(Acceptance::Accepted))
        .unwrap();
    assert!(second.commit().unwrap().changed());

    let replayed = patch.apply(&initial).unwrap();
    let mut replay_edit = replayed.snapshot().transaction().unwrap();
    replay_edit
        .set_acceptance("i", Some(Acceptance::Rejected))
        .unwrap();
    assert!(replay_edit.commit().unwrap().changed());

    let restored = patch.inverse().apply(committed.snapshot()).unwrap();
    let mut inverse_edit = restored.snapshot().transaction().unwrap();
    inverse_edit.insert(0, donor).unwrap();
    assert!(inverse_edit.commit().unwrap().changed());
}

#[test]
fn bulk_acceptance_and_replace_keep_many_inbound_rejecting_edges_consistent() {
    let mut records = String::new();
    for index in 0..16 {
        records.push_str(&format!(
            r#"<table:insertion table:id="r{index}" table:acceptance-state="rejected" table:type="row" table:position="{index}">{}</table:insertion><table:insertion table:id="a{index}" table:rejecting-change-id="r{index}" table:type="row" table:position="{}">{}</table:insertion>"#,
            info("rejected target"),
            index + 32,
            info("rejecting record")
        ));
    }
    let owner = format!("<table:tracked-changes>{records}</table:tracked-changes>");
    let snapshot = Snapshot::parse(content(&owner)).unwrap();
    let mut transaction = snapshot.transaction().unwrap();
    for index in 0..16 {
        let target_id = format!("r{index}");
        let mut replacement = snapshot
            .changes()
            .unwrap()
            .changes
            .iter()
            .find(|change| change.metadata().id == target_id)
            .unwrap()
            .clone();
        let Change::Insertion(insertion) = &mut replacement else {
            unreachable!()
        };
        insertion.position = (index + 100).into();
        transaction.replace(&target_id, replacement).unwrap();
        transaction
            .set_acceptance(&format!("a{index}"), Some(Acceptance::Accepted))
            .unwrap();
    }
    let commit = transaction.commit().unwrap();
    for index in 0..16 {
        assert_eq!(
            commit.snapshot().acceptance(&format!("a{index}")).unwrap(),
            Some(Acceptance::Accepted)
        );
        let target_id = format!("r{index}");
        let target = commit
            .snapshot()
            .changes()
            .unwrap()
            .changes
            .iter()
            .find(|change| change.metadata().id == target_id)
            .unwrap();
        let Change::Insertion(insertion) = target else {
            unreachable!()
        };
        assert_eq!(insertion.position.as_str(), (index + 100).to_string());
    }
}

#[test]
fn multi_deletion_run_breaking_mutations_fail_immediately_and_atomically() {
    let multi = format!(
        r#"<table:tracked-changes><table:deletion table:id="d1" table:type="row" table:position="3" table:table="0" table:multi-deletion-spanned="2">{}</table:deletion><table:deletion table:id="d2" table:type="row" table:position="3" table:table="0">{}</table:deletion></table:tracked-changes>"#,
        info("multi first"),
        info("multi follower")
    );
    let snapshot = Snapshot::parse(content(&multi)).unwrap();
    let donor = Snapshot::parse(content(&format!(
        "<table:tracked-changes>{}</table:tracked-changes>",
        insertion("x")
    )))
    .unwrap()
    .changes()
    .unwrap()
    .changes[0]
        .clone();

    let mut insert = snapshot.transaction().unwrap();
    let before = insert.changes().cloned();
    assert!(insert.insert(1, donor).is_err());
    assert_eq!(insert.changes(), before.as_ref());
    assert!(!insert.is_changed());
    assert!(!insert.commit().unwrap().changed());

    let mut remove = snapshot.transaction().unwrap();
    let before = remove.changes().cloned();
    assert!(remove.remove("d2").is_err());
    assert_eq!(remove.changes(), before.as_ref());
    assert!(!remove.is_changed());
    assert!(!remove.commit().unwrap().changed());

    let mut mismatch = snapshot.changes().unwrap().changes[1].clone();
    let Change::Deletion(deletion) = &mut mismatch else {
        unreachable!()
    };
    deletion.position = 4.into();
    let mut replace = snapshot.transaction().unwrap();
    let before = replace.changes().cloned();
    assert!(replace.replace("d2", mismatch).is_err());
    assert_eq!(replace.changes(), before.as_ref());
    assert!(!replace.is_changed());
    assert!(!replace.commit().unwrap().changed());
}

#[test]
fn package_patch_application_is_source_checked_and_inverse_restores_exact_content() {
    let bytes = package(&content(&independent_owner(None)), false);
    let before = Spreadsheet::from_bytes(bytes.clone())
        .unwrap()
        .tracked_changes()
        .unwrap();
    let mut transaction = before.transaction().unwrap();
    transaction.set_tracking(Some(true)).unwrap();
    let patch: Patch = transaction.commit().unwrap().into_patch();

    let mut mutable = Spreadsheet::from_bytes(bytes).unwrap();
    mutable.apply_tracked_changes_patch(&patch).unwrap();
    assert_eq!(mutable.tracked_changes().unwrap().tracking(), Some(true));
    mutable
        .apply_tracked_changes_patch(&patch.inverse())
        .unwrap();
    assert_eq!(
        mutable.tracked_changes().unwrap().source_xml(),
        before.source_xml()
    );
    assert!(
        mutable
            .apply_tracked_changes_patch(&patch.inverse())
            .is_err()
    );
}
