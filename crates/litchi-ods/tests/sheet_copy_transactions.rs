#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_core::{Result, xml::escape_xml};
use litchi_odf_common::{
    constants,
    core::{PackageWriter, Profile},
    package::raw_identical_members,
};
use litchi_ods::{
    Spreadsheet,
    document::{Limits, Patch, SheetPosition, Snapshot},
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

fn content(tables: &str, extra_namespaces: &str) -> String {
    let extra_namespaces = if extra_namespaces.is_empty() {
        String::new()
    } else {
        format!(" {extra_namespaces}")
    };
    format!(
        r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE}" xmlns:t="{TABLE}" xmlns:x="{TEXT}"{extra_namespaces} office:version="1.3"><office:body><office:spreadsheet>{tables}</office:spreadsheet></office:body></office:document-content>"#
    )
}

fn table(name: &str, token: &str) -> String {
    format!(
        r#"<t:table t:name="{name}" t:style-name="style-{token}"><t:table-row><t:table-cell office:value-type="string"><x:p>{token}</x:p></t:table-cell></t:table-row></t:table>"#
    )
}

fn package(content_xml: &str, additions: &[(&str, &[u8], &str)]) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype(constants::ODF_SPREADSHEET)
        .expect("fixed ODS mimetype");
    writer
        .add_file("content.xml", content_xml.as_bytes())
        .expect("fixed content.xml");
    writer
        .add_file_with_media_type(
            "Assets/untouched.bin",
            b"exact opaque package payload",
            "application/octet-stream",
        )
        .expect("fixed opaque member");
    for (path, bytes, media_type) in additions {
        writer
            .add_file_with_media_type(path, bytes, media_type)
            .expect("fixed test member");
    }
    writer.finish_to_bytes().expect("fixed ODS package")
}

fn names(bytes: &[u8]) -> Result<Vec<String>> {
    Ok(Spreadsheet::from_bytes(bytes.to_vec())?
        .sheets()
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect())
}

#[test]
fn first_middle_last_copies_keep_exact_table_bytes_and_durable_lineage() -> Result<()> {
    let tables = format!(
        "{}<!--gap-a-->{}<!--gap-b-->{}",
        table("A", "alpha"),
        table("B", "beta"),
        table("C", "gamma")
    );
    let source = package(&content(&tables, ""), &[]);
    let cases = [
        (SheetPosition::First, "first", vec!["first", "A", "B", "C"]),
        (
            SheetPosition::Index(1),
            "middle & <copy> \"'",
            vec!["A", "middle & <copy> \"'", "B", "C"],
        ),
        (SheetPosition::Last, "last", vec!["A", "B", "C", "last"]),
    ];

    for (position, destination, expected_names) in cases {
        let snapshot = Snapshot::from_bytes(source.clone())?;
        let mut edit = snapshot.edit();
        edit.copy_sheet("A", destination, position)?;
        assert_eq!(names(edit.as_bytes())?, expected_names);
        let commit = edit.commit()?;
        let target_xml = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?
            .content_xml()
            .to_string();
        assert!(target_xml.contains(&table("A", "alpha")));
        assert!(target_xml.contains(&table(&escape_xml(destination), "alpha")));
        assert_eq!(target_xml.matches("<x:p>alpha</x:p>").count(), 2);
        assert!(target_xml.contains("<!--gap-a-->"));
        assert!(target_xml.contains("<!--gap-b-->"));

        let identical =
            raw_identical_members(&source, commit.snapshot().as_bytes()).ok_or_else(|| {
                litchi_core::Error::InvalidFormat("raw comparison failed".to_string())
            })?;
        assert!(identical.contains("Assets/untouched.bin"));
        assert!(identical.contains("META-INF/manifest.xml"));

        let wire = commit.patch().to_deterministic_json()?;
        let durable = Patch::from_deterministic_json(&wire, snapshot.limits())?;
        assert_eq!(durable.operations()[0].op, "ods.sheet.copy");
        assert_eq!(
            durable.apply(&snapshot)?.snapshot().as_bytes(),
            commit.snapshot().as_bytes()
        );
        let inverse = Patch::from_deterministic_json(
            &durable.inverse().to_deterministic_json()?,
            snapshot.limits(),
        )?;
        assert_eq!(
            inverse.apply(commit.snapshot())?.snapshot().as_bytes(),
            source
        );
        let stale = Snapshot::from_bytes(package(&content(&table("Other", "x"), ""), &[]))?;
        assert!(durable.apply(&stale).is_err());
    }
    Ok(())
}

#[test]
fn same_document_hidden_sheet_style_is_safe_and_source_exact() -> Result<()> {
    let xml = format!(
        r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE}" xmlns:t="{TABLE}" xmlns:x="{TEXT}" xmlns:s="{STYLE}" office:version="1.3"><office:automatic-styles><s:style s:name="hidden" s:family="table"><s:table-properties t:display="false"/></s:style></office:automatic-styles><office:body><office:spreadsheet><t:table t:name="Hidden" t:style-name="hidden"><t:table-row><t:table-cell><x:p>opaque spacing</x:p></t:table-cell></t:table-row></t:table><t:table t:name="Visible"/></office:spreadsheet></office:body></office:document-content>"#
    );
    let source = package(
        &xml,
        &[("styles.xml", b"<styles>untouched</styles>", "text/xml")],
    );
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    edit.copy_sheet("Hidden", "Hidden copy", SheetPosition::Last)?;
    let commit = edit.commit()?;
    let target = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
    assert_eq!(
        target
            .sheets()
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        ["Hidden", "Visible", "Hidden copy"]
    );
    assert_eq!(
        target
            .content_xml()
            .matches("t:style-name=\"hidden\"")
            .count(),
        2
    );
    assert_eq!(target.content_xml().matches("opaque spacing").count(), 2);
    let identical = raw_identical_members(&source, commit.snapshot().as_bytes())
        .ok_or_else(|| litchi_core::Error::InvalidFormat("raw comparison failed".to_string()))?;
    assert!(identical.contains("styles.xml"));
    Ok(())
}

#[test]
fn inherited_prefix_single_quotes_and_attribute_order_remain_exact() -> Result<()> {
    let source_table = r#"<t:table t:style-name='style-alpha' t:name='A &amp; old'><t:table-row><t:table-cell office:value-type='string'><x:p>  exact body  </x:p></t:table-cell></t:table-row></t:table>"#;
    let source = package(
        &content(&format!("{source_table}{}", table("B", "b")), ""),
        &[],
    );
    let snapshot = Snapshot::from_bytes(source)?;
    let mut edit = snapshot.edit();
    edit.copy_sheet("A & old", "new <name> & \"quote\"", SheetPosition::Last)?;
    let commit = edit.commit()?;
    let xml = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?
        .content_xml()
        .to_string();
    let expected_copy =
        source_table.replacen("A &amp; old", &escape_xml("new <name> & \"quote\""), 1);
    assert!(xml.contains(source_table));
    assert!(xml.contains(&expected_copy));
    Ok(())
}

#[test]
fn attribute_whitespace_names_reopen_exactly_and_collisions_are_atomic() -> Result<()> {
    let source = package(
        &content(&format!("{}{}", table("A", "a"), table("B", "b")), ""),
        &[],
    );
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    let destination = "tab\tline\nreturn\r";
    edit.copy_sheet("A", destination, SheetPosition::Last)?;
    let commit = edit.commit()?;
    assert_eq!(names(commit.snapshot().as_bytes())?[2], destination);
    let xml = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?
        .content_xml()
        .to_string();
    assert!(xml.contains(r#"t:name="tab&#x9;line&#xA;return&#xD;""#));
    let wire = commit.patch().to_deterministic_json()?;
    let durable = Patch::from_deterministic_json(&wire, snapshot.limits())?;
    assert_eq!(
        durable.apply(&snapshot)?.snapshot().as_bytes(),
        commit.snapshot().as_bytes()
    );

    let collision = package(
        &content(
            &format!("{}{}", table("A", "a"), table("has&#x9;tab", "b")),
            "",
        ),
        &[],
    );
    let snapshot = Snapshot::from_bytes(collision.clone())?;
    let mut edit = snapshot.edit();
    assert!(
        edit.copy_sheet("A", "has\ttab", SheetPosition::Last)
            .is_err()
    );
    assert_eq!(edit.as_bytes(), collision);
    Ok(())
}

#[test]
fn selectors_names_positions_and_output_bounds_refuse_atomically() -> Result<()> {
    let source = package(
        &content(&format!("{}{}", table("A", "a"), table("B", "b")), ""),
        &[],
    );
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    for result in [
        edit.copy_sheet("missing", "C", SheetPosition::First),
        edit.copy_sheet("A", "B", SheetPosition::First),
        edit.copy_sheet("A", "", SheetPosition::First),
        edit.copy_sheet("A", "C", SheetPosition::Index(3)),
        edit.copy_sheet("A", &"x".repeat(1_025), SheetPosition::Last),
    ] {
        assert!(result.is_err());
        assert_eq!(edit.as_bytes(), source);
    }
    assert!(!edit.commit()?.patch().changed());

    let defaults = Limits::default();
    let exact_source_limit = Limits::new(
        source.len(),
        defaults.max_resources(),
        defaults.max_resource_bytes(),
        defaults.patch(),
        defaults.composition(),
        defaults.history(),
    );
    let snapshot = Snapshot::from_bytes_with(source.clone(), exact_source_limit)?;
    let mut edit = snapshot.edit();
    assert!(edit.copy_sheet("A", "C", SheetPosition::Last).is_err());
    assert_eq!(edit.as_bytes(), source);
    Ok(())
}

#[test]
fn long_source_name_keeps_the_durable_operation_target_bounded() -> Result<()> {
    let source_name = "source-".to_string() + &"x".repeat(5_000);
    let source = package(
        &content(
            &format!("{}{}", table(&source_name, "a"), table("B", "b")),
            "",
        ),
        &[],
    );
    let snapshot = Snapshot::from_bytes(source)?;
    let mut edit = snapshot.edit();
    edit.copy_sheet(&source_name, "copy", SheetPosition::Last)?;
    let commit = edit.commit()?;
    assert_eq!(names(commit.snapshot().as_bytes())?[2], "copy");
    let operation = &commit.patch().operations()[0];
    assert!(operation.target.len() < 256);
    let wire = commit.patch().to_deterministic_json()?;
    let durable = Patch::from_deterministic_json(&wire, snapshot.limits())?;
    assert_eq!(
        durable.apply(&snapshot)?.snapshot().as_bytes(),
        commit.snapshot().as_bytes()
    );
    Ok(())
}

#[test]
fn formula_range_validation_active_external_and_opaque_owners_refuse() -> Result<()> {
    let cases = [
        content(
            &format!(
                r#"<t:table t:name="A"><t:table-row><t:table-cell t:formula="of:=1+1" office:value-type="float" office:value="2"/></t:table-row></t:table>{}"#,
                table("B", "b")
            ),
            "",
        ),
        content(
            &format!(
                r#"<t:table t:name="A" t:print-ranges="A.A1:A.A2"/>{}"#,
                table("B", "b")
            ),
            "",
        ),
        content(
            &format!(
                r#"<t:table t:name="A"><t:table-row><t:table-cell t:content-validation-name="v"/></t:table-row></t:table>{}"#,
                table("B", "b")
            ),
            "",
        ),
        content(
            &format!(
                r#"<t:table t:name="A"><t:table-row><t:table-cell><x:p><x:a l:href="https://example.test/never-fetch">external</x:a></x:p></t:table-cell></t:table-row></t:table>{}"#,
                table("B", "b")
            ),
            r#"xmlns:l="http://www.w3.org/1999/xlink""#,
        ),
        content(
            &format!(
                r#"<t:table t:name="A" xml:id="duplicate-me"/>{}"#,
                table("B", "b")
            ),
            r#"xmlns:xml="http://www.w3.org/XML/1998/namespace""#,
        ),
        content(
            &format!(r#"{}<vendor:owner/>{}"#, table("A", "a"), table("B", "b")),
            r#"xmlns:vendor="urn:example:vendor""#,
        ),
        content(
            &format!(
                r#"<t:table t:name="A" vendor-owner="opaque"/>{}"#,
                table("B", "b")
            ),
            "",
        ),
        content(
            &format!(
                r#"<t:table t:name="A" t:vendor-extension="opaque"/>{}"#,
                table("B", "b")
            ),
            "",
        ),
        content(
            &format!(
                r#"<t:table t:name="A"><t:table-row t:name="duplicate-owner"/></t:table>{}"#,
                table("B", "b")
            ),
            "",
        ),
        content(
            &format!(
                r#"<t:table t:name="A"><t:shapes><d:frame d:name="duplicate-frame"/></t:shapes></t:table>{}"#,
                table("B", "b")
            ),
            r#"xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0""#,
        ),
        content(
            &format!(
                r#"<t:table t:name="A"><office:forms><f:form f:name="duplicate-form"/></office:forms></t:table>{}"#,
                table("B", "b")
            ),
            r#"xmlns:f="urn:oasis:names:tc:opendocument:xmlns:form:1.0""#,
        ),
        content(
            &format!(
                r#"<t:table t:name="A"><office:event-listeners><s:event-listener s:event-name="dom:click"/></office:event-listeners></t:table>{}"#,
                table("B", "b")
            ),
            r#"xmlns:s="urn:oasis:names:tc:opendocument:xmlns:script:1.0""#,
        ),
        content(
            &format!(
                r#"<t:table t:name="A"><t:table-source t:table-name="remote"/></t:table>{}"#,
                table("B", "b")
            ),
            "",
        ),
        content(
            &format!(r#"{}<?vendor active?>{}"#, table("A", "a"), table("B", "b")),
            "",
        ),
        content(
            &format!(
                r#"{}<office:event-listeners/>{}"#,
                table("A", "a"),
                table("B", "b")
            ),
            "",
        ),
        content(
            &format!(
                r#"<t:table t:name="A" t:protected="true"/>{}"#,
                table("B", "b")
            ),
            "",
        ),
        format!(
            r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE}" xmlns:t="{TABLE}" xmlns:x="{TEXT}" office:version="1.3"><office:body><office:spreadsheet t:structure-protected="true">{}{}</office:spreadsheet></office:body></office:document-content>"#,
            table("A", "a"),
            table("B", "b")
        ),
        content(
            &format!(
                r#"<t:table t:name="A" t:protection-key="AA==" t:protection-key-digest-algorithm="urn:example:digest"/>{}"#,
                table("B", "b")
            ),
            "",
        ),
        content(
            &format!(
                r#"{}<mc:AlternateContent/>{}"#,
                table("A", "a"),
                table("B", "b")
            ),
            r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#,
        ),
    ];
    for content_xml in cases {
        let source = package(&content_xml, &[]);
        let snapshot = Snapshot::from_bytes(source.clone())?;
        let mut edit = snapshot.edit();
        assert!(edit.copy_sheet("A", "A copy", SheetPosition::Last).is_err());
        assert_eq!(edit.as_bytes(), source);
    }
    Ok(())
}

#[test]
fn dependent_package_members_signatures_and_encryption_refuse() -> Result<()> {
    let ordinary = content(&format!("{}{}", table("A", "a"), table("B", "b")), "");
    for path in [
        "settings.xml",
        "Pictures/image.bin",
        "Object 1/content.xml",
        "Basic/Standard/Module1.xml",
        "META-INF/documentsignatures.xml",
    ] {
        let payload: &[u8] = if path.ends_with(".xml") {
            b"<opaque/>"
        } else {
            b"opaque"
        };
        let source = package(&ordinary, &[(path, payload, "application/octet-stream")]);
        let snapshot = Snapshot::from_bytes(source.clone())?;
        let mut edit = snapshot.edit();
        assert!(edit.copy_sheet("A", "A copy", SheetPosition::Last).is_err());
        assert_eq!(edit.as_bytes(), source);
    }

    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_SPREADSHEET)?;
    writer.set_encryption("secret", Profile::compatible())?;
    writer.add_file("content.xml", ordinary.as_bytes())?;
    let encrypted = writer.finish_to_bytes()?;
    assert!(Snapshot::from_bytes(encrypted).is_err());
    Ok(())
}

#[test]
fn sheet_count_and_patch_conflicts_are_bounded() -> Result<()> {
    let mut tables = String::new();
    for index in 0..1_024 {
        tables.push_str(&format!(r#"<t:table t:name="S{index}"/>"#));
    }
    let source = package(&content(&tables, ""), &[]);
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    assert!(edit.copy_sheet("S0", "copy", SheetPosition::Last).is_err());
    assert_eq!(edit.as_bytes(), source);

    let source = package(
        &content(&format!("{}{}", table("A", "a"), table("B", "b")), ""),
        &[],
    );
    let snapshot = Snapshot::from_bytes(source)?;
    let mut left = snapshot.edit();
    left.copy_sheet("A", "A copy", SheetPosition::Last)?;
    let left = left.commit()?;
    let mut right = snapshot.edit();
    right.copy_sheet("B", "B copy", SheetPosition::Last)?;
    let right = right.commit()?;
    assert!(left.patch().join(right.patch()).is_err());
    Ok(())
}
