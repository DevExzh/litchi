#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_core::Result;
use litchi_odf_common::{constants, core::PackageWriter, package::raw_identical_members};
use litchi_ods::{
    Spreadsheet,
    document::{Patch, SheetPosition, Snapshot},
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

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
            "Pictures/untouched.bin",
            b"exact opaque worksheet payload",
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

fn fragments(xml: &str) -> Vec<&str> {
    let mut remaining = xml;
    let mut found = Vec::new();
    while let Some(start) = remaining.find("<t:table ") {
        remaining = &remaining[start..];
        let end = remaining.find("</t:table>").unwrap() + "</t:table>".len();
        found.push(&remaining[..end]);
        remaining = &remaining[end..];
    }
    found
}

#[test]
fn first_middle_last_moves_preserve_exact_fragments_and_durable_inverse() -> Result<()> {
    let tables = format!(
        "{}<!--gap-a-->{}<!--gap-b-->{}",
        table("A", "alpha"),
        table("B", "beta"),
        table("C", "gamma")
    );
    let source = package(&content(&tables, ""), &[]);
    let source_xml = Spreadsheet::from_bytes(source.clone())?
        .content_xml()
        .to_string();
    let source_fragments = fragments(&source_xml);
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    edit.move_sheet("A", SheetPosition::Last)?;
    assert_eq!(names(edit.as_bytes())?, ["B", "C", "A"]);
    edit.move_sheet("C", SheetPosition::First)?;
    assert_eq!(names(edit.as_bytes())?, ["C", "B", "A"]);
    edit.move_sheet("A", SheetPosition::Index(1))?;
    assert_eq!(names(edit.as_bytes())?, ["C", "A", "B"]);
    let commit = edit.commit()?;

    let target_xml = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?
        .content_xml()
        .to_string();
    assert_eq!(
        fragments(&target_xml),
        [
            source_fragments[2],
            source_fragments[0],
            source_fragments[1]
        ]
    );
    assert!(target_xml.contains("<!--gap-a-->"));
    assert!(target_xml.contains("<!--gap-b-->"));
    let identical = raw_identical_members(&source, commit.snapshot().as_bytes())
        .ok_or_else(|| litchi_core::Error::InvalidFormat("raw comparison failed".to_string()))?;
    assert!(identical.contains("Pictures/untouched.bin"));

    let wire = commit.patch().to_deterministic_json()?;
    let durable = Patch::from_deterministic_json(&wire, snapshot.limits())?;
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
    Ok(())
}

#[test]
fn exact_noop_bypasses_changed_source_refusals() -> Result<()> {
    let tables = r#"<t:table t:name="A" t:protected="true"/><t:table t:name="B"/>"#;
    let source = package(&content(tables, ""), &[]);
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    edit.move_sheet("A", SheetPosition::First)?;
    let commit = edit.commit()?;
    assert!(!commit.patch().changed());
    assert_eq!(commit.snapshot().as_bytes(), source);
    Ok(())
}

#[test]
fn missing_duplicate_and_out_of_range_refusals_are_atomic() -> Result<()> {
    let source = package(
        &content(&format!("{}{}", table("A", "a"), table("B", "b")), ""),
        &[],
    );
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    assert!(edit.move_sheet("missing", SheetPosition::First).is_err());
    assert!(
        edit.move_sheet("A", SheetPosition::Index(usize::MAX))
            .is_err()
    );
    assert_eq!(edit.as_bytes(), source);
    assert!(!edit.commit()?.patch().changed());

    let duplicate = package(
        &content(&format!("{}{}", table("A", "a"), table("A", "b")), ""),
        &[],
    );
    assert!(Snapshot::from_bytes(duplicate).is_err());
    Ok(())
}

#[test]
fn dependency_mce_unknown_namespace_and_settings_refuse_before_staging() -> Result<()> {
    let cases = [
        content(
            &format!(
                r#"{}{}<t:named-expressions/>"#,
                table("A", "a"),
                table("B", "b")
            ),
            "",
        ),
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
            &format!(r#"{}<vendor:owner/>{}"#, table("A", "a"), table("B", "b")),
            r#"xmlns:vendor="urn:example:vendor""#,
        ),
        content(
            &format!(
                r#"<t:table t:name="A" vendor:owner="opaque"/>{}"#,
                table("B", "b")
            ),
            r#"xmlns:vendor="urn:example:vendor""#,
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
        assert!(edit.move_sheet("A", SheetPosition::Last).is_err());
        assert_eq!(edit.as_bytes(), source);
    }

    let source = package(
        &content(&format!("{}{}", table("A", "a"), table("B", "b")), ""),
        &[("settings.xml", b"<settings/>", "text/xml")],
    );
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    assert!(edit.move_sheet("A", SheetPosition::Last).is_err());
    assert_eq!(edit.as_bytes(), source);

    for path in [
        "Basic/Standard/Module1.xml",
        "Configurations2/views/current.xml",
    ] {
        let source = package(
            &content(&format!("{}{}", table("A", "a"), table("B", "b")), ""),
            &[(path, b"<opaque/>", "text/xml")],
        );
        let snapshot = Snapshot::from_bytes(source.clone())?;
        let mut edit = snapshot.edit();
        assert!(edit.move_sheet("A", SheetPosition::Last).is_err());
        assert_eq!(edit.as_bytes(), source);
    }
    Ok(())
}

#[test]
fn changed_signed_and_protected_sources_refuse_publication() -> Result<()> {
    let ordinary = content(&format!("{}{}", table("A", "a"), table("B", "b")), "");
    let signed = package(
        &ordinary,
        &[(
            "META-INF/documentsignatures.xml",
            br#"<document-signatures/>"#,
            "text/xml",
        )],
    );
    let signed = Snapshot::from_bytes(signed)?;
    let mut edit = signed.edit();
    edit.move_sheet("A", SheetPosition::Last)?;
    assert!(edit.commit().is_err());

    let protected = package(
        &content(
            r#"<t:table t:name="A" t:protected="true"/><t:table t:name="B"/>"#,
            "",
        ),
        &[],
    );
    let protected = Snapshot::from_bytes(protected.clone())?;
    let mut edit = protected.edit();
    assert!(edit.move_sheet("A", SheetPosition::Last).is_err());
    assert_eq!(edit.as_bytes(), protected.as_bytes());
    Ok(())
}

#[test]
fn sheet_count_work_bound_refuses_before_replacement_allocation() -> Result<()> {
    let mut tables = String::new();
    for index in 0..=1_024 {
        tables.push_str(&format!(r#"<t:table t:name="S{index}"/>"#));
    }
    let source = package(&content(&tables, ""), &[]);
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    assert!(edit.move_sheet("S0", SheetPosition::Last).is_err());
    assert_eq!(edit.as_bytes(), source);
    Ok(())
}
