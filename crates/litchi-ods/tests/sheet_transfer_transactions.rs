#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on fixed in-memory fixtures"
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

fn scalar_table(name: &str, token: &str) -> String {
    format!(
        r#"<t:table t:name="{name}"><t:table-row t:number-rows-repeated="2"><t:table-cell office:value-type="float" office:value="1.50"><x:p>{token}</x:p></t:table-cell><t:table-cell office:value-type="boolean" office:boolean-value="true"/></t:table-row><t:table-row><t:table-cell office:value-type="date" office:date-value="2026-08-14"><x:p>date</x:p></t:table-cell><t:table-cell office:value-type="time" office:time-value="PT1H"><x:p>time</x:p></t:table-cell></t:table-row></t:table>"#
    )
}

fn plain_table(name: &str, token: &str) -> String {
    format!(
        r#"<t:table t:name="{name}"><t:table-row><t:table-cell office:value-type="string"><x:p>{token}</x:p></t:table-cell></t:table-row></t:table>"#
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
            b"opaque destination payload",
            "application/octet-stream",
        )
        .expect("fixed opaque member");
    for (path, bytes, media_type) in additions {
        writer
            .add_file_with_media_type(path, bytes, media_type)
            .expect("fixed package member");
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
fn scalar_sheet_transfer_preserves_fragment_members_and_durable_inverse() -> Result<()> {
    let source_fragment = scalar_table("Source", "1.50");
    let source_bytes = package(
        &content(&source_fragment, ""),
        &[(
            "Source/opaque.bin",
            b"source opaque",
            "application/octet-stream",
        )],
    );
    let destination_bytes = package(
        &content(
            &format!(
                "{}<!--destination-gap-->{}",
                plain_table("A", "alpha"),
                plain_table("B", "beta")
            ),
            "",
        ),
        &[(
            "Destination/opaque.bin",
            b"destination opaque",
            "application/octet-stream",
        )],
    );
    let source = Snapshot::from_bytes(source_bytes.clone())?;
    for (position, expected) in [
        (SheetPosition::First, vec!["Source copy", "A", "B"]),
        (SheetPosition::Index(1), vec!["A", "Source copy", "B"]),
        (SheetPosition::Last, vec!["A", "B", "Source copy"]),
    ] {
        let destination = Snapshot::from_bytes(destination_bytes.clone())?;
        let mut edit = destination.edit();
        edit.transfer_plain_scalar_sheet_from(&source, "Source", "Source copy", position)?;
        assert_eq!(names(edit.as_bytes())?, expected);
        assert_eq!(source.as_bytes(), source_bytes.as_slice());

        let copied = scalar_table("Source copy", "1.50");
        let target_xml = Spreadsheet::from_bytes(edit.as_bytes().to_vec())?
            .content_xml()
            .to_string();
        assert!(target_xml.contains(&copied));
        assert_eq!(target_xml.matches("office:value-type=\"float\"").count(), 1);
        assert_eq!(
            target_xml.matches("office:value-type=\"boolean\"").count(),
            1
        );
        assert!(target_xml.contains("<!--destination-gap-->"));

        let commit = edit.commit()?;
        assert_eq!(commit.patch().operations()[0].op, "ods.sheet.transfer");
        assert!(commit.patch().operations()[0].target.contains("source:"));
        let identical = raw_identical_members(&destination_bytes, commit.snapshot().as_bytes())
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat("raw comparison failed".to_string())
            })?;
        assert!(identical.contains("Assets/untouched.bin"));
        assert!(identical.contains("Destination/opaque.bin"));

        let wire = commit.patch().to_deterministic_json()?;
        let durable = Patch::from_deterministic_json(&wire, destination.limits())?;
        assert_eq!(
            durable.apply(&destination)?.snapshot().as_bytes(),
            commit.snapshot().as_bytes()
        );
        assert!(durable.apply(&source).is_err());
        let inverse = Patch::from_deterministic_json(
            &durable.inverse().to_deterministic_json()?,
            destination.limits(),
        )?;
        assert_eq!(
            inverse.apply(commit.snapshot())?.snapshot().as_bytes(),
            destination_bytes.as_slice()
        );
    }
    Ok(())
}

#[test]
fn scalar_sheet_transfer_refuses_dependencies_and_keeps_destination_atomic() -> Result<()> {
    let destination_bytes = package(&content(&plain_table("Dest", "keep"), ""), &[]);
    let destination = Snapshot::from_bytes(destination_bytes.clone())?;
    let transfer_source =
        Snapshot::from_bytes(package(&content(&plain_table("Source", "value"), ""), &[]))?;
    let refused = [
        r#"<t:table t:name="Source"><t:table-row><t:table-cell t:formula="of:=A1"/></t:table-row></t:table>"#
            .to_string(),
        r#"<t:table t:name="Source"><t:table-row><t:table-cell t:style-name="foreign"/></t:table-row></t:table>"#
            .to_string(),
        r#"<t:table t:name="Source"><t:table-row><t:table-cell t:number-columns-spanned="2"/></t:table-row></t:table>"#
            .to_string(),
        r#"<t:table t:name="Source"><t:table-row><t:table-cell><x:p><x:span>opaque</x:span></x:p></t:table-cell></t:table-row></t:table>"#
            .to_string(),
        r#"<t:table t:name="Source"><t:table-row><t:table-cell office:value-type="currency" office:value="1"/></t:table-row></t:table>"#
            .to_string(),
        r#"<t:table t:name="Source"><t:table-row><t:table-cell office:value-type="date" office:date-value="2026-02-30"/></t:table-row></t:table>"#
            .to_string(),
        r#"<t:table t:name="Source"><t:table-row><t:table-cell office:value-type="time" office:time-value="not-a-duration"/></t:table-row></t:table>"#
            .to_string(),
    ];
    let mut transfer_checked = 0usize;
    for table in refused {
        let Ok(source) = Snapshot::from_bytes(package(&content(&table, ""), &[])) else {
            continue;
        };
        transfer_checked += 1;
        let mut edit = destination.edit();
        assert!(
            edit.transfer_plain_scalar_sheet_from(&source, "Source", "Copied", SheetPosition::Last)
                .is_err()
        );
        assert_eq!(edit.as_bytes(), destination_bytes.as_slice());
    }
    assert!(transfer_checked >= 4);

    let unknown = Snapshot::from_bytes(package(
        &content(
            r#"<t:table t:name="Source"><t:table-row><t:table-cell><x:p>opaque</x:p><u:extension/></t:table-cell></t:table-row></t:table>"#,
            r#"xmlns:u="urn:example:unknown""#,
        ),
        &[],
    ))?;
    let mut edit = destination.edit();
    assert!(
        edit.transfer_plain_scalar_sheet_from(&unknown, "Source", "Copied", SheetPosition::Last)
            .is_err()
    );
    assert_eq!(edit.as_bytes(), destination_bytes.as_slice());

    let signed = Snapshot::from_bytes(package(
        &content(&plain_table("Source", "signed"), ""),
        &[(
            "META-INF/documentsignatures.xml",
            b"<signature/>",
            "application/xml",
        )],
    ))?;
    let mut edit = destination.edit();
    assert!(
        edit.transfer_plain_scalar_sheet_from(&signed, "Source", "Copied", SheetPosition::Last)
            .is_err()
    );
    assert_eq!(edit.as_bytes(), destination_bytes.as_slice());

    let signed_destination = Snapshot::from_bytes(package(
        &content(&plain_table("Dest", "keep"), ""),
        &[(
            "META-INF/documentsignatures.xml",
            b"<signature/>",
            "application/xml",
        )],
    ))?;
    let signed_before = signed_destination.as_bytes().to_vec();
    let mut signed_edit = signed_destination.edit();
    assert!(
        signed_edit
            .transfer_plain_scalar_sheet_from(
                &transfer_source,
                "Source",
                "Copied",
                SheetPosition::Last,
            )
            .is_err()
    );
    assert_eq!(signed_edit.as_bytes(), signed_before.as_slice());

    let protected_destination = Snapshot::from_bytes(package(
        &content(
            r#"<t:table t:name="Dest" t:protected="true"><t:table-row><t:table-cell office:value-type="string"><x:p>keep</x:p></t:table-cell></t:table-row></t:table>"#,
            "",
        ),
        &[],
    ))?;
    let protected_before = protected_destination.as_bytes().to_vec();
    let mut protected_edit = protected_destination.edit();
    assert!(
        protected_edit
            .transfer_plain_scalar_sheet_from(
                &transfer_source,
                "Source",
                "Copied",
                SheetPosition::Last,
            )
            .is_err()
    );
    assert_eq!(protected_edit.as_bytes(), protected_before.as_slice());
    Ok(())
}

#[test]
fn scalar_sheet_transfer_refuses_name_and_namespace_ambiguity() -> Result<()> {
    let source = Snapshot::from_bytes(package(&content(&plain_table("Source", "value"), ""), &[]))?;
    let collision = Snapshot::from_bytes(package(&content(&plain_table("Dest", "a"), ""), &[]))?;
    let before = collision.as_bytes().to_vec();
    let mut edit = collision.edit();
    assert!(
        edit.transfer_plain_scalar_sheet_from(&source, "Source", "Dest", SheetPosition::Last)
            .is_err()
    );
    assert_eq!(edit.as_bytes(), before.as_slice());

    let destination = Snapshot::from_bytes(package(
        &format!(
            r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE}" xmlns:t="{TABLE}" xmlns:y="{TEXT}" office:version="1.3"><office:body><office:spreadsheet><t:table t:name="Dest"><t:table-row><t:table-cell office:value-type="string"><y:p>value</y:p></t:table-cell></t:table-row></t:table></office:spreadsheet></office:body></office:document-content>"#
        ),
        &[],
    ))?;
    let mut edit = destination.edit();
    assert!(
        edit.transfer_plain_scalar_sheet_from(&source, "Source", "Copied", SheetPosition::Last)
            .is_err()
    );
    assert_eq!(edit.as_bytes(), destination.as_bytes());

    let shadow_destination = Snapshot::from_bytes(package(
        &format!(
            r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE}" xmlns:t="{TABLE}" xmlns:d="{TABLE}" xmlns:x="{TEXT}" xmlns:y="{TEXT}" office:version="1.3"><office:body><office:spreadsheet xmlns:t="urn:example:shadow"><d:table d:name="Dest"><d:table-row><d:table-cell office:value-type="string"><y:p>value</y:p></d:table-cell></d:table-row></d:table></office:spreadsheet></office:body></office:document-content>"#
        ),
        &[],
    ))?;
    let shadow_before = shadow_destination.as_bytes().to_vec();
    let mut shadow_edit = shadow_destination.edit();
    assert!(
        shadow_edit
            .transfer_plain_scalar_sheet_from(&source, "Source", "Copied", SheetPosition::Last,)
            .is_err()
    );
    assert_eq!(shadow_edit.as_bytes(), shadow_before.as_slice());

    let self_closing_destination = Snapshot::from_bytes(package(
        &format!(
            r#"<?xml version="1.0"?><office:document-content xmlns:office="{OFFICE}" xmlns:t="{TABLE}" xmlns:x="{TEXT}" office:version="1.3"><office:body><office:spreadsheet/></office:body></office:document-content>"#
        ),
        &[],
    ))?;
    let self_closing_before = self_closing_destination.as_bytes().to_vec();
    let mut self_closing_edit = self_closing_destination.edit();
    assert!(
        self_closing_edit
            .transfer_plain_scalar_sheet_from(&source, "Source", "Copied", SheetPosition::Last,)
            .is_err()
    );
    assert_eq!(self_closing_edit.as_bytes(), self_closing_before.as_slice());

    let mut edit = Snapshot::from_bytes(destination_bytes_for_empty())?.edit();
    assert!(
        edit.transfer_plain_scalar_sheet_from(&source, "Source", "", SheetPosition::Last)
            .is_err()
    );
    Ok(())
}

fn destination_bytes_for_empty() -> Vec<u8> {
    package(
        &content("", ""),
        &[(
            "Empty/opaque.bin",
            b"empty destination",
            "application/octet-stream",
        )],
    )
}
