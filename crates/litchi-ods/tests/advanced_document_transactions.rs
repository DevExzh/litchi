mod support;

use std::{fs, path::Path};

use litchi_core::{MergeChoice, Result};
use litchi_odf_common::{chart::ChartClass, core::OwnedPackage};
use litchi_ods::{
    Builder, Cell, CellValue, Row, Sheet, Spreadsheet,
    document::{
        BoundFormControl, CellProperties, CellStyle, CellStyleNode, ChartObject, Collision,
        Drawing, DrawingFrame, EncryptionWritePolicy, FormControl, History, JoinFailure,
        NumberStyleNode, Resource, RichRun, RichText, SecurityWritePolicy, SignatureWritePolicy,
        Snapshot, StyleGraph, TextProperties, TextStyleNode,
    },
    model::{
        conditional_format::{Condition, Format},
        sparkline::{Group, Item},
        structure::{Column, Visibility},
    },
};

fn compact_source() -> Result<Vec<u8>> {
    let mut builder = Builder::new();
    builder.add_sheet(Sheet::new("Data")?)?;
    builder.set_cell(
        "Data",
        0,
        0,
        Cell::new(CellValue::Text("seed".to_string()), "seed"),
    )?;
    builder.build()
}

fn deep_style_graph() -> StyleGraph {
    StyleGraph {
        number_styles: vec![NumberStyleNode {
            name: "Money2".to_string(),
            decimal_places: 2,
            min_integer_digits: 1,
            prefix: Some("$".to_string()),
            suffix: None,
        }],
        text_styles: vec![TextStyleNode {
            name: "StrongRun".to_string(),
            parent: None,
            text: TextProperties {
                color: Some("#224466".to_string()),
                font_family: Some("Liberation Sans".to_string()),
                font_size_pt: Some(11.0),
                bold: Some(true),
                italic: Some(false),
                underline: Some(true),
            },
        }],
        cell_styles: vec![
            CellStyleNode {
                name: "BaseInput".to_string(),
                parent: None,
                data_style: None,
                cell: CellProperties::default(),
                text: TextProperties::default(),
            },
            CellStyleNode {
                name: "MoneyInput".to_string(),
                parent: Some("BaseInput".to_string()),
                data_style: Some("Money2".to_string()),
                cell: CellProperties {
                    background: Some("#fff4cc".to_string()),
                    horizontal_align: Some("end".to_string()),
                    vertical_align: Some("middle".to_string()),
                    wrap: Some(true),
                    border: Some("0.02cm solid #224466".to_string()),
                },
                text: TextProperties::default(),
            },
        ],
    }
}

#[test]
fn deep_style_form_drawing_and_chart_dependencies_are_atomic_and_durable() -> Result<()> {
    let source = compact_source()?;
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    let rich = RichText::new(vec![vec![
        RichRun::Text("Total ".to_string()),
        RichRun::Span {
            text: "$12.00".to_string(),
            style_name: "StrongRun".to_string(),
        },
    ]])?;
    edit.set_rich_cell_text_with_styles("Data", 0, 0, &rich, &deep_style_graph())?;
    edit.set_cell_style("Data", 0, 0, "MoneyInput")?;
    edit.set_bound_form_controls_with_resources(
        &[BoundFormControl {
            id: "amount-picker".to_string(),
            label: "Choose amount".to_string(),
            linked_cell: Some("Data.A1".to_string()),
            source_range: Some("Data.A1:A2".to_string()),
            image_path: Some("Pictures/control.png".to_string()),
        }],
        vec![Resource::new(
            "Pictures/control.png",
            "image/png",
            vec![1, 2, 3],
        )?],
        Collision::Reject,
    )?;
    edit.put_drawing_frame_with_resource(
        "Data",
        &DrawingFrame {
            name: "PositionedImage".to_string(),
            resource_path: "Pictures/frame.png".to_string(),
            anchor_cell: Some("Data.B2".to_string()),
            x_cm: 1.0,
            y_cm: 2.0,
            width_cm: 4.5,
            height_cm: 3.0,
            z_index: 7,
        },
        Resource::new("Pictures/frame.png", "image/png", vec![4, 5, 6])?,
        Collision::Reject,
    )?;
    edit.put_chart_definition(
        "Data",
        "SalesChart",
        "Object_1",
        &litchi_ods::charts::Definition::new(ChartClass::bar()),
        Collision::Reject,
    )?;

    let commit = edit.commit()?;
    let reopened = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
    let charts = reopened.charts()?;
    assert_eq!(charts.len(), 1);
    assert!(charts.named("SalesChart")?.is_some());
    for expected in [
        "style:name=\"Money2\"",
        "style:name=\"StrongRun\"",
        "style:parent-style-name=\"BaseInput\"",
        "form:linked-cell=\"Data.A1\"",
        "draw:name=\"PositionedImage\"",
        "draw:name=\"SalesChart\"",
    ] {
        assert!(reopened.content_xml().contains(expected));
    }
    for path in [
        "Pictures/control.png",
        "Pictures/frame.png",
        "Object_1/content.xml",
    ] {
        assert!(commit.snapshot().resource(path)?.is_some());
    }
    let wire = commit.patch().to_deterministic_json()?;
    let decoded = litchi_ods::document::Patch::from_deterministic_json(&wire, snapshot.limits())?;
    assert_eq!(
        decoded.apply(&snapshot)?.snapshot().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())?
            .snapshot()
            .as_bytes(),
        source
    );
    Ok(())
}

#[test]
fn deep_dependency_failures_leave_the_candidate_unchanged() -> Result<()> {
    let snapshot = Snapshot::from_bytes(compact_source()?)?;
    let mut edit = snapshot.edit();
    let before = edit.as_bytes().to_vec();
    let unresolved_rich = RichText::new(vec![vec![RichRun::Span {
        text: "unresolved".to_string(),
        style_name: "MissingStyle".to_string(),
    }]])?;
    assert!(
        edit.set_rich_cell_text_with_styles(
            "Data",
            0,
            0,
            &unresolved_rich,
            &StyleGraph::default(),
        )
        .is_err()
    );
    assert_eq!(edit.as_bytes(), before);
    let cyclic = StyleGraph {
        cell_styles: vec![
            CellStyleNode {
                name: "A".to_string(),
                parent: Some("B".to_string()),
                data_style: None,
                cell: CellProperties::default(),
                text: TextProperties::default(),
            },
            CellStyleNode {
                name: "B".to_string(),
                parent: Some("A".to_string()),
                data_style: None,
                cell: CellProperties::default(),
                text: TextProperties::default(),
            },
        ],
        ..StyleGraph::default()
    };
    assert!(edit.put_style_graph(&cyclic).is_err());
    assert_eq!(edit.as_bytes(), before);
    assert!(
        edit.set_bound_form_controls_with_resources(
            &[BoundFormControl {
                id: "missing-image".to_string(),
                label: "Missing".to_string(),
                linked_cell: Some("Data.A1:A2".to_string()),
                source_range: None,
                image_path: Some("Pictures/missing.png".to_string()),
            }],
            Vec::new(),
            Collision::Reject,
        )
        .is_err()
    );
    assert_eq!(edit.as_bytes(), before);
    assert!(
        edit.put_drawing_frame_with_resource(
            "Data",
            &DrawingFrame {
                name: "WrongSheet".to_string(),
                resource_path: "Pictures/wrong-sheet.png".to_string(),
                anchor_cell: Some("Other.A1".to_string()),
                x_cm: 0.0,
                y_cm: 0.0,
                width_cm: 1.0,
                height_cm: 1.0,
                z_index: 0,
            },
            Resource::new("Pictures/wrong-sheet.png", "image/png", vec![9])?,
            Collision::Reject,
        )
        .is_err()
    );
    assert_eq!(edit.as_bytes(), before);
    assert!(
        edit.put_chart_object(
            "Data",
            &ChartObject {
                name: "PrettyChart".to_string(),
                object_path: "Object_pretty".to_string(),
                content_xml: "<office:document-content>\n</office:document-content>".to_string(),
            },
            Collision::Reject,
        )
        .is_err()
    );
    assert_eq!(edit.as_bytes(), before);
    Ok(())
}

#[test]
fn unified_advanced_crud_reopens_and_is_durable() -> Result<()> {
    let source = compact_source()?;
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    edit.put_cell_style(&CellStyle {
        name: "Emphasis".to_string(),
        background: Some("#ffeeaa".to_string()),
        color: Some("#112233".to_string()),
        bold: Some(true),
    })?;
    let rich = RichText::new(vec![vec![
        RichRun::Text("Revenue".to_string()),
        RichRun::Space(1),
        RichRun::Span {
            text: "Q1".to_string(),
            style_name: "Emphasis".to_string(),
        },
        RichRun::LineBreak,
        RichRun::Link {
            text: "details".to_string(),
            href: "https://example.invalid/details".to_string(),
        },
    ]])?;
    edit.set_rich_cell_text("Data", 0, 0, &rich)?;
    edit.set_cell_formula("Data", 0, 0, "of:=SUM([.B1:.B2])")?;
    edit.set_cell_style("Data", 0, 0, "Emphasis")?;

    let mut row = Row::new();
    row.push_cell(Cell::new(CellValue::Number(3.0), "3"))?;
    edit.append_row("Data", &row)?;
    edit.append_column(
        "Data",
        &Column {
            index: 0,
            style_name: None,
            default_cell_style_name: Some("Emphasis".to_string()),
            visibility: Visibility::default(),
        },
    )?;
    edit.append_sheet(&Sheet::new("Summary")?)?;

    let format = Format::new(
        vec!["Data.A1:Data.A2".to_string()],
        vec![Condition::new("cell-content()>0", "Emphasis").into()],
    )?;
    edit.set_conditional_formats("Data", &[format])?;
    edit.set_sparkline_groups(
        "Data",
        &[Group::new(vec![Item::new(
            "Data.C1",
            vec!["Data.A1:Data.B1".to_string()],
        )])],
    )?;
    edit.set_form_controls(&[FormControl {
        id: "refresh".to_string(),
        label: "Refresh".to_string(),
    }])?;
    let resource = Resource::new("Pictures/chart.bin", "image/png", vec![1, 2, 3, 4])?;
    assert_eq!(
        edit.put_drawing_with_resource(
            "Data",
            &Drawing {
                name: "RevenueImage".to_string(),
                resource_path: "Pictures/chart.bin".to_string(),
            },
            resource,
            Collision::Reject,
        )?,
        litchi_ods::document::TransferDisposition::Added
    );

    let commit = edit.commit()?;
    let reopened = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
    assert_eq!(reopened.sheets().len(), 2);
    assert!(reopened.content_xml().contains("<text:span"));
    assert!(
        reopened
            .content_xml()
            .contains("<calcext:conditional-formats")
    );
    assert!(reopened.content_xml().contains("<calcext:sparkline-groups"));
    assert!(reopened.content_xml().contains("<office:forms"));
    assert!(reopened.content_xml().contains("RevenueImage"));
    assert!(commit.snapshot().resource("Pictures/chart.bin")?.is_some());
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())?
            .snapshot()
            .as_bytes(),
        source
    );
    let wire = commit.patch().to_deterministic_json()?;
    let decoded = litchi_ods::document::Patch::from_deterministic_json(&wire, snapshot.limits())?;
    assert_eq!(
        decoded.apply(&snapshot)?.snapshot().as_bytes(),
        commit.snapshot().as_bytes()
    );
    Ok(())
}

#[test]
fn real_calc_pretty_xml_keeps_source_formatting_outside_splice() -> Result<()> {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/odf/corpus/calc-two-sheets.ods");
    let source = fs::read(path)?;
    let snapshot = Snapshot::from_bytes(source)?;
    let before = Spreadsheet::from_bytes(snapshot.as_bytes().to_vec())?
        .content_xml()
        .to_string();
    let mut edit = snapshot.edit();
    edit.set_cell_formula("sheet1", 0, 0, "of:=1+2")?;
    let commit = edit.commit()?;
    let reopened = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
    assert_eq!(reopened.sheets().len(), 2);
    assert!(reopened.content_xml().contains("table:formula=\"of:=1+2\""));
    let preserved_prefix = before
        .split("<table:table-cell")
        .next()
        .ok_or_else(|| litchi_core::Error::InvalidFormat("cell prefix missing".to_string()))?;
    assert!(reopened.content_xml().starts_with(preserved_prefix));
    Ok(())
}

#[test]
fn raw_zip_pretty_provenance_and_malformed_negative_stay_test_only() -> Result<()> {
    let compact = Spreadsheet::from_bytes(compact_source()?)?
        .content_xml()
        .to_string();
    let pretty = compact.replace("><", ">\n  <");
    let raw = support::raw_package(&[("content.xml", pretty.as_bytes(), "text/xml")]);
    let snapshot = Snapshot::from_bytes(raw)?;
    let mut edit = snapshot.edit();
    edit.set_cell_formula("Data", 0, 0, "of:=40+2")?;
    let commit = edit.commit()?;
    let content = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?
        .content_xml()
        .to_string();
    let untouched = pretty
        .split("<table:table-cell")
        .next()
        .ok_or_else(|| litchi_core::Error::InvalidFormat("cell prefix missing".to_string()))?;
    assert!(content.starts_with(untouched));
    assert!(content.contains("table:formula=\"of:=40+2\""));

    let malformed = support::raw_package(&[(
        "content.xml",
        b"<office:document-content><office:body>",
        "text/xml",
    )]);
    assert!(Snapshot::from_bytes(malformed).is_err());
    Ok(())
}

#[test]
fn explicit_signature_policy_strips_stale_signatures_for_commit_and_apply() -> Result<()> {
    let content = Spreadsheet::from_bytes(compact_source()?)?
        .content_xml()
        .to_string();
    let signed_bytes = support::raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        (
            "META-INF/documentsignatures.xml",
            br#"<ds:document-signatures xmlns:ds="urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"/>"#,
            "text/xml",
        ),
    ]);
    let signed = Snapshot::from_bytes(signed_bytes)?;
    let mut refused = signed.edit();
    refused.set_cell_formula("Data", 0, 0, "of:=6*7")?;
    assert!(refused.commit().is_err());

    let policy = SecurityWritePolicy {
        signatures: SignatureWritePolicy::StripInvalidated,
        encryption: EncryptionWritePolicy::Refuse,
    };
    let mut allowed = signed.edit();
    allowed.set_cell_formula("Data", 0, 0, "of:=6*7")?;
    let commit = allowed.commit_with_security(policy)?;
    let package = OwnedPackage::from_bytes(commit.snapshot().as_bytes().to_vec())?;
    assert!(!package.has_file("META-INF/documentsignatures.xml")?);
    assert!(commit.patch().apply(&signed).is_err());
    assert_eq!(
        commit
            .patch()
            .apply_with_security(&signed, policy)?
            .snapshot()
            .as_bytes(),
        commit.snapshot().as_bytes()
    );
    Ok(())
}

#[test]
fn dependency_and_adversarial_inputs_fail_atomically() -> Result<()> {
    let snapshot = Snapshot::from_bytes(compact_source()?)?;
    let mut edit = snapshot.edit();
    let before = edit.as_bytes().to_vec();
    let mismatch = edit.put_drawing_with_resource(
        "Data",
        &Drawing {
            name: "bad".to_string(),
            resource_path: "Pictures/expected.bin".to_string(),
        },
        Resource::new("Pictures/actual.bin", "application/octet-stream", vec![1])?,
        Collision::Reject,
    );
    assert!(mismatch.is_err());
    assert_eq!(edit.as_bytes(), before);
    assert!(RichText::new(vec![vec![RichRun::Space(0)]]).is_err());
    assert!(
        RichText::new(vec![vec![RichRun::Link {
            text: "bad".to_string(),
            href: "javascript:alert(1)".to_string(),
        }]])
        .is_err()
    );
    assert!(
        edit.put_cell_style(&CellStyle {
            name: "Bad".to_string(),
            background: Some("red".to_string()),
            color: None,
            bold: None,
        })
        .is_err()
    );
    assert_eq!(edit.as_bytes(), before);

    let mut source_edit = snapshot.edit();
    let _source_disposition = source_edit.put_resource(
        Resource::new("Pictures/source.bin", "image/png", vec![8, 9])?,
        Collision::Reject,
    )?;
    let transfer_source = source_edit.commit()?.into_snapshot();
    let mut destination_edit = snapshot.edit();
    let _transfer_disposition = destination_edit.transfer_drawing(
        &transfer_source,
        "Pictures/source.bin",
        "Data",
        "Transferred",
        "Pictures/copied.bin",
        Collision::Reject,
    )?;
    let transferred = destination_edit.commit()?;
    assert!(
        transferred
            .snapshot()
            .resource("Pictures/copied.bin")?
            .is_some()
    );
    assert!(
        Spreadsheet::from_bytes(transferred.snapshot().as_bytes().to_vec())?
            .content_xml()
            .contains("xlink:href=\"Pictures/copied.bin\"")
    );
    Ok(())
}

#[test]
fn advanced_operations_use_join_three_way_and_history() -> Result<()> {
    let snapshot = Snapshot::from_bytes(compact_source()?)?;
    let mut rich_edit = snapshot.edit();
    rich_edit.set_rich_cell_text(
        "Data",
        0,
        0,
        &RichText::new(vec![vec![RichRun::Text("left".to_string())]])?,
    )?;
    let rich_commit = rich_edit.commit()?;

    let mut resource_edit = snapshot.edit();
    let _disposition = resource_edit.put_resource(
        Resource::new("Data/independent.bin", "application/octet-stream", vec![7])?,
        Collision::Reject,
    )?;
    let resource_commit = resource_edit.commit()?;
    let joined = rich_commit
        .patch()
        .join(resource_commit.patch())
        .map_err(|error| litchi_core::Error::InvalidFormat(error.detail().to_string()))?;
    let joined_commit = joined.apply(&snapshot)?;
    assert!(
        joined_commit
            .snapshot()
            .resource("Data/independent.bin")?
            .is_some()
    );
    assert!(
        Spreadsheet::from_bytes(joined_commit.snapshot().as_bytes().to_vec())?
            .content_xml()
            .contains("left")
    );

    let mut competing_edit = snapshot.edit();
    competing_edit.set_cell_formula("Data", 0, 0, "of:=2")?;
    let competing = competing_edit.commit()?;
    let conflict = rich_commit.patch().join(competing.patch()).err();
    assert!(matches!(
        conflict
            .as_ref()
            .map(litchi_ods::document::JoinError::failure),
        Some(JoinFailure::Conflict(_))
    ));
    let mut plan = litchi_ods::document::Patch::three_way(rich_commit.patch(), competing.patch())?;
    let _plan = plan.resolve(MergeChoice::Left);
    assert!(
        Spreadsheet::from_bytes(
            plan.finish()?
                .apply(&snapshot)?
                .snapshot()
                .as_bytes()
                .to_vec()
        )?
        .content_xml()
        .contains("left")
    );

    let mut history = History::new(snapshot);
    let _discarded = history.record(rich_commit)?;
    assert!(history.undo());
    assert!(history.redo());
    Ok(())
}
