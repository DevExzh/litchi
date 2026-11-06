//! Typed visual-object reads and source edits on unmodified PowerPoint fixtures.

use litchi_core::Position;
use litchi_pptx::{
    Package,
    shape::{Scene, Shape},
};
use std::collections::BTreeMap;

fn native(name: &str) -> Package {
    Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(name),
    )
    .unwrap()
}

fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let package = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    package
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let data = package.read_member(&name).unwrap();
            (name, data)
        })
        .collect()
}

fn assert_only_changed(before: &[u8], after: &[u8], expected: &[&str]) {
    let a = members(before);
    let b = members(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    let changed = a
        .iter()
        .filter(|(name, bytes)| **bytes != b[*name])
        .map(|(name, _)| name.as_str())
        .collect::<Vec<_>>();
    assert_eq!(changed, expected);
}

fn table_values(package: &Package) -> Vec<Vec<String>> {
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let shapes = slides[2].shapes().unwrap();
    let tables = shapes
        .iter()
        .filter_map(|s| {
            if let Shape::Table(table) = s {
                Some(table)
            } else {
                None
            }
        })
        .collect::<Vec<_>>();
    assert_eq!(tables.len(), 1);
    let table =
        litchi_pptx::table::Table::from_graphic_frame(tables[0].common().xml().unwrap()).unwrap();
    table
        .rows()
        .unwrap()
        .iter()
        .map(|row| {
            row.cells()
                .unwrap()
                .iter()
                .map(|cell| cell.text().unwrap())
                .collect()
        })
        .collect()
}

#[test]
fn native_table_reads_every_cell_dimension_and_style_switch() {
    use litchi_pptx::table::style::{Id, Parts};
    let package = native("powerpoint-text-table.pptx");
    assert_eq!(
        table_values(&package),
        [
            ["Item", "Quantity", "Total"],
            ["Apples", "3", "3.75"],
            ["Pears", "5", "12.50"],
            ["Tea", "2", "9.50"]
        ]
    );
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let shapes = slides[2].shapes().unwrap();
    let shape = shapes.iter().find(|s| s.id() == Some(4)).unwrap();
    assert!(matches!(shape, Shape::Table(_)));
    let table = litchi_pptx::table::Table::from_graphic_frame(shape.xml().unwrap()).unwrap();
    assert_eq!(
        (table.row_count().unwrap(), table.column_count().unwrap()),
        (4, 3)
    );
    assert_eq!(table.grid_columns().unwrap(), [3_505_199; 3]);
    assert_eq!(
        table
            .rows()
            .unwrap()
            .iter()
            .map(|r| r.height().unwrap())
            .collect::<Vec<_>>(),
        [Some(370_840); 4]
    );
    let properties = table.properties().unwrap().unwrap();
    assert_eq!(
        properties,
        litchi_pptx::table::Properties {
            first_row: true,
            band_row: true,
            style_id: Some("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}".into()),
            ..Default::default()
        }
    );
    for row in table.rows().unwrap() {
        for cell in row.cells().unwrap() {
            assert_eq!(
                cell.properties().unwrap(),
                litchi_pptx::table::CellProperties::default()
            );
        }
    }
    let styles = package.styles().unwrap().unwrap();
    let id = Id::parse(properties.style_id.as_deref().unwrap()).unwrap();
    assert_eq!(styles.default(), id);
    let style = styles.get(id).unwrap();
    assert_eq!(style.name(), "Medium Style 2 - Accent 1");
    for part in [
        Parts::WHOLE,
        Parts::FIRST_ROW,
        Parts::ODD_ROW,
        Parts::EVEN_ROW,
    ] {
        assert!(style.has(part));
    }
}

#[test]
fn native_table_row_column_and_style_edits_reopen_and_preserve_other_parts() {
    let package = native("powerpoint-text-table.pptx");
    let before = package.to_bytes().unwrap();
    let original = table_values(&package);
    let mut edit = package.edit().unwrap();
    edit.insert_table_row(
        2_usize,
        "Content Placeholder 3",
        Position::new(4),
        Some(400_000),
        &["Coffee", "4", "20.00"],
    )
    .unwrap();
    edit.insert_table_column(
        2_usize,
        "Content Placeholder 3",
        Position::new(3),
        Some(1_000_000),
        "new",
    )
    .unwrap();
    edit.set_table_column_width(
        2_usize,
        "Content Placeholder 3",
        Position::new(1),
        3_000_000,
    )
    .unwrap();
    assert!(
        edit.set_table_properties(2_usize, "Content Placeholder 3", |p| {
            p.band_row = false;
            p.first_col = true;
        })
        .unwrap()
    );
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let out = tempfile::NamedTempFile::with_suffix(".pptx").unwrap();
    commit.package().save(out.path()).unwrap();
    let reopened = Package::open(out.path()).unwrap();
    let mut expected = original.clone();
    expected.push(vec!["Coffee".into(), "4".into(), "20.00".into()]);
    for row in &mut expected {
        row.push("new".into());
    }
    assert_eq!(table_values(&reopened), expected);
    let presentation = reopened.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let shapes = slides[2].shapes().unwrap();
    let shape = shapes.iter().find(|s| s.id() == Some(4)).unwrap();
    let table = litchi_pptx::table::Table::from_graphic_frame(shape.xml().unwrap()).unwrap();
    assert_eq!(
        table.grid_columns().unwrap(),
        [3_505_199, 3_000_000, 3_505_199, 1_000_000]
    );
    assert_eq!(
        table
            .rows()
            .unwrap()
            .iter()
            .map(|r| r.height().unwrap())
            .collect::<Vec<_>>(),
        [
            Some(370_840),
            Some(370_840),
            Some(370_840),
            Some(370_840),
            Some(400_000)
        ]
    );
    let properties = table.properties().unwrap().unwrap();
    assert!(!properties.band_row);
    assert!(properties.first_row && properties.first_col);
    assert_eq!(
        properties.style_id.as_deref(),
        Some("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
    );
    let bounds = shape.bounds().unwrap();
    assert_eq!((bounds.width(), bounds.height()), (11_010_398, 1_883_360));
    assert_only_changed(&before, &after, &["ppt/slides/slide3.xml"]);
    let mut remove = reopened.edit().unwrap();
    remove
        .remove_table_column(2_usize, "Content Placeholder 3", Position::new(3))
        .unwrap();
    remove
        .remove_table_row(2_usize, "Content Placeholder 3", Position::new(4))
        .unwrap();
    let removed = remove.commit().unwrap();
    assert_eq!(table_values(removed.package()), original);
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        before
    );
    assert_eq!(package.to_bytes().unwrap(), before);
}

fn assert_chart_values(chart: &litchi_drawingml::chart::Chart) {
    use litchi_drawingml::chart::{
        TitleText,
        plot_area::TypeGroup,
        types::{BarDirection, BarGrouping},
    };
    assert_eq!(chart.plot_area.type_groups.len(), 1);
    let TypeGroup::Bar(group) = &chart.plot_area.type_groups[0] else {
        panic!("native column group")
    };
    assert_eq!(group.direction, BarDirection::Column);
    assert_eq!(group.grouping, BarGrouping::Clustered);
    assert_eq!(group.common.series.len(), 3);
    for (i, (series, values)) in group
        .common
        .series
        .iter()
        .zip([
            [10.0, 20.0, 30.0, 40.0],
            [15.0, 18.0, 21.0, 24.0],
            [8.0, 12.0, 16.0, 20.0],
        ])
        .enumerate()
    {
        let index = u32::try_from(i).unwrap();
        assert_eq!((series.index, series.order), (index, index));
        let column = ["B", "C", "D"][i];
        let TitleText::Reference(title) = series.title.as_ref().unwrap() else {
            panic!("series title formula")
        };
        assert_eq!(title.formula, format!("Sheet1!${column}$1"));
        let categories = series.categories.as_ref().unwrap();
        assert_eq!(categories.values, ["Q1", "Q2", "Q3", "Q4"]);
        assert_eq!(
            categories.source_ref.as_ref().unwrap().formula,
            "Sheet1!$A$2:$A$5"
        );
        let data = series.values.as_ref().unwrap();
        assert_eq!(data.values, values);
        assert_eq!(
            data.source_ref.as_ref().unwrap().formula,
            format!("Sheet1!${column}$2:${column}$5")
        );
    }
    assert_eq!(
        chart
            .external_data
            .as_ref()
            .unwrap()
            .relationship_id
            .as_deref(),
        Some("rId3")
    );
}

#[test]
fn native_classic_chart_reads_exact_series_caches_formulas_and_owner() {
    for name in [
        "powerpoint-visual-objects.pptx",
        "powerpoint-visual-objects.potx",
        "powerpoint-visual-objects.ppsx",
    ] {
        let package = native(name);
        let presentation = package.presentation().unwrap();
        let slides = presentation.slides().unwrap();
        let charts = slides[3].charts().unwrap();
        assert_eq!(charts.len(), 1);
        let info = charts[0].chart_info().unwrap();
        assert_eq!(info.chart_type, litchi_pptx::chart::Type::Column);
        assert_eq!(info.title.as_deref(), Some("REGIONAL_TOTALS_002"));
        assert!(info.has_legend);
        assert!(info.tolerance_report().is_clean());
        let part = charts[0].part();
        assert_eq!(part.partname().as_str(), "/ppt/charts/chart1.xml");
        let chart = litchi_drawingml::chart::reader::read(
            part.blob(),
            &litchi_drawingml::chart::Limits::DEFAULT,
        )
        .unwrap();
        assert_chart_values(&chart);
        let workbook = part.rels().get("rId3").unwrap();
        assert!(!workbook.is_external());
        assert_eq!(
            workbook.reltype(),
            litchi_opc::constants::relationship_type::PACKAGE
        );
        assert_eq!(
            workbook.target_partname().unwrap().as_str(),
            "/ppt/embeddings/Microsoft_Excel_Worksheet.xlsx"
        );
    }
}

#[test]
fn native_chart_title_edit_retains_series_workbook_relationships_and_slide() {
    use litchi_drawingml::chart::{TitleText, edit};
    let package = native("powerpoint-visual-objects.pptx");
    let before = package.to_bytes().unwrap();
    let mut tx = package.edit().unwrap();
    let (xml, workbook) = tx.chart(3_usize, "Content Placeholder 3").unwrap();
    assert!(workbook.is_some());
    let chart = edit::Package::from_xml(xml.clone()).unwrap();
    let mut edit = chart.edit();
    edit.set_title(Some(TitleText::from_string("Native chart revised")))
        .unwrap();
    let chart_commit = edit.commit().unwrap();
    tx.set_chart(
        3_usize,
        "Content Placeholder 3",
        chart_commit.patch().after_xml().to_vec(),
        None,
    )
    .unwrap();
    let commit = tx.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let presentation = reopened.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let charts = slides[3].charts().unwrap();
    assert_eq!(
        charts[0].chart_info().unwrap().title.as_deref(),
        Some("Native chart revised")
    );
    assert_chart_values(
        &litchi_drawingml::chart::reader::read(
            charts[0].part().blob(),
            &litchi_drawingml::chart::Limits::DEFAULT,
        )
        .unwrap(),
    );
    assert_only_changed(&before, &after, &["ppt/charts/chart1.xml"]);
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        before
    );
}

#[test]
fn native_smartart_outline_and_text_edit_keep_cached_drawing_and_companions() {
    use litchi_drawingml::diagram::{Limits, definition::DiagramDefinition, edit};
    let package = native("powerpoint-visual-objects.pptx");
    let before = package.to_bytes().unwrap();
    let mut tx = package.edit().unwrap();
    let parts = tx.smartart(2_usize, "Content Placeholder 3").unwrap();
    assert_eq!(
        litchi_pptx::shape::text::extract(parts.drawing.as_deref().unwrap(), None).unwrap(),
        "AcquireTransformVerify"
    );
    let limits = Limits::default();
    let outline = edit::Outline::read(&parts.data, &limits).unwrap();
    assert_eq!(outline.document, "{7D12FA0F-A88C-8D4D-9939-BC5701C7BCC8}");
    assert_eq!(outline.drawing_relationship.as_deref(), Some("rId6"));
    assert_eq!(
        outline
            .entries
            .iter()
            .map(|e| (e.id.as_str(), e.text.as_str(), e.parent, e.depth))
            .collect::<Vec<_>>(),
        [
            ("{463C37F4-C1FC-EF4B-B90F-1B19D4D7634C}", "Acquire", None, 0),
            (
                "{B3EA89B3-4BA0-EB49-8F0D-EBA487AA6177}",
                "Transform",
                None,
                0
            ),
            ("{72FA4759-4F7A-744F-BDAC-5C08F83BAE8E}", "Verify", None, 0)
        ]
    );
    let layout =
        DiagramDefinition::parse_layout(std::str::from_utf8(&parts.layout).unwrap(), &limits)
            .unwrap();
    assert_eq!(
        layout.unique_id.as_deref(),
        Some("urn:microsoft.com/office/officeart/2005/8/layout/process1")
    );
    assert!(
        layout
            .categories
            .iter()
            .any(|c| c.category_type == "process")
    );
    let id = &outline.entries[1].id;
    let mut changed = parts.clone();
    changed.drawing = Some(
        edit::drawing_text(
            parts.drawing.as_deref().unwrap(),
            &parts.data,
            id,
            "Transform revised",
            &limits,
        )
        .unwrap()
        .expect("native cached drawing maps to content node"),
    );
    changed.data = edit::set_text(&parts.data, id, "Transform revised", &limits).unwrap();
    assert!(
        tx.set_smartart(2_usize, "Content Placeholder 3", changed.clone())
            .unwrap()
    );
    let commit = tx.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let actual = reopened
        .edit()
        .unwrap()
        .smartart(2_usize, "Content Placeholder 3")
        .unwrap();
    assert_eq!(actual, changed);
    assert_eq!(
        litchi_pptx::shape::text::extract(actual.drawing.as_deref().unwrap(), None).unwrap(),
        "AcquireTransform revisedVerify"
    );
    let mut expected = outline;
    expected.entries[1].text = "Transform revised".into();
    assert_eq!(
        edit::Outline::read(&actual.data, &limits).unwrap(),
        expected
    );
    assert_only_changed(
        &before,
        &after,
        &["ppt/diagrams/data1.xml", "ppt/diagrams/drawing1.xml"],
    );
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        before
    );
}

#[test]
fn native_group_preserves_order_ids_transforms_and_source_text_edit() {
    let package = native("powerpoint-visual-objects.pptx");
    let before = package.to_bytes().unwrap();
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let scene = Scene::read(
        slides[4].part().part().blob(),
        litchi_pptx::shape::Limits::DEFAULT,
    )
    .unwrap();
    let Shape::Group(group) = scene.get("Group 6").unwrap().unwrap() else {
        panic!("native group")
    };
    assert_eq!(group.common().id(), Some(7));
    let children = group.shapes().collect::<Vec<_>>();
    assert_eq!(children.len(), 3);
    assert_eq!(
        children.iter().map(|s| s.id()).collect::<Vec<_>>(),
        [Some(4), Some(5), Some(6)]
    );
    assert_eq!(children[0].text(), Some("SOURCE_002"));
    assert_eq!(children[1].text(), Some("TARGET_002"));
    assert!(matches!(children[2], Shape::Connector(_)));
    for (shape, expected) in children.iter().zip([
        (1_371_600, 2_743_200, 2_743_200, 914_400),
        (7_772_400, 4_114_800, 2_743_200, 914_400),
        (4_114_800, 3_200_400, 3_657_600, 1_371_600),
    ]) {
        let b = shape.bounds().unwrap();
        assert_eq!((b.x(), b.y(), b.width(), b.height()), expected);
    }
    let transform = group.common().transform().unwrap().unwrap();
    assert_eq!(
        transform.offset(),
        Some(&litchi_drawingml::transform::Point::emu(1_371_600, 2_743_200).unwrap())
    );
    assert_eq!(
        transform.extent(),
        Some(litchi_drawingml::transform::Size::emu(9_144_000, 2_286_000).unwrap())
    );
    assert_eq!(transform.offset(), transform.child_offset());
    assert_eq!(transform.extent(), transform.child_extent());
    assert!(!transform.flip_horizontal());
    assert!(!transform.flip_vertical());
    let mut tx = package.edit().unwrap();
    tx.set_shape_text(4_usize, "Rounded Rectangle 3", "SOURCE_REVISED")
        .unwrap();
    let commit = tx.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let presentation = reopened.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let updated = Scene::read(
        slides[4].part().part().blob(),
        litchi_pptx::shape::Limits::DEFAULT,
    )
    .unwrap();
    assert_eq!(
        updated.get("Rounded Rectangle 3").unwrap().unwrap().text(),
        Some("SOURCE_REVISED")
    );
    for name in ["Oval 4", "Elbow Connector 5"] {
        assert_eq!(
            updated.get(name).unwrap().unwrap().xml().unwrap(),
            scene.get(name).unwrap().unwrap().xml().unwrap()
        );
    }
    assert_only_changed(&before, &after, &["ppt/slides/slide5.xml"]);
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        before
    );
}
