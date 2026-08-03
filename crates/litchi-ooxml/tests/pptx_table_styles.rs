use litchi_ooxml::OoxmlError;
use litchi_ooxml::pptx::Package;
use litchi_ooxml::pptx::shape::Shape;
use litchi_ooxml::pptx::table::Table;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::part::BlobPart;
use litchi_opc::{PackURI, PackageWriter};
use litchi_pptx::table::style::{Conformance, Def, Id, Parts};
use tempfile::NamedTempFile;

const SHAPES: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/ooxml/pptx/shapes.pptx"
);

#[test]
fn package_loads_table_styles_part() {
    let package = Package::open(SHAPES).unwrap();

    let styles = package.styles().unwrap().unwrap();
    assert_eq!(
        styles.default().to_string(),
        "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
    );
    assert_eq!(styles.styles().len(), 2);

    let medium = styles
        .get(Id::parse("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}").unwrap())
        .unwrap();
    assert_eq!(medium.name(), "Medium Style 2 - Accent 1");
    for part in [
        Parts::WHOLE,
        Parts::ODD_ROW,
        Parts::EVEN_ROW,
        Parts::ODD_COLUMN,
        Parts::EVEN_COLUMN,
        Parts::FIRST_COLUMN,
        Parts::FIRST_ROW,
        Parts::LAST_COLUMN,
        Parts::LAST_ROW,
    ] {
        assert!(medium.has(part), "missing part style {part:?}");
    }
    assert!(!medium.has(Parts::NORTH_WEST));

    let plain = styles
        .get(Id::parse("{5940675A-B579-460E-94D1-54222C63F5DA}").unwrap())
        .unwrap();
    assert!(plain.has(Parts::WHOLE));
    assert!(!plain.has(Parts::FIRST_ROW));
}

#[test]
fn slide_tables_report_style_switches_and_references() {
    let package = Package::open(SHAPES).unwrap();
    let presentation = package.presentation().unwrap();

    let mut found = Vec::new();
    for slide in presentation.slides().unwrap() {
        for shape in slide.shapes().unwrap().iter() {
            if let Shape::Table(shape) = shape {
                let table = Table::from_graphic_frame_xml(shape.common().xml().unwrap()).unwrap();
                let properties = table.properties().unwrap().unwrap();
                found.push(properties);
            }
        }
    }

    assert!(!found.is_empty());
    assert!(
        found
            .iter()
            .all(|properties| properties.first_row && properties.band_row)
    );
    let medium = found.iter().find(|properties| {
        properties.style_id.as_deref() == Some("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
    });
    let plain = found.iter().find(|properties| {
        properties.style_id.as_deref() == Some("{5940675A-B579-460E-94D1-54222C63F5DA}")
    });
    assert!(medium.is_some());
    assert!(plain.is_some());

    // Every referenced style resolves in the package's table styles part.
    let styles = package.styles().unwrap().unwrap();
    for properties in &found {
        let style_id = properties.style_id.as_deref().unwrap();
        assert!(
            styles.get(Id::parse(style_id).unwrap()).is_some(),
            "unresolved style {style_id}"
        );
    }
}

#[test]
fn presentation_styles_match_package_level() {
    let package = Package::open(SHAPES).unwrap();
    let from_package = package.styles().unwrap().unwrap();
    let from_presentation = package.presentation().unwrap().styles().unwrap().unwrap();
    assert_eq!(from_package.default(), from_presentation.default());
    assert_eq!(from_package.source_xml(), from_presentation.source_xml());
}

#[test]
fn package_updates_and_removes_catalog_without_legacy_writer_state() {
    let mut package = Package::open(SHAPES).unwrap();
    let medium = Id::parse("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}").unwrap();
    let mut styles = package.styles().unwrap().unwrap();
    styles.rename(medium, "Renamed safely").unwrap();
    assert!(package.put_styles(styles).unwrap());

    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    package.save(output.path()).unwrap();
    let mut reopened = Package::open(output.path()).unwrap();
    assert_eq!(
        reopened
            .styles()
            .unwrap()
            .unwrap()
            .get(medium)
            .unwrap()
            .name(),
        "Renamed safely"
    );

    let removed = reopened.remove_styles().unwrap().unwrap();
    assert_eq!(
        removed.default().to_string(),
        "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
    );
    assert!(reopened.styles().unwrap().is_none());
    assert!(reopened.remove_styles().unwrap().is_none());
}

#[test]
fn style_edit_composes_with_a_new_deck_slide_edit() {
    let mut package = Package::new().unwrap();
    let created = Id::parse("{11111111-2222-3333-4444-555555555555}").unwrap();
    let mut styles = package.styles().unwrap().unwrap();
    styles.add(Def::new(created, "Created").unwrap()).unwrap();
    styles.set_default(created);
    assert!(package.put_styles(styles).unwrap());

    package.presentation_mut().unwrap().add_slide().unwrap();
    assert_eq!(package.styles().unwrap().unwrap().default(), created);

    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    package.save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    assert_eq!(reopened.presentation().unwrap().slide_count().unwrap(), 1);
    assert_eq!(reopened.styles().unwrap().unwrap().default(), created);
}

#[test]
fn style_removal_composes_with_a_new_deck_slide_edit() {
    let mut package = Package::new().unwrap();
    assert!(package.remove_styles().unwrap().is_some());
    package.presentation_mut().unwrap().add_slide().unwrap();
    assert!(package.styles().unwrap().is_none());

    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    package.save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    assert_eq!(reopened.presentation().unwrap().slide_count().unwrap(), 1);
    assert!(reopened.styles().unwrap().is_none());
}

#[test]
fn noncanonical_style_target_survives_transactional_raw_save() {
    let mut package = Package::new().unwrap();
    let styles = package.remove_styles().unwrap().unwrap();
    package
        .edit_opc(|opc| {
            opc.add_part(Box::new(BlobPart::new(
                PackURI::new("/ppt/tableStyles.xml").unwrap(),
                "application/xml".into(),
                b"<occupied/>".to_vec(),
            )));
            Ok(())
        })
        .unwrap();
    assert!(package.put_styles(styles).unwrap());

    // Exercise a producer-selected style reference that occupies the
    // historical rId1 slide-master slot.
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let (master_id, master_target, style_id, style_target) = {
        let relationships = package
            .opc()
            .unwrap()
            .get_part(&presentation_name)
            .unwrap()
            .rels();
        let master = relationships
            .iter()
            .find(|relationship| relationship.reltype() == rt::SLIDE_MASTER)
            .unwrap();
        let style = relationships
            .iter()
            .find(|relationship| relationship.reltype() == rt::TABLE_STYLES)
            .unwrap();
        (
            master.r_id().to_owned(),
            master.target_ref().to_owned(),
            style.r_id().to_owned(),
            style.target_ref().to_owned(),
        )
    };
    package
        .edit_opc(|opc| {
            let presentation = opc.get_part_mut(&presentation_name)?;
            assert!(presentation.rels_mut().remove(&master_id).is_some());
            assert!(presentation.rels_mut().remove(&style_id).is_some());
            presentation.rels_mut().add_relationship(
                rt::SLIDE_MASTER.into(),
                master_target,
                style_id,
                false,
            );
            presentation.rels_mut().add_relationship(
                rt::TABLE_STYLES.into(),
                style_target,
                master_id,
                false,
            );
            Ok(())
        })
        .unwrap();

    let before = package
        .opc()
        .unwrap()
        .main_document_part()
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::TABLE_STYLES)
        .unwrap();
    assert_eq!(before.target_ref(), "tableStyles2.xml");
    let before_id = before.r_id().to_owned();
    assert_eq!(before_id, "rId1");

    assert!(matches!(
        package.presentation_mut(),
        Err(OoxmlError::UnsafeEdit {
            operation: "presentation_mut",
            ..
        })
    ));
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    package.save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    let after = reopened
        .opc()
        .unwrap()
        .main_document_part()
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::TABLE_STYLES)
        .unwrap();
    assert_eq!(after.r_id(), before_id);
    assert_eq!(after.target_ref(), "tableStyles2.xml");
    assert!(reopened.styles().unwrap().is_some());
    assert_eq!(reopened.presentation().unwrap().slide_count().unwrap(), 0);
}

#[test]
fn strict_raw_edit_disables_the_transitional_legacy_writer() {
    let mut package = Package::new().unwrap();
    make_strict(&mut package);
    assert_eq!(
        package.styles().unwrap().unwrap().conformance(),
        Conformance::Strict
    );
    assert!(matches!(
        package.presentation_mut(),
        Err(OoxmlError::UnsafeEdit {
            operation: "presentation_mut",
            ..
        })
    ));
    let before = PackageWriter::to_bytes(package.opc().unwrap()).unwrap();
    let output = NamedTempFile::with_suffix(".pptx").unwrap();

    package.save(output.path()).unwrap();
    assert_eq!(
        PackageWriter::to_bytes(package.opc().unwrap()).unwrap(),
        before
    );
}

fn make_strict(package: &mut Package) {
    const TRANSITIONAL_PRESENTATION: &str =
        "http://schemas.openxmlformats.org/presentationml/2006/main";
    const STRICT_PRESENTATION: &str = "http://purl.oclc.org/ooxml/presentationml/main";
    const TRANSITIONAL_DRAWING: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const STRICT_DRAWING: &str = "http://purl.oclc.org/ooxml/drawingml/main";
    const STRICT_TABLE_STYLES: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/tableStyles";

    package
        .edit_opc(|opc| {
            let (root_id, root_target) = opc
                .rels()
                .iter()
                .find(|relationship| relationship.reltype() == rt::OFFICE_DOCUMENT)
                .map(|relationship| {
                    (
                        relationship.r_id().to_owned(),
                        relationship.target_ref().to_owned(),
                    )
                })
                .unwrap();
            assert!(opc.rels_mut().remove(&root_id).is_some());
            opc.rels_mut().add_relationship(
                rt::STRICT_OFFICE_DOCUMENT.into(),
                root_target,
                root_id,
                false,
            );

            let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
            let (style_id, style_target) = opc
                .get_part(&presentation_name)
                .unwrap()
                .rels()
                .iter()
                .find(|relationship| relationship.reltype() == rt::TABLE_STYLES)
                .map(|relationship| {
                    (
                        relationship.r_id().to_owned(),
                        relationship.target_ref().to_owned(),
                    )
                })
                .unwrap();
            let presentation = opc.get_part_mut(&presentation_name)?;
            let presentation_xml = String::from_utf8(presentation.blob().to_vec())
                .unwrap()
                .replace(TRANSITIONAL_PRESENTATION, STRICT_PRESENTATION);
            presentation.set_blob(presentation_xml.into_bytes());
            assert!(presentation.rels_mut().remove(&style_id).is_some());
            presentation.rels_mut().add_relationship(
                STRICT_TABLE_STYLES.into(),
                style_target,
                style_id,
                false,
            );

            let style_name = PackURI::new("/ppt/tableStyles.xml").unwrap();
            let style = opc.get_part_mut(&style_name)?;
            let style_xml = String::from_utf8(style.blob().to_vec())
                .unwrap()
                .replace(TRANSITIONAL_DRAWING, STRICT_DRAWING);
            style.set_blob(style_xml.into_bytes());
            Ok(())
        })
        .unwrap();
}
