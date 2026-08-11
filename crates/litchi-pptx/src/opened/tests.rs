//! Opened-presentation transaction regression tests.

use std::collections::BTreeMap;

use litchi_opc::{BlobPart, PackURI, TargetMode};

use super::{History, Limits, Patch, Resolution, ShapeTextReplacement};
use crate::{Error, Package, Result};

fn opened_two_slide_package() -> Result<Package> {
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        let first = presentation.add_slide()?;
        first.set_title("First title");
        first.add_text_box("First body", 10, 20, 300, 400);
        first.set_notes("First notes");
        let second = presentation.add_slide()?;
        second.set_title("Second title");
        second.add_text_box("Second body", 10, 20, 300, 400);
        second.set_notes("Second notes");
    }
    let bytes = package.to_bytes()?;
    Package::from_bytes(&bytes)
}

fn part_states(package: &Package) -> BTreeMap<String, (Vec<u8>, Vec<String>)> {
    package
        .opc
        .iter_parts()
        .map(|part| {
            let mut relationships: Vec<_> = part
                .rels()
                .iter()
                .map(|relationship| {
                    format!(
                        "{}\0{}\0{}\0{}",
                        relationship.r_id(),
                        relationship.reltype(),
                        relationship.target_ref(),
                        relationship.is_external()
                    )
                })
                .collect();
            relationships.sort_unstable();
            (
                part.partname().as_str().to_owned(),
                (part.blob().to_vec(), relationships),
            )
        })
        .collect()
}

fn add_group_connector_transfer_fixture(package: &mut Package) -> Result<()> {
    let slide = package.opened_presentation()?.slides()[0]
        .part_name()
        .clone();
    let image_name = PackURI::new("/ppt/media/group-transfer.png").map_err(Error::Invalid)?;
    package.opc.try_add_part(Box::new(BlobPart::new(
        image_name.clone(),
        "image/png".into(),
        vec![137, 80, 78, 71, 9, 8, 7, 6],
    )))?;
    let diagram_parts = [
        (
            "/ppt/diagrams/data1.xml",
            litchi_opc::constants::content_type::DML_DIAGRAM_DATA,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
            "rIdDiagramData",
        ),
        (
            "/ppt/diagrams/layout1.xml",
            litchi_opc::constants::content_type::DML_DIAGRAM_LAYOUT,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
            "rIdDiagramLayout",
        ),
        (
            "/ppt/diagrams/quickStyle1.xml",
            litchi_opc::constants::content_type::DML_DIAGRAM_STYLE,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
            "rIdDiagramStyle",
        ),
        (
            "/ppt/diagrams/colors1.xml",
            litchi_opc::constants::content_type::DML_DIAGRAM_COLORS,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
            "rIdDiagramColors",
        ),
    ];
    for (part_name, content_type, _, _) in diagram_parts {
        package.opc.try_add_part(Box::new(BlobPart::new(
            PackURI::new(part_name).map_err(Error::Invalid)?,
            content_type.into(),
            format!(r#"<dgm:test xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" name="{part_name}"/>"#).into_bytes(),
        )))?;
    }
    let ole_name =
        PackURI::new("/ppt/embeddings/oleObject-transfer.bin").map_err(Error::Invalid)?;
    package.opc.try_add_part(Box::new(BlobPart::new(
        ole_name.clone(),
        litchi_opc::constants::content_type::OFC_OLE_OBJECT.into(),
        vec![0xD0, 0xCF, 0x11, 0xE0, 1, 2, 3, 4],
    )))?;
    let target = image_name.relative_ref(slide.base_uri());
    let relationship_id = package
        .opc
        .get_part_mut(&slide)?
        .relate_to(&target, litchi_opc::constants::relationship_type::IMAGE);
    for (part_name, _, relationship_type, relationship_id) in diagram_parts {
        let part_name = PackURI::new(part_name).map_err(Error::Invalid)?;
        package
            .opc
            .get_part_mut(&slide)?
            .rels_mut()
            .try_add_relationship(
                relationship_type.into(),
                part_name.relative_ref(slide.base_uri()),
                relationship_id.into(),
                TargetMode::Internal,
            )?;
    }
    package
        .opc
        .get_part_mut(&slide)?
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::OLE_OBJECT.into(),
            ole_name.relative_ref(slide.base_uri()),
            "rIdOleTransfer".into(),
            TargetMode::Internal,
        )?;
    let fragment = format!(
        r#"<p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGrpSpPr><p:cNvPr id="20" name="Transfer group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/><a:chOff x="0" y="0"/><a:chExt cx="2000000" cy="1000000"/></a:xfrm></p:grpSpPr><p:pic><p:nvPicPr><p:cNvPr id="21" name="Group picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{relationship_id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="900000" cy="900000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic><p:sp><p:nvSpPr><p:cNvPr id="22" name="Group target"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1100000" y="0"/><a:ext cx="900000" cy="900000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp><p:cxnSp><p:nvCxnSpPr><p:cNvPr id="23" name="Group connector"/><p:cNvCxnSpPr><a:stCxn id="21" idx="0"/><a:endCxn id="22" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="900000" y="450000"/><a:ext cx="200000" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp></p:grpSp>"#
    );
    let free_connector = r#"<p:cxnSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvCxnSpPr><p:cNvPr id="24" name="Free connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>"#;
    let external_connector = r#"<p:cxnSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvCxnSpPr><p:cNvPr id="25" name="External connector"/><p:cNvCxnSpPr><a:stCxn id="20" idx="0"/><a:endCxn id="22" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>"#;
    let unresolved_connector = r#"<p:cxnSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvCxnSpPr><p:cNvPr id="26" name="Unresolved connector"/><p:cNvCxnSpPr><a:stCxn id="999" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>"#;
    let diagram = r#"<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="27" name="Dependency diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds rel:dm="rIdDiagramData" rel:lo="rIdDiagramLayout" rel:qs="rIdDiagramStyle" rel:cs="rIdDiagramColors"/></a:graphicData></a:graphic></p:graphicFrame>"#;
    let opaque_frame = r#"<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:opaque"><p:nvGraphicFramePr><p:cNvPr id="28" name="Opaque frame"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></p:xfrm><a:graphic><a:graphicData uri="urn:opaque"><x:payload/></a:graphicData></a:graphic></p:graphicFrame>"#;
    let ole = r#"<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:producer"><p:nvGraphicFramePr><p:cNvPr id="29" name="Inert OLE transfer"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj name="Inert" progId="Opaque.App" r:id="rIdOleTransfer" x:provenance="rIdOleTransfer"><p:embed/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>"#;
    let content_part =
        r#"<p:contentPart xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
    let unknown = r#"<p14:sp xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p14:nvSpPr><p14:cNvPr id="30" name="Opaque extension shape"/></p14:nvSpPr></p14:sp>"#;
    let owner = package.opc.get_part(&slide)?;
    let xml = super::xml::append_shape(owner.blob(), fragment.as_bytes())?;
    let xml = super::xml::append_shape(&xml, free_connector.as_bytes())?;
    let xml = super::xml::append_shape(&xml, external_connector.as_bytes())?;
    let xml = super::xml::append_shape(&xml, unresolved_connector.as_bytes())?;
    let xml = super::xml::append_shape(&xml, diagram.as_bytes())?;
    let xml = super::xml::append_shape(&xml, opaque_frame.as_bytes())?;
    let xml = super::xml::append_shape(&xml, ole.as_bytes())?;
    let xml = super::xml::append_shape(&xml, content_part.as_bytes())?;
    let xml = super::xml::append_shape(&xml, unknown.as_bytes())?;
    package.opc.get_part_mut(&slide)?.set_blob(xml);
    Ok(())
}

#[test]
fn composes_order_shape_and_notes_in_one_durable_inverse_commit() -> Result<()> {
    let mut package = opened_two_slide_package()?;
    let source = package.opened_presentation()?;
    let source_revision = source.revision();
    let before = part_states(&package);
    let mut edit = source.edit();
    assert!(edit.move_slide(0, 1)?);
    assert!(edit.set_shape_text(0_usize, 0_usize, "Edited <second> & title")?);
    assert!(edit.set_notes_text(0_usize, "Edited & second notes")?);
    let commit = edit.commit()?;
    assert_eq!(commit.patch().resource_count(), 3);
    let durable = commit.patch().to_bytes()?;
    let decoded = Patch::from_bytes(&durable)?;
    assert_eq!(&decoded, commit.patch());
    let inverse = decoded.inverse();
    let touched: Vec<_> = decoded
        .resources()
        .map(|name| name.as_str().to_owned())
        .collect();

    let published = package.apply_opened_presentation_patch(&decoded)?;
    assert_eq!(published.slides()[0].name(), "Slide 257");
    let presentation = package.presentation()?;
    assert_eq!(
        presentation
            .slide(0)?
            .ok_or_else(|| Error::Invalid("published slide disappeared".into()))?
            .text()?,
        "Edited <second> & title\nSecond body"
    );
    let notes = package
        .notes()?
        .ok_or_else(|| Error::Invalid("published notes disappeared".into()))?;
    assert_eq!(
        notes.slides()[0].text()?.as_deref(),
        Some("Edited & second notes")
    );
    let after = part_states(&package);
    for (name, state) in &before {
        if !touched.contains(name) {
            assert_eq!(after.get(name), Some(state), "untouched part {name}");
        }
    }

    let bytes = package.to_bytes()?;
    let mut reopened = Package::from_bytes(&bytes)?;
    assert_eq!(reopened.opened_presentation()?.slides()[0].id(), 257);
    let restored = reopened.apply_opened_presentation_patch(&inverse)?;
    assert_eq!(restored.revision(), source_revision);
    assert_eq!(part_states(&reopened), before);
    Ok(())
}

#[test]
fn atomic_shape_text_batch_composes_in_one_opened_transaction() -> Result<()> {
    let mut package = opened_two_slide_package()?;
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    assert_eq!(
        edit.set_shape_texts(
            0_usize,
            &[
                ShapeTextReplacement::at(1, "Batch body & <changed>"),
                ShapeTextReplacement::at(0, "Batch title"),
            ],
        )?,
        2
    );
    let commit = edit.commit()?;
    package.apply_opened_presentation_patch(commit.patch())?;
    let slide = package
        .presentation()?
        .slide(0)?
        .ok_or_else(|| Error::Invalid("batch slide disappeared".into()))?;
    let scene = slide.shapes()?;
    assert_eq!(scene.at(0)?.common().text(), Some("Batch title"));
    assert_eq!(scene.at(1)?.common().text(), Some("Batch body & <changed>"));
    Ok(())
}

#[test]
fn disjoint_patches_join_and_same_part_edits_conflict() -> Result<()> {
    let mut package = opened_two_slide_package()?;
    let source = package.opened_presentation()?;

    let mut shape = source.edit();
    shape.set_shape_text(0_usize, 0_usize, "Joined shape")?;
    let shape_patch = shape.commit()?.into_patch();

    let mut notes = source.edit();
    notes.set_notes_text(1_usize, "Joined notes")?;
    let notes_patch = notes.commit()?.into_patch();
    assert!(!shape_patch.conflicts_with(&notes_patch));
    let joined = shape_patch.join(&notes_patch)?;
    assert_eq!(joined.resource_count(), 2);
    package.apply_opened_presentation_patch(&joined)?;
    assert!(
        package
            .presentation()?
            .slide(0)?
            .is_some_and(|slide| { slide.text().is_ok_and(|text| text.contains("Joined shape")) })
    );

    let mut competing = source.edit();
    competing.set_shape_text(0_usize, 1_usize, "Competing shape")?;
    let competing = competing.commit()?.into_patch();
    assert!(shape_patch.conflicts_with(&competing));
    assert!(shape_patch.join(&competing).is_err());
    Ok(())
}

#[test]
fn stale_and_unsupported_raw_xml_fail_before_publication() -> Result<()> {
    let mut package = opened_two_slide_package()?;
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    edit.set_shape_text(0_usize, 0_usize, "Stale target")?;
    let patch = edit.commit()?.into_patch();
    let slide_name = patch
        .resources()
        .find(|name| name.as_str().contains("/slides/"))
        .cloned()
        .ok_or_else(|| Error::Invalid("shape patch has no slide resource".into()))?;
    let mut changed = package.opc.get_part(&slide_name)?.blob().to_vec();
    changed.extend_from_slice(b" ");
    package
        .opc
        .get_part_mut(&slide_name)?
        .set_blob(changed.clone());
    assert!(package.apply_opened_presentation_patch(&patch).is_err());
    assert_eq!(package.opc.get_part(&slide_name)?.blob(), changed);

    let mut package = opened_two_slide_package()?;
    let main = PackURI::new("/ppt/presentation.xml").map_err(Error::Invalid)?;
    let xml = std::str::from_utf8(package.opc.get_part(&main)?.blob())
        .map_err(|error| Error::Xml(error.to_string()))?
        .replacen("</p:sldIdLst>", "<p:ext/></p:sldIdLst>", 1)
        .into_bytes();
    package.opc.get_part_mut(&main)?.set_blob(xml);
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    assert!(edit.move_slide(0, 1).is_err());
    assert!(!edit.is_changed());

    let mut package = opened_two_slide_package()?;
    let slide = package.opened_presentation()?.slides()[0]
        .part_name()
        .clone();
    let xml = std::str::from_utf8(package.opc.get_part(&slide)?.blob())
        .map_err(|error| Error::Xml(error.to_string()))?
        .replacen(
            "<a:t>First title</a:t>",
            "<a:t><x:v xmlns:x=\"urn:hostile\"/></a:t>",
            1,
        )
        .into_bytes();
    package.opc.get_part_mut(&slide)?.set_blob(xml);
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    assert!(edit.set_shape_text(0_usize, 0_usize, "rejected").is_err());
    assert!(!edit.is_changed());
    Ok(())
}

#[test]
fn durable_decoder_and_history_enforce_finite_bounds() -> Result<()> {
    let package = opened_two_slide_package()?;
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    edit.set_shape_text(0_usize, 0_usize, "History")?;
    let patch = edit.commit()?.into_patch();
    let mut trailing = patch.to_bytes()?;
    trailing.push(0);
    assert!(Patch::from_bytes(&trailing).is_err());
    let tiny = Limits::new(1, 16, 1, 1, 1)
        .ok_or_else(|| Error::Invalid("test limits are invalid".into()))?;
    assert!(Patch::from_bytes_with_limits(&patch.to_bytes()?, tiny).is_err());

    let history_limits = Limits::new(
        4_096,
        128 * 1024 * 1024,
        8 * 1024 * 1024,
        2,
        256 * 1024 * 1024,
    )
    .ok_or_else(|| Error::Invalid("test history limits are invalid".into()))?;
    let mut history = History::new(history_limits);
    history.push(patch.clone())?;
    history.push(patch.clone())?;
    history.push(patch.clone())?;
    assert_eq!(history.len(), 2);
    assert!(history.encoded_bytes() > 0);
    assert_eq!(history.pop_inverse(), Some(patch.inverse()));
    assert_eq!(history.len(), 1);
    Ok(())
}

#[test]
fn real_powerpoint_fixture_round_trips_an_exact_no_op() -> Result<()> {
    let bytes = std::fs::read("../../test-data/ooxml/pptx/shapes.pptx")?;
    let mut package = Package::from_bytes(&bytes)?;
    let source = package.opened_presentation()?;
    assert!(source.slides().len() >= 2);
    let original_first = source.slides()[0].id();
    let slide = package
        .presentation()?
        .slide(0)?
        .ok_or_else(|| Error::Invalid("fixture slide disappeared".into()))?;
    let (shape_position, original_text) = slide
        .shapes()?
        .iter()
        .enumerate()
        .find_map(|(position, shape)| {
            shape
                .common()
                .text()
                .map(|text| (position, text.to_owned()))
        })
        .ok_or_else(|| Error::Invalid("fixture slide has no text shape".into()))?;
    let mut edit = source.edit();
    assert!(!edit.move_slide(0, 0)?);
    assert!(!edit.set_shape_text(0_usize, shape_position, &original_text)?);
    let patch = Patch::from_bytes(&edit.commit()?.into_patch().to_bytes()?)?;
    assert!(patch.is_empty());
    package.apply_opened_presentation_patch(&patch)?;
    let saved = package.to_bytes()?;
    let reopened = Package::from_bytes(&saved)?;
    assert_eq!(
        reopened.opened_presentation()?.slides()[0].id(),
        original_first
    );
    Ok(())
}

#[test]
fn one_opened_transaction_spans_complex_ordinary_domains_and_full_reopen() -> Result<()> {
    let mut package = opened_two_slide_package()?;
    let before = part_states(&package);
    let mut styles = package
        .styles()?
        .ok_or_else(|| Error::Invalid("authored package has no table styles".into()))?;
    let default_style = styles.default();
    styles.add(crate::table::style::Def::new(
        default_style,
        "Opened transaction style",
    )?)?;
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    edit.move_slide(0, 1)?;
    edit.set_shape_text(0_usize, 0_usize, "Complex transaction title")?;
    edit.set_notes_text(0_usize, "Complex transaction notes")?;
    assert!(edit.put_table_styles(styles)?);
    edit.add_table(
        0_usize,
        &[
            vec!["Quarter".into(), "Revenue".into()],
            vec!["Q1".into(), "10".into()],
        ],
        (100, 100, 2_000_000, 1_000_000),
    )?;

    let chart =
        crate::chart::Chart::new(crate::chart::Type::Column, 100, 200, 3_000_000, 2_000_000)
            .with_title("Quarterly chart")
            .add_series(
                crate::chart::Series::new("Revenue")
                    .with_categories(vec!["Q1".into(), "Q2".into()])
                    .with_values(vec![10.0, 12.0]),
            );
    edit.add_chart(0_usize, &chart)?;

    let media = crate::media_parts::List {
        pictures: vec![crate::media_parts::Picture {
            shape_id: 40,
            name: "clip.mp4".into(),
            kind: crate::media_parts::Kind::Video,
            relationship_id: "rIdOpenedVideo".into(),
            resource: Some(crate::media_parts::Resource::new(
                "/ppt/media/opened-video.mp4",
                "video/mp4",
                vec![0, 1, 2, 3, 4],
            )),
            poster: Some(crate::media_parts::Poster {
                relationship_id: "rIdOpenedPoster".into(),
                resource: Some(crate::media_parts::Resource::new(
                    "/ppt/media/opened-poster.png",
                    "image/png",
                    vec![137, 80, 78, 71],
                )),
            }),
            transform: Some(crate::media_parts::Transform::emu(10, 20, 300, 400)?),
            office_extension: None,
        }],
    };
    edit.store_media(
        1_usize,
        &media,
        crate::media_parts::Conformance::Transitional,
    )?;

    let master = edit.add_slide_master()?;
    edit.add_slide_layout(
        &master.part_name,
        crate::master_layout::SlideLayoutKind::Blank,
        "Opened Blank",
        &[],
    )?;
    edit.add_comment_author(
        crate::comments::Author::new(7, "Opened Author", "OA"),
        crate::comments::Conformance::Transitional,
    )?;
    edit.add_comment(
        0_usize,
        crate::comments::Comment::new(7, "Opened comment", 10, 20),
        crate::comments::Conformance::Transitional,
    )?;

    let commit = edit.commit()?;
    assert!(commit.patch().resource_count() >= 10);
    let durable = Patch::from_bytes(&commit.patch().to_bytes()?)?;
    let inverse = durable.inverse();
    package.apply_opened_presentation_patch(&durable)?;
    let bytes = package.to_bytes()?;
    let mut reopened = Package::from_bytes(&bytes)?;

    let opened = reopened.opened_presentation()?;
    let chart_slide = reopened.opc.get_part(opened.slides()[0].part_name())?;
    assert_eq!(crate::chart::related(&reopened.opc, chart_slide)?.len(), 1);
    assert!(
        chart_slide
            .blob()
            .windows(b"<a:tbl>".len())
            .any(|window| window == b"<a:tbl>")
    );
    assert_eq!(
        crate::media_parts::load(&reopened.opc, opened.slides()[1].part_name())?
            .pictures
            .len(),
        1
    );
    assert_eq!(
        crate::comments::load_presentation_comments(&reopened.opc)?
            .ok_or_else(|| Error::Invalid("comments disappeared".into()))?
            .slides[0]
            .comments[0]
            .text,
        "Opened comment"
    );
    assert!(
        reopened
            .styles()?
            .is_some_and(|catalog| catalog.named("Opened transaction style").count() == 1)
    );
    crate::master_layout::validate_master_layout_graph(&reopened.opc)?;

    reopened.apply_opened_presentation_patch(&inverse)?;
    assert_eq!(part_states(&reopened), before);
    Ok(())
}

#[test]
fn three_way_plan_is_immutable_and_history_supports_redo() -> Result<()> {
    let mut package = opened_two_slide_package()?;
    let base = package.opened_presentation()?;
    let mut left = base.edit();
    left.set_shape_text(0_usize, 0_usize, "Left title")?;
    left.set_notes_text(1_usize, "Merged notes")?;
    let left = left.commit()?.into_patch();
    let mut right = base.edit();
    right.set_shape_text(0_usize, 0_usize, "Right title")?;
    let right = right.commit()?.into_patch();

    let plan = Patch::three_way(&base, &left, &right)?;
    assert_eq!(plan.unresolved_count(), 1);
    assert_eq!(
        plan.unresolved_count(),
        1,
        "the original plan remains immutable"
    );
    let selected = plan.resolve(0, Resolution::Right)?;
    assert_eq!(plan.unresolved_count(), 1);
    let merged = selected.finish()?;
    package.apply_opened_presentation_patch(&merged)?;
    assert!(
        package
            .presentation()?
            .slide(0)?
            .is_some_and(|slide| slide.text().is_ok_and(|text| text.contains("Right title")))
    );
    assert_eq!(
        package
            .notes()?
            .ok_or_else(|| Error::Invalid("notes disappeared".into()))?
            .slides()[1]
            .text()?
            .as_deref(),
        Some("Merged notes")
    );

    let mut history = History::new(Limits::default());
    history.push(merged.clone())?;
    let undo = history
        .pop_undo()
        .ok_or_else(|| Error::Invalid("undo disappeared".into()))?;
    assert_eq!(history.redo_len(), 1);
    package.apply_opened_presentation_patch(&undo)?;
    let redo = history
        .pop_redo()
        .ok_or_else(|| Error::Invalid("redo disappeared".into()))?;
    package.apply_opened_presentation_patch(&redo)?;
    assert_eq!(history.redo_len(), 0);
    Ok(())
}

#[test]
fn slide_removal_and_dependency_transfer_are_durable_and_reversible() -> Result<()> {
    let source_bytes = std::fs::read("../../test-data/ooxml/pptx/line-chart.pptx")?;
    let source_package = Package::from_bytes(&source_bytes)?;
    let source_root = source_package.opened_presentation()?;
    let source_slide = source_root.slides()[0].part_name().clone();
    let relationship_id = source_package
        .opc
        .get_part(&source_slide)?
        .rels()
        .iter()
        .find(|relationship| {
            crate::parts::is_relationship_type(
                relationship.reltype(),
                litchi_opc::constants::relationship_type::CHART,
                "chart",
            )
        })
        .map(|relationship| relationship.r_id().to_owned())
        .ok_or_else(|| Error::Invalid("real chart fixture has no chart relationship".into()))?;

    let mut destination = opened_two_slide_package()?;
    let destination_before = part_states(&destination);
    let destination_root = destination.opened_presentation()?;
    let mut transfer = destination_root.edit();
    let copied_relationship = transfer.transfer_relationship_closure(
        &source_root,
        &source_slide,
        &relationship_id,
        0_usize,
    )?;
    assert!(!copied_relationship.is_empty());
    let transfer_patch = transfer.commit()?.into_patch();
    destination.apply_opened_presentation_patch(&transfer_patch)?;
    let opened = destination.opened_presentation()?;
    let slide = destination.opc.get_part(opened.slides()[0].part_name())?;
    let related = crate::chart::related(&destination.opc, slide)?;
    assert_eq!(related.len(), 1);
    let copied_chart = related[0].part();
    assert!(!copied_chart.rels().is_empty());
    for relationship in copied_chart
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        assert!(
            destination
                .opc
                .contains_part(&relationship.target_partname()?)
        );
    }
    destination.apply_opened_presentation_patch(&transfer_patch.inverse())?;
    assert_eq!(part_states(&destination), destination_before);

    let mut removal = destination.opened_presentation()?.edit();
    let removed = removal.remove_slide(1_usize)?;
    let removed_part = removed.part_name().clone();
    let removal_patch = removal.commit()?.into_patch();
    destination.apply_opened_presentation_patch(&removal_patch)?;
    assert_eq!(destination.opened_presentation()?.slides().len(), 1);
    assert!(!destination.opc.contains_part(&removed_part));
    destination.apply_opened_presentation_patch(&removal_patch.inverse())?;
    assert_eq!(part_states(&destination), destination_before);
    Ok(())
}

#[test]
fn real_complex_chart_deck_survives_transaction_and_full_reopen() -> Result<()> {
    let bytes = std::fs::read("../../test-data/ooxml/pptx/line-chart.pptx")?;
    let mut package = Package::from_bytes(&bytes)?;
    let before = part_states(&package);
    let source = package.opened_presentation()?;
    let first_slide = source
        .slides()
        .first()
        .ok_or_else(|| Error::Invalid("real chart fixture has no slide".into()))?
        .part_name()
        .clone();
    assert!(!crate::chart::related(&package.opc, package.opc.get_part(&first_slide)?)?.is_empty());

    let mut edit = source.edit();
    edit.add_table(
        0_usize,
        &[
            vec!["Series".into(), "Value".into()],
            vec!["Preserved chart".into(), "42".into()],
        ],
        (100, 100, 2_000_000, 800_000),
    )?;
    edit.add_comment_author(
        crate::comments::Author::new(91, "Fixture Author", "FA"),
        crate::comments::Conformance::Transitional,
    )?;
    edit.add_comment(
        0_usize,
        crate::comments::Comment::new(91, "Fixture transaction", 50, 50),
        crate::comments::Conformance::Transitional,
    )?;
    let patch = edit.commit()?.into_patch();
    let inverse = patch.inverse();
    package.apply_opened_presentation_patch(&patch)?;
    let saved = package.to_bytes()?;
    let mut reopened = Package::from_bytes(&saved)?;
    let reopened_root = reopened.opened_presentation()?;
    let reopened_slide = reopened
        .opc
        .get_part(reopened_root.slides()[0].part_name())?;
    assert!(!crate::chart::related(&reopened.opc, reopened_slide)?.is_empty());
    assert!(
        crate::comments::load_presentation_comments(&reopened.opc)?
            .is_some_and(|comments| comments.slides[0].comments[0].text == "Fixture transaction")
    );
    reopened.apply_opened_presentation_patch(&inverse)?;
    assert_eq!(part_states(&reopened), before);
    Ok(())
}

fn modern_author(id: &str, name: &str) -> crate::modern_comments::Author {
    crate::modern_comments::Author {
        id: id.into(),
        name: name.into(),
        initials: Some("OA".into()),
        user_id: "opened@example.com".into(),
        provider_id: String::new(),
        namespace_declarations: Vec::new(),
        extension_xml: None,
    }
}

fn modern_comment(id: &str, author_id: &str, title: &str) -> crate::modern_comments::Comment {
    crate::modern_comments::Comment {
        id: id.into(),
        author_id: author_id.into(),
        status: Some(crate::modern_comments::Status::Active),
        created: "2026-08-10T10:00:00Z".into(),
        start_date: None,
        due_date: None,
        assigned_to: None,
        complete: None,
        title: Some(title.into()),
        namespace_declarations: Vec::new(),
        anchors: Vec::new(),
        position: Some(crate::modern_comments::Position { x: 10, y: 20 }),
        reply_list_namespace_declarations: Vec::new(),
        replies: Vec::new(),
        reply_list_present: false,
        text_body_xml: None,
        extension_xml: None,
    }
}

fn modern_reply(id: &str, author_id: &str) -> crate::modern_comments::Reply {
    crate::modern_comments::Reply {
        id: id.into(),
        author_id: author_id.into(),
        status: Some(crate::modern_comments::Status::Active),
        created: "2026-08-10T10:01:00Z".into(),
        namespace_declarations: Vec::new(),
        text_body_xml: None,
        extension_xml: None,
    }
}

#[test]
fn general_shapes_picture_parts_and_modern_extensions_are_atomic_and_reversible() -> Result<()> {
    const AUTHOR: &str = "{CD37207E-7903-4ED4-8AE8-017538D2DF7E}";
    const COMMENT: &str = "{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}";
    const REPLY: &str = "{BCA5ED0E-707B-4D89-8B89-62D96F48E871}";
    let mut package = opened_two_slide_package()?;
    let before = part_states(&package);
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    let text_id = edit.add_text_box(
        0_usize,
        "Opened text <safe> & compact",
        (100, 200, 2_000_000, 500_000),
    )?;
    let rectangle_id =
        edit.add_rectangle(0_usize, (100, 800_000, 800_000, 500_000), Some("12ABEF"))?;
    let ellipse_id = edit.add_ellipse(0_usize, (1_000_000, 800_000, 800_000, 500_000), None)?;
    let picture_id = edit.add_picture(
        0_usize,
        "Opened picture",
        &crate::media_parts::Resource::new(
            "/ppt/media/opened-picture.png",
            "image/png",
            vec![137, 80, 78, 71, 13, 10, 26, 10],
        ),
        (2_000_000, 800_000, 800_000, 500_000),
    )?;
    edit.add_modern_comment_author(modern_author(AUTHOR, "Opened Author"))
        .map_err(|error| Error::Invalid(format!("add modern author: {error}")))?;
    assert!(
        edit.replace_modern_comment_author(AUTHOR, modern_author(AUTHOR, "Opened Author Updated"))?
    );
    assert_eq!(
        edit.reorder_modern_comment_authors(&[AUTHOR.to_owned()])?
            .len(),
        1
    );
    edit.add_modern_comment(1_usize, modern_comment(COMMENT, AUTHOR, "Opened task"))
        .map_err(|error| Error::Invalid(format!("add modern comment: {error}")))?;
    assert!(edit.replace_modern_comment(
        1_usize,
        COMMENT,
        modern_comment(COMMENT, AUTHOR, "Opened task updated")
    )?);
    assert_eq!(
        edit.reorder_modern_comments(1_usize, &[COMMENT.to_owned()])?
            .len(),
        1
    );
    assert!(
        edit.add_modern_comment_reply(1_usize, COMMENT, modern_reply(REPLY, AUTHOR))
            .map_err(|error| Error::Invalid(format!("add modern reply: {error}")))?
    );
    let mut replacement_reply = modern_reply(REPLY, AUTHOR);
    replacement_reply.status = Some(crate::modern_comments::Status::Resolved);
    assert!(edit.replace_modern_comment_reply(1_usize, COMMENT, REPLY, replacement_reply)?);
    assert!(
        edit.update_modern_comment_extensions(1_usize, COMMENT, |extensions| {
            extensions
                .replace_reactions(
                    Some("{E1E2D3D4-C5C6-47A8-99AA-BBCCDDEEFF00}"),
                    Some(crate::modern_comments::semantic::reactions::List::default()),
                )
                .expect("test reaction extension is valid");
        })
        .map_err(|error| Error::Invalid(format!("update comment extensions: {error}")))?
    );
    assert!(
        edit.update_modern_comment_reply_extensions(1_usize, COMMENT, REPLY, |extensions| {
            extensions
                .replace_reactions(
                    Some("{10203040-5060-4780-90A0-B0C0D0E0F000}"),
                    Some(crate::modern_comments::semantic::reactions::List::default()),
                )
                .expect("test reply extension is valid");
        })
        .map_err(|error| Error::Invalid(format!("update reply extensions: {error}")))?
    );

    let patch = Patch::from_bytes(
        &edit
            .commit()
            .map_err(|error| Error::Invalid(format!("commit modern extensions: {error}")))?
            .into_patch()
            .to_bytes()?,
    )?;
    assert!(patch.resource_count() >= 5);
    package
        .apply_opened_presentation_patch(&patch)
        .map_err(|error| Error::Invalid(format!("apply modern patch: {error}")))?;
    let bytes = package
        .to_bytes()
        .map_err(|error| Error::Invalid(format!("write modern package: {error}")))?;
    let mut reopened = Package::from_bytes(&bytes)
        .map_err(|error| Error::Invalid(format!("reopen modern package: {error}")))?;
    let opened = reopened.opened_presentation()?;
    let slide = reopened.opc.get_part(opened.slides()[0].part_name())?;
    let scene = crate::shape::Scene::read(slide.blob())?;
    for id in [text_id, rectangle_id, ellipse_id, picture_id] {
        assert!(scene.iter().any(|shape| shape.id() == Some(id)));
    }
    assert!(scene.iter().any(|shape| {
        matches!(shape, crate::shape::Shape::Picture(_)) && shape.id() == Some(picture_id)
    }));
    let image = slide
        .rels()
        .iter()
        .find(|relationship| {
            crate::parts::is_relationship_type(
                relationship.reltype(),
                litchi_opc::constants::relationship_type::IMAGE,
                "image",
            )
        })
        .ok_or_else(|| Error::Invalid("opened picture relationship disappeared".into()))?;
    assert_eq!(
        reopened.opc.get_part(&image.target_partname()?)?.blob()[..4],
        [137, 80, 78, 71]
    );
    let graph = crate::modern_comments::load_modern_comment_graph(&reopened.opc)
        .map_err(|error| Error::Invalid(format!("reload modern graph: {error}")))?;
    assert_eq!(
        graph
            .authors
            .as_ref()
            .map(|part| part.authors.authors.len()),
        Some(1)
    );
    let comment = &graph.comments[0].comments.comments[0];
    assert_eq!(comment.replies.len(), 1);
    assert!(comment.reactions()?.is_some());
    assert!(comment.replies[0].reactions()?.is_some());

    let mut history = History::new(Limits::default());
    history.push(patch.clone())?;
    let undo = history
        .pop_undo()
        .ok_or_else(|| Error::Invalid("shape/comment undo disappeared".into()))?;
    reopened.apply_opened_presentation_patch(&undo)?;
    assert_eq!(part_states(&reopened), before);
    let redo = history
        .pop_redo()
        .ok_or_else(|| Error::Invalid("shape/comment redo disappeared".into()))?;
    reopened.apply_opened_presentation_patch(&redo)?;
    assert!(
        crate::modern_comments::load_modern_comment_graph(&reopened.opc)?
            .authors
            .is_some()
    );
    let with_modern_graph = part_states(&reopened);
    let mut cleanup = reopened.opened_presentation()?.edit();
    assert!(cleanup.remove_modern_comment_reply(1_usize, COMMENT, REPLY)?);
    assert!(cleanup.remove_modern_comment(1_usize, COMMENT)?);
    assert!(cleanup.remove_modern_comment_author(AUTHOR)?);
    let cleanup = cleanup.commit()?.into_patch();
    reopened.apply_opened_presentation_patch(&cleanup)?;
    let empty_graph = crate::modern_comments::load_modern_comment_graph(&reopened.opc)?;
    assert!(empty_graph.authors.is_none());
    assert!(empty_graph.comments.is_empty());
    reopened.apply_opened_presentation_patch(&cleanup.inverse())?;
    assert_eq!(part_states(&reopened), with_modern_graph);
    Ok(())
}

#[test]
fn picture_shape_transfer_copies_relationship_closure_and_remaps_identity() -> Result<()> {
    let mut source_package = opened_two_slide_package()?;
    let source_root = source_package.opened_presentation()?;
    let mut source_edit = source_root.edit();
    let picture_id = source_edit.add_picture(
        0_usize,
        "Transfer picture",
        &crate::media_parts::Resource::new(
            "/ppt/media/transfer-picture.png",
            "image/png",
            vec![137, 80, 78, 71, 1, 2, 3, 4],
        ),
        (100, 100, 900_000, 600_000),
    )?;
    let source_patch = source_edit.commit()?.into_patch();
    source_package.apply_opened_presentation_patch(&source_patch)?;
    let source_bytes = source_package.to_bytes()?;
    let source_package = Package::from_bytes(&source_bytes)?;
    let source_root = source_package.opened_presentation()?;
    let source_slide = source_package
        .opc
        .get_part(source_root.slides()[0].part_name())?;
    let source_position = crate::shape::Scene::read(source_slide.blob())?
        .iter()
        .position(|shape| shape.id() == Some(picture_id))
        .ok_or_else(|| Error::Invalid("source picture shape disappeared".into()))?;

    let mut destination = opened_two_slide_package()?;
    let before = part_states(&destination);
    let destination_root = destination.opened_presentation()?;
    let mut transfer = destination_root.edit();
    let transferred_id =
        transfer.transfer_shape(&source_root, 0_usize, source_position, 1_usize)?;
    let transfer_patch = transfer.commit()?.into_patch();
    destination.apply_opened_presentation_patch(&transfer_patch)?;
    let bytes = destination.to_bytes()?;
    let mut reopened = Package::from_bytes(&bytes)?;
    let opened = reopened.opened_presentation()?;
    let slide = reopened.opc.get_part(opened.slides()[1].part_name())?;
    let transferred_scene = crate::shape::Scene::read(slide.blob())?;
    let transferred = transferred_scene
        .iter()
        .find(|shape| shape.id() == Some(transferred_id))
        .ok_or_else(|| Error::Invalid("transferred picture shape disappeared".into()))?;
    assert!(matches!(transferred, crate::shape::Shape::Picture(_)));
    assert!(transferred.name().is_some_and(|name| name.contains("Copy")));
    let transferred_name = transferred.name().unwrap_or("missing").to_owned();
    let image_target = slide
        .rels()
        .iter()
        .find(|relationship| {
            crate::parts::is_relationship_type(
                relationship.reltype(),
                litchi_opc::constants::relationship_type::IMAGE,
                "image",
            )
        })
        .ok_or_else(|| Error::Invalid("transferred image relationship disappeared".into()))?
        .target_partname()?;
    assert_eq!(
        reopened.opc.get_part(&image_target)?.blob()[..4],
        [137, 80, 78, 71]
    );

    let mut removal = reopened.opened_presentation()?.edit();
    assert!(removal.remove_shape(1_usize, transferred_name.as_str())?);
    let removal_patch = removal.commit()?.into_patch();
    reopened.apply_opened_presentation_patch(&removal_patch)?;
    assert!(!reopened.opc.contains_part(&image_target));
    reopened.apply_opened_presentation_patch(&removal_patch.inverse())?;
    reopened.apply_opened_presentation_patch(&transfer_patch.inverse())?;
    assert_eq!(part_states(&reopened), before);
    Ok(())
}

#[test]
fn grouped_connector_transfer_remaps_identity_closure_and_is_durable() -> Result<()> {
    let mut source_package = opened_two_slide_package()?;
    let authored_source = source_package.opened_presentation()?;
    let mut authored_shapes = authored_source.edit();
    let table_id = authored_shapes.add_table(
        0_usize,
        &[vec!["Transfer".into(), "Table".into()]],
        (0, 0, 1_000_000, 500_000),
    )?;
    let chart = crate::chart::Chart::new(crate::chart::Type::Column, 0, 0, 1_000_000, 1_000_000)
        .with_title("Transfer chart")
        .add_series(
            crate::chart::Series::new("Series")
                .with_categories(vec!["A".into()])
                .with_values(vec![1.0]),
        );
    let _chart_relationship_id = authored_shapes.add_chart(0_usize, &chart)?;
    let authored_patch = authored_shapes.commit()?.into_patch();
    source_package.apply_opened_presentation_patch(&authored_patch)?;
    add_group_connector_transfer_fixture(&mut source_package)?;
    let source_package = Package::from_bytes(&source_package.to_bytes()?)?;
    let source_root = source_package.opened_presentation()?;
    let source_slide = source_package
        .opc
        .get_part(source_root.slides()[0].part_name())?;
    let source_scene = crate::shape::Scene::read(source_slide.blob())?;
    let unknown_position = source_scene
        .iter()
        .position(|shape| matches!(shape, crate::shape::Shape::Unknown(_)))
        .ok_or_else(|| Error::Invalid("source unknown shape disappeared".into()))?;
    let free_connector_position = source_scene
        .iter()
        .position(|shape| shape.name() == Some("Free connector"))
        .ok_or_else(|| Error::Invalid("source free connector disappeared".into()))?;
    let external_connector_position = source_scene
        .iter()
        .position(|shape| shape.name() == Some("External connector"))
        .ok_or_else(|| Error::Invalid("source external connector disappeared".into()))?;
    let unresolved_connector_position = source_scene
        .iter()
        .position(|shape| shape.name() == Some("Unresolved connector"))
        .ok_or_else(|| Error::Invalid("source unresolved connector disappeared".into()))?;
    let content_position = source_scene
        .iter()
        .position(|shape| matches!(shape, crate::shape::Shape::Content(_)))
        .ok_or_else(|| Error::Invalid("source content-part shape disappeared".into()))?;
    let diagram_position = source_scene
        .iter()
        .position(|shape| shape.name() == Some("Dependency diagram"))
        .ok_or_else(|| Error::Invalid("source dependency diagram disappeared".into()))?;
    let frame_position = source_scene
        .iter()
        .position(|shape| shape.name() == Some("Opaque frame"))
        .ok_or_else(|| Error::Invalid("source opaque frame disappeared".into()))?;
    let ole_position = source_scene
        .iter()
        .position(|shape| shape.name() == Some("Inert OLE transfer"))
        .ok_or_else(|| Error::Invalid("source inert OLE shape disappeared".into()))?;
    let nested_position = source_scene
        .iter()
        .position(|shape| shape.name() == Some("Group picture"))
        .ok_or_else(|| Error::Invalid("source nested group picture disappeared".into()))?;
    let table_position = source_scene
        .iter()
        .position(|shape| shape.id() == Some(table_id))
        .ok_or_else(|| Error::Invalid("source table disappeared".into()))?;
    let chart_position = source_scene
        .iter()
        .position(|shape| matches!(shape, crate::shape::Shape::Chart(_)))
        .ok_or_else(|| Error::Invalid("source chart disappeared".into()))?;

    let mut destination = opened_two_slide_package()?;
    let before = part_states(&destination);
    let destination_root = destination.opened_presentation()?;
    let mut refusal = destination_root.edit();
    let refusal_error = refusal
        .transfer_shape(&source_root, 0_usize, unknown_position, 1_usize)
        .expect_err("unknown shape transfer must be refused");
    assert!(matches!(
        refusal_error,
        Error::ShapeTransfer {
            kind: crate::ShapeTransferRefusal::UnknownExtensionShape,
            ..
        }
    ));
    assert!(!refusal.is_changed());
    let mut unresolved_refusal = destination_root.edit();
    let unresolved_error = unresolved_refusal
        .transfer_shape(
            &source_root,
            0_usize,
            unresolved_connector_position,
            1_usize,
        )
        .expect_err("unresolved connector endpoint transfer must be refused");
    assert!(matches!(
        unresolved_error,
        Error::ShapeTransfer {
            kind: crate::ShapeTransferRefusal::UnresolvedConnectorEndpoint,
            ..
        }
    ));
    assert!(!unresolved_refusal.is_changed());
    let mut content_refusal = destination_root.edit();
    let content_error = content_refusal
        .transfer_shape(&source_root, 0_usize, content_position, 1_usize)
        .expect_err("content-part shape transfer must be classified");
    assert!(matches!(
        content_error,
        Error::ShapeTransfer {
            kind: crate::ShapeTransferRefusal::ContentPart,
            ..
        }
    ));
    assert!(!content_refusal.is_changed());
    let mut frame_refusal = destination_root.edit();
    let frame_error = frame_refusal
        .transfer_shape(&source_root, 0_usize, frame_position, 1_usize)
        .expect_err("unclassified frame transfer must be classified");
    assert!(matches!(
        frame_error,
        Error::ShapeTransfer {
            kind: crate::ShapeTransferRefusal::UnclassifiedGraphicFrame,
            ..
        }
    ));
    assert!(!frame_refusal.is_changed());
    let mut nested_refusal = destination_root.edit();
    let nested_error = nested_refusal
        .transfer_shape(&source_root, 0_usize, nested_position, 1_usize)
        .expect_err("nested shape transfer must identify its owner boundary");
    assert!(matches!(
        nested_error,
        Error::ShapeTransfer {
            kind: crate::ShapeTransferRefusal::NestedShape,
            ..
        }
    ));
    assert!(!nested_refusal.is_changed());

    let mut transfer = destination_root.edit();
    let transferred_external_connector_id =
        transfer.transfer_shape(&source_root, 0_usize, external_connector_position, 1_usize)?;
    let transferred_free_connector_id =
        transfer.transfer_shape(&source_root, 0_usize, free_connector_position, 1_usize)?;
    let transferred_diagram_id =
        transfer.transfer_shape(&source_root, 0_usize, diagram_position, 1_usize)?;
    let transferred_ole_id =
        transfer.transfer_shape(&source_root, 0_usize, ole_position, 1_usize)?;
    let transferred_table_id =
        transfer.transfer_shape(&source_root, 0_usize, table_position, 1_usize)?;
    let transferred_chart_id =
        transfer.transfer_shape(&source_root, 0_usize, chart_position, 1_usize)?;
    let transfer_patch = transfer.commit()?.into_patch();
    let transfer_patch = Patch::from_bytes(&transfer_patch.to_bytes()?)?;
    let mut notes = destination_root.edit();
    notes.set_notes_text(0_usize, "Disjoint grouped-transfer note")?;
    let notes_patch = notes.commit()?.into_patch();
    assert!(!transfer_patch.conflicts_with(&notes_patch));
    let merged = Patch::three_way(&destination_root, &transfer_patch, &notes_patch)?.finish()?;
    let durable = Patch::from_bytes(&merged.to_bytes()?)?;
    destination.apply_opened_presentation_patch(&durable)?;

    let mut reopened = Package::from_bytes(&destination.to_bytes()?)?;
    let opened = reopened.opened_presentation()?;
    let destination_slide = reopened.opc.get_part(opened.slides()[1].part_name())?;
    assert!(!destination_slide.blob().contains(&b'\n'));
    let scene = crate::shape::Scene::read(destination_slide.blob())?;
    let transferred_group = scene
        .roots()
        .find(|shape| {
            matches!(shape, crate::shape::Shape::Group(_))
                && shape
                    .name()
                    .is_some_and(|name| name.starts_with("Transfer group Copy "))
        })
        .ok_or_else(|| Error::Invalid("connector endpoint group closure disappeared".into()))?;
    let transferred_group_id = transferred_group
        .id()
        .ok_or_else(|| Error::Invalid("transferred group has no identity".into()))?;
    let external_connector = scene
        .roots()
        .find(|shape| shape.id() == Some(transferred_external_connector_id))
        .ok_or_else(|| Error::Invalid("transferred external connector disappeared".into()))?;
    assert!(matches!(
        external_connector,
        crate::shape::Shape::Connector(_)
    ));
    let free_connector = scene
        .roots()
        .find(|shape| shape.id() == Some(transferred_free_connector_id))
        .ok_or_else(|| Error::Invalid("transferred free connector disappeared".into()))?;
    assert!(matches!(free_connector, crate::shape::Shape::Connector(_)));
    assert!(
        free_connector
            .name()
            .is_some_and(|name| name.contains("Copy"))
    );
    let transferred_diagram = scene
        .roots()
        .find(|shape| shape.id() == Some(transferred_diagram_id))
        .ok_or_else(|| Error::Invalid("transferred diagram disappeared".into()))?;
    assert!(matches!(
        transferred_diagram,
        crate::shape::Shape::Diagram(_)
    ));
    let diagram_xml = std::str::from_utf8(transferred_diagram.xml()?)
        .map_err(|error| Error::Xml(error.to_string()))?;
    for old_relationship in [
        "rIdDiagramData",
        "rIdDiagramLayout",
        "rIdDiagramStyle",
        "rIdDiagramColors",
    ] {
        assert!(!diagram_xml.contains(old_relationship));
    }
    for attribute in ["rel:dm=", "rel:lo=", "rel:qs=", "rel:cs="] {
        assert!(diagram_xml.contains(attribute));
    }
    let transferred_ole = scene
        .roots()
        .find(|shape| shape.id() == Some(transferred_ole_id))
        .ok_or_else(|| Error::Invalid("transferred inert OLE shape disappeared".into()))?;
    assert!(matches!(transferred_ole, crate::shape::Shape::Ole(_)));
    let ole_xml = std::str::from_utf8(transferred_ole.xml()?)
        .map_err(|error| Error::Xml(error.to_string()))?;
    assert!(!ole_xml.contains("r:id=\"rIdOleTransfer\""));
    assert!(ole_xml.contains("x:provenance=\"rIdOleTransfer\""));
    assert!(scene.roots().any(|shape| {
        shape.id() == Some(transferred_table_id) && matches!(shape, crate::shape::Shape::Table(_))
    }));
    assert!(scene.roots().any(|shape| {
        shape.id() == Some(transferred_chart_id) && matches!(shape, crate::shape::Shape::Chart(_))
    }));
    let transferred = scene
        .iter()
        .find(|shape| shape.id() == Some(transferred_group_id))
        .ok_or_else(|| Error::Invalid("transferred group disappeared".into()))?;
    let crate::shape::Shape::Group(group) = transferred else {
        return Err(Error::Invalid(
            "transferred grouped shape changed kind".into(),
        ));
    };
    let children: Vec<_> = group.shapes().collect();
    assert_eq!(children.len(), 3);
    let picture = children
        .iter()
        .copied()
        .find(|shape| matches!(shape, crate::shape::Shape::Picture(_)))
        .ok_or_else(|| Error::Invalid("transferred group picture disappeared".into()))?;
    let target = children
        .iter()
        .copied()
        .find(|shape| matches!(shape, crate::shape::Shape::Auto(_)))
        .ok_or_else(|| Error::Invalid("transferred group target disappeared".into()))?;
    let connector = children
        .iter()
        .copied()
        .find(|shape| matches!(shape, crate::shape::Shape::Connector(_)))
        .ok_or_else(|| Error::Invalid("transferred group connector disappeared".into()))?;
    let picture_id = picture
        .id()
        .ok_or_else(|| Error::Invalid("transferred group picture has no identity".into()))?;
    let target_id = target
        .id()
        .ok_or_else(|| Error::Invalid("transferred group target has no identity".into()))?;
    let connector_id = connector
        .id()
        .ok_or_else(|| Error::Invalid("transferred group connector has no identity".into()))?;
    let transferred_ids = [transferred_group_id, picture_id, target_id, connector_id];
    assert_eq!(
        transferred_ids
            .into_iter()
            .collect::<std::collections::HashSet<_>>()
            .len(),
        4
    );
    assert!(transferred_ids.iter().all(|id| !(20..=23).contains(id)));
    let connector_xml =
        std::str::from_utf8(connector.xml()?).map_err(|error| Error::Xml(error.to_string()))?;
    assert!(connector_xml.contains(&format!("<a:stCxn id=\"{picture_id}\" idx=\"0\"/>")));
    assert!(connector_xml.contains(&format!("<a:endCxn id=\"{target_id}\" idx=\"0\"/>")));
    let external_connector_xml = std::str::from_utf8(external_connector.xml()?)
        .map_err(|error| Error::Xml(error.to_string()))?;
    assert!(external_connector_xml.contains(&format!(
        "<a:stCxn id=\"{transferred_group_id}\" idx=\"0\"/>"
    )));
    assert!(external_connector_xml.contains(&format!("<a:endCxn id=\"{target_id}\" idx=\"0\"/>")));
    assert!(picture.name().is_some_and(|name| name.contains("Copy")));
    assert!(connector.name().is_some_and(|name| name.contains("Copy")));

    let mut copied_image = None;
    for relationship in destination_slide.rels().iter() {
        if !crate::parts::is_relationship_type(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::IMAGE,
            "image",
        ) {
            continue;
        }
        let target_name = relationship.target_partname()?;
        if reopened.opc.get_part(&target_name)?.blob() == [137, 80, 78, 71, 9, 8, 7, 6] {
            copied_image = Some(target_name);
            break;
        }
    }
    assert!(copied_image.is_some());
    let mut copied_diagram_types = std::collections::BTreeSet::new();
    for relationship in destination_slide.rels().iter() {
        if relationship.is_external() {
            continue;
        }
        let target_name = relationship.target_partname()?;
        let part = reopened.opc.get_part(&target_name)?;
        if part
            .content_type()
            .starts_with("application/vnd.openxmlformats-officedocument.drawingml.diagram")
        {
            copied_diagram_types.insert(part.content_type().to_owned());
        }
    }
    assert_eq!(
        copied_diagram_types,
        [
            litchi_opc::constants::content_type::DML_DIAGRAM_COLORS.to_owned(),
            litchi_opc::constants::content_type::DML_DIAGRAM_DATA.to_owned(),
            litchi_opc::constants::content_type::DML_DIAGRAM_LAYOUT.to_owned(),
            litchi_opc::constants::content_type::DML_DIAGRAM_STYLE.to_owned(),
        ]
        .into_iter()
        .collect()
    );
    let mut copied_ole = false;
    for relationship in destination_slide.rels().iter() {
        if relationship.is_external() {
            continue;
        }
        let target_name = relationship.target_partname()?;
        let part = reopened.opc.get_part(&target_name)?;
        if part.content_type() == litchi_opc::constants::content_type::OFC_OLE_OBJECT
            && part.blob() == [0xD0, 0xCF, 0x11, 0xE0, 1, 2, 3, 4]
        {
            copied_ole = true;
            break;
        }
    }
    assert!(copied_ole);
    assert_eq!(
        crate::chart::related(&reopened.opc, destination_slide)?.len(),
        1
    );
    assert_eq!(
        reopened
            .notes()?
            .ok_or_else(|| Error::Invalid("reopened notes disappeared".into()))?
            .slides()[0]
            .text()?
            .as_deref(),
        Some("Disjoint grouped-transfer note")
    );

    let after = part_states(&reopened);
    let mut history = History::new(Limits::default());
    history.push(durable.clone())?;
    let undo = history
        .pop_undo()
        .ok_or_else(|| Error::Invalid("group transfer undo disappeared".into()))?;
    assert_eq!(undo, durable.inverse());
    reopened.apply_opened_presentation_patch(&undo)?;
    assert_eq!(part_states(&reopened), before);
    let redo = history
        .pop_redo()
        .ok_or_else(|| Error::Invalid("group transfer redo disappeared".into()))?;
    reopened.apply_opened_presentation_patch(&redo)?;
    assert_eq!(part_states(&reopened), after);
    Ok(())
}

#[test]
fn disjoint_new_shape_and_modern_comment_changes_merge_automatically() -> Result<()> {
    const AUTHOR: &str = "{0B2043D4-0908-4C42-8A79-51EA2CC309F7}";
    const COMMENT: &str = "{8F23E89D-A0D4-4AB0-AF88-6DEEA19D812A}";
    let mut package = opened_two_slide_package()?;
    let base = package.opened_presentation()?;
    let mut shape = base.edit();
    shape.add_rectangle(0_usize, (10, 10, 500_000, 500_000), Some("00AAFF"))?;
    let shape = shape.commit()?.into_patch();
    let mut modern = base.edit();
    modern.add_modern_comment_author(modern_author(AUTHOR, "Merge Author"))?;
    modern.add_modern_comment(1_usize, modern_comment(COMMENT, AUTHOR, "Merge comment"))?;
    let modern = modern.commit()?.into_patch();
    assert!(!shape.conflicts_with(&modern));
    let merged = Patch::three_way(&base, &shape, &modern)?.finish()?;
    package.apply_opened_presentation_patch(&merged)?;
    let reopened = Package::from_bytes(&package.to_bytes()?)?;
    assert!(
        crate::modern_comments::load_modern_comment_graph(&reopened.opc)?
            .authors
            .is_some()
    );
    assert!(
        crate::shape::Scene::read(
            reopened
                .opc
                .get_part(reopened.opened_presentation()?.slides()[0].part_name())?
                .blob()
        )?
        .iter()
        .any(|shape| shape
            .name()
            .is_some_and(|name| name.starts_with("Rectangle ")))
    );
    Ok(())
}
