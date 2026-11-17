use std::collections::BTreeMap;

use base64::Engine as _;
use litchi_equation::Limits;
use litchi_opc::PackURI;

use super::{Source, markup};
use crate::Package;

thread_local! {
    /// The `Equation Native` bytes decoded on this thread.
    pub(super) static DECODED: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

/// A workbook whose first sheet holds one written object and `count`
/// more, each with its own relationship, naming the embedding parts
/// `oleObject1.bin` to `oleObject{parts}.bin` in turn; each part holds an
/// `Equation Native` stream of `native` bytes.
fn shared_embeddings(count: u32, parts: u32, native: usize) -> Package {
    let (package, _) = add(&Package::new().unwrap(), 0, 0, Source::Latex("x"));
    let mut opc = package.opc().clone();
    let mut stream = vec![0_u8; native];
    stream[..2].copy_from_slice(&28_u16.to_le_bytes());
    let object =
        litchi_equation::equation_object(&litchi_equation::MtefTarget::MathType, &stream).unwrap();
    for part in 1..=parts {
        let name = PackURI::new(format!("/xl/embeddings/oleObject{part}.bin")).unwrap();
        if opc.get_part(&name).is_ok() {
            opc.get_part_mut(&name).unwrap().set_blob(object.clone());
        } else {
            opc.add_part(Box::new(litchi_opc::BlobPart::new(
                name,
                "application/vnd.openxmlformats-officedocument.oleObject".to_owned(),
                object.clone(),
            )));
        }
    }
    let sheet = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let mut entries = String::new();
    for index in 0..count {
        let id = format!("rIdShared{index}");
        let target = format!("../embeddings/oleObject{}.bin", index % parts + 1);
        opc.get_part_mut(&sheet)
            .unwrap()
            .rels_mut()
            .add_relationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
                    .to_owned(),
                target,
                id.clone(),
                false,
            );
        entries.push_str(&format!(
            "<oleObject xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" progId=\"Equation.DSMT4\" shapeId=\"{}\" r:id=\"{id}\"/>",
            3000 + index
        ));
    }
    let xml = text(&package, "/xl/worksheets/sheet1.xml")
        .replace("</oleObjects>", &format!("{entries}</oleObjects>"));
    opc.get_part_mut(&sheet).unwrap().set_blob(xml.into_bytes());
    Package::from_opc(opc).unwrap()
}

#[test]
fn objects_sharing_an_embedding_decode_it_once_and_only_when_read() {
    const NATIVE: usize = 2 * 1024 * 1024;
    let package = shared_embeddings(200, 1, NATIVE);
    DECODED.with(|decoded| decoded.set(0));
    let objects = package.equation_objects(0).unwrap();
    assert_eq!(objects.len(), 201);
    // Listing reads no stream.
    assert_eq!(DECODED.with(std::cell::Cell::get), 0);
    // One object's stream, then another's from the same part: one decode.
    assert_eq!(objects[1].native().unwrap().len(), NATIVE);
    assert_eq!(objects[150].native().unwrap().len(), NATIVE);
    for object in &objects {
        assert!(object.error().is_none());
    }
    assert_eq!(DECODED.with(std::cell::Cell::get), NATIVE);
    // The LaTeX for queries is converted from the same shared stream (which
    // holds no equation, so there is none).
    for object in &objects {
        assert_eq!(object.default_latex(), None);
    }
    assert_eq!(DECODED.with(std::cell::Cell::get), NATIVE);
}

#[test]
fn a_spent_listing_budget_is_a_per_object_error() {
    const NATIVE: usize = 1024 * 1024;
    let package = shared_embeddings(4, 2, NATIVE);
    let sheet = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let embedding = package
        .opc()
        .get_part(&PackURI::new("/xl/embeddings/oleObject1.bin").unwrap())
        .unwrap()
        .blob()
        .len() as u64;
    // Room for one part and its stream, not for the second part.
    let budget = embedding + NATIVE as u64 + embedding / 2;
    let objects = super::list_within(package.opc(), &sheet, budget).unwrap();
    assert_eq!(objects.len(), 5);
    // The extra objects name oleObject1, then oleObject2, in turn.
    assert!(objects[1].native().is_some());
    let error = objects[2].error().unwrap();
    assert!(error.contains("byte budget"), "{error}");
    assert_eq!(objects[2].default_latex(), None);
    // The first part stays readable; the listing itself did not fail.
    assert!(objects[3].native().is_some());
    assert!(objects[4].native().is_none());
}

/// A new workbook with a second sheet, `Sheet2`; when `vml` is given, that
/// sheet's legacy drawing is `/xl/drawings/vmlDrawing9.vml` holding `vml`.
fn two_sheets(vml: Option<&str>) -> Package {
    use litchi_opc::BlobPart;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};

    let package = Package::new().unwrap();
    let mut opc = package.opc().clone();
    let first = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let second = PackURI::new("/xl/worksheets/sheet2.xml").unwrap();
    let mut sheet = text(&package, "/xl/worksheets/sheet1.xml");
    if vml.is_some() {
        sheet = sheet.replace(
            "</worksheet>",
            "<legacyDrawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdVml\"/></worksheet>",
        );
        assert!(sheet.contains("legacyDrawing"));
    }
    opc.add_part(Box::new(BlobPart::new(
        second.clone(),
        opc.get_part(&first).unwrap().content_type().to_owned(),
        sheet.into_bytes(),
    )));
    if let Some(vml) = vml {
        opc.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/drawings/vmlDrawing9.vml").unwrap(),
            ct::OFC_VML_DRAWING.to_owned(),
            vml.as_bytes().to_vec(),
        )));
        opc.get_part_mut(&second)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::VML_DRAWING.to_owned(),
                "../drawings/vmlDrawing9.vml".to_owned(),
                "rIdVml".to_owned(),
                false,
            );
    }
    let workbook = PackURI::new("/xl/workbook.xml").unwrap();
    let xml = text(&package, "/xl/workbook.xml").replace(
        "</sheets>",
        r#"<sheet name="Sheet2" sheetId="2" r:id="rIdSheet2"/></sheets>"#,
    );
    let part = opc.get_part_mut(&workbook).unwrap();
    part.set_blob(xml.into_bytes());
    part.rels_mut().add_relationship(
        rt::WORKSHEET.to_owned(),
        "worksheets/sheet2.xml".to_owned(),
        "rIdSheet2".to_owned(),
        false,
    );
    Package::from_opc(opc).unwrap()
}

#[test]
fn a_new_vml_drawing_takes_a_shape_id_block_no_other_sheet_claims() {
    // A second sheet whose VML drawing claims block 1 (shape ids 1025...),
    // as Excel numbers blocks in the order drawings are created.
    let package = two_sheets(Some(&markup::new_vml_drawing(
        1,
        "<v:shape id=\"_x0000_s1025\" type=\"#_x0000_t75\"><x:ClientData ObjectType=\"Pict\"/></v:shape>",
    )));

    // The first sheet's new VML drawing takes block 2, not block 1.
    let (after, shape) = add(&package, 0, 0, Source::Latex("x"));
    assert_eq!(shape, 2049);
    let drawing = text(&after, "/xl/drawings/vmlDrawing1.vml");
    assert!(drawing.contains("<o:idmap v:ext=\"edit\" data=\"2\"/>"));
    assert!(drawing.contains("id=\"_x0000_s2049\""));
    assert_eq!(after.equation_objects(0).unwrap()[0].shape_id(), 2049);
    assert_eq!(
        crate::vml_ids::idmap_blocks(text(&after, "/xl/drawings/vmlDrawing9.vml").as_bytes()),
        [1]
    );
    // A note on that sheet takes its block too, past the object.
    let after = note(&after, 0, "C5");
    assert_vml_ids_unique(&after);
}

/// `package` with a note at `cell` of the sheet at `sheet`.
fn note(package: &Package, sheet: usize, cell: &str) -> Package {
    let mut edit = package.edit_comments().unwrap();
    edit.set(sheet, cell, "Ann", "Check this").unwrap();
    edit.commit().unwrap().into_package()
}

/// Every VML shape id of `package` is unique, lies in a block its own
/// drawing claims, and no two drawings claim one block; every equation
/// object's shape id is a VML shape's.
fn assert_vml_ids_unique(package: &Package) -> Vec<u32> {
    let mut ids = Vec::new();
    let mut blocks = Vec::new();
    for part in package.opc().iter_parts() {
        if part.content_type() != "application/vnd.openxmlformats-officedocument.vmlDrawing" {
            continue;
        }
        let xml = text(package, part.partname().as_str());
        let claimed = crate::vml_ids::idmap_blocks(xml.as_bytes());
        for block in &claimed {
            assert!(!blocks.contains(block), "block {block} claimed twice");
        }
        blocks.extend(claimed.iter().copied());
        for (at, _) in xml.match_indices("<v:shape ") {
            let id = crate::vml_ids::shape_id(&xml.as_bytes()[at..]).unwrap();
            assert!(
                claimed.contains(&((id - 1) / 1024)),
                "{id} outside {claimed:?}"
            );
            assert!(!ids.contains(&id), "shape id {id} used twice");
            ids.push(id);
        }
    }
    for sheet in 0..package.workbook_ref().len() {
        for object in package.equation_objects(sheet).unwrap() {
            assert!(ids.contains(&object.shape_id()));
        }
    }
    ids
}

#[test]
fn notes_and_equation_objects_share_one_drawing_without_duplicate_ids() {
    let blank = Package::new().unwrap();
    // A note after an object.
    let (one, _) = add(&blank, 0, 0, Source::Latex("x"));
    let one = note(&one, 0, "B2");
    assert_eq!(assert_vml_ids_unique(&one).len(), 2);
    // An object after a note.
    let two = note(&blank, 0, "B2");
    let (two, shape) = add(&two, 0, 3, Source::Latex("y"));
    assert_eq!(shape, 1026);
    assert_eq!(assert_vml_ids_unique(&two).len(), 2);
    // Interleaved.
    let (three, _) = add(&blank, 0, 0, Source::Latex("a"));
    let three = note(&three, 0, "B2");
    let (three, _) = add(&three, 4, 0, Source::Latex("b"));
    let three = note(&three, 0, "D9");
    assert_eq!(assert_vml_ids_unique(&three).len(), 4);
    assert_eq!(three.equation_objects(0).unwrap().len(), 2);
}

#[test]
fn notes_and_equation_objects_on_two_sheets_claim_distinct_blocks() {
    let package = two_sheets(None);
    // A note on the second sheet, an object on the first, then the other
    // way round.
    let package = note(&package, 1, "B2");
    let mut edit = package.edit_equation_objects().unwrap();
    edit.add(0, 0, 0, Source::Latex("x"), &limits()).unwrap();
    edit.add(1, 5, 5, Source::Latex("y"), &limits()).unwrap();
    let package = edit.commit().unwrap().into_package();
    let package = note(&package, 0, "C3");
    assert_eq!(assert_vml_ids_unique(&package).len(), 4);

    // Objects first, on the second sheet, then notes.
    let package = two_sheets(None);
    let mut edit = package.edit_equation_objects().unwrap();
    edit.add(1, 0, 0, Source::Latex("x"), &limits()).unwrap();
    let package = edit.commit().unwrap().into_package();
    let package = note(&package, 0, "B2");
    let package = note(&package, 1, "D4");
    let mut edit = package.edit_equation_objects().unwrap();
    edit.add(0, 2, 2, Source::Latex("z"), &limits()).unwrap();
    let package = edit.commit().unwrap().into_package();
    assert_eq!(assert_vml_ids_unique(&package).len(), 4);
}

#[test]
fn idmap_blocks_read_every_claimed_block() {
    assert_eq!(
        crate::vml_ids::idmap_blocks(b"<xml><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1,3\"/></o:shapelayout></xml>"),
        [1, 3]
    );
    assert!(crate::vml_ids::idmap_blocks(b"<xml/>").is_empty());
}

#[test]
fn listing_budget_takes_the_smaller_of_the_package_and_module_limits() {
    let budget = super::Budget::new(u64::MAX);
    assert!(budget.charge(super::MAX_LISTING_NATIVE_BYTES).is_ok());
    assert!(budget.charge(1).is_err());
    let budget = super::Budget::new(10);
    assert!(budget.charge(11).is_err());
    assert!(budget.charge(10).is_ok());
}

const POI: &[u8] =
    include_bytes!("../../../../test-data/poi/test-data/spreadsheet/bug64512_embed.xlsx");
const EXCEL: &[u8] =
    include_bytes!("../../../../test-data/microsoft-office/excel/excel-links-objects.xlsx");

/// An Equation Editor 3.0 `Equation Native` stream holding the digit 2.
fn equation3() -> Vec<u8> {
    base64::engine::general_purpose::STANDARD
        .decode("HAAAAAIAbcENAAAAAAAAAAAAAAAAAAAAAAAAAAMBAQMACgECiDIAAAA=")
        .unwrap()
}

/// LaTeX without the spaces and braces the MTEF reader adds.
fn norm(latex: &str) -> String {
    latex.replace([' ', '{', '}'], "")
}

fn limits() -> Limits {
    Limits::DEFAULT
}

/// Every part of a package: content type, bytes and relationships.
fn parts(package: &Package) -> BTreeMap<String, (String, Vec<u8>, Vec<(String, String, String)>)> {
    package
        .opc()
        .iter_parts()
        .map(|part| {
            let mut rels = part
                .rels()
                .iter()
                .map(|r| {
                    (
                        r.r_id().to_owned(),
                        r.reltype().to_owned(),
                        r.target_ref().to_owned(),
                    )
                })
                .collect::<Vec<_>>();
            rels.sort();
            (
                part.partname().to_string(),
                (
                    part.content_type().to_owned(),
                    package
                        .opc()
                        .get_part(part.partname())
                        .unwrap()
                        .blob()
                        .to_vec(),
                    rels,
                ),
            )
        })
        .collect()
}

fn text(package: &Package, name: &str) -> String {
    let name = PackURI::new(name).unwrap();
    String::from_utf8(package.opc().get_part(&name).unwrap().blob().to_vec()).unwrap()
}

fn add(package: &Package, row: u32, column: u32, source: Source<'_>) -> (Package, u32) {
    let mut edit = package.edit_equation_objects().unwrap();
    let written = edit.add(0, row, column, source, &limits()).unwrap();
    (edit.commit().unwrap().into_package(), written.shape_id())
}

#[test]
fn a_new_object_is_written_as_excel_writes_one() {
    let package = Package::new().unwrap();
    assert!(package.equation_objects(0).unwrap().is_empty());
    let (after, shape) = add(&package, 2, 1, Source::Latex(r"\frac{a}{b}"));
    assert_eq!(shape, 1025);
    let objects = after.equation_objects(0).unwrap();
    assert_eq!(objects.len(), 1);
    let object = &objects[0];
    assert_eq!(object.prog_id(), "Equation.DSMT4");
    assert_eq!(object.shape_id(), 1025);
    assert_eq!(object.error(), None);
    assert_eq!(object.latex(&limits()).unwrap(), r"\frac{a}{b}");
    let info = litchi_equation::read_mtef(object.native().unwrap(), &limits()).unwrap();
    assert_eq!(info.info().unwrap().version, 5);
    let anchor = object.anchor().unwrap();
    assert_eq!((anchor.from.column, anchor.from.row), (1, 2));
    assert_eq!(object.embedding(), Some("/xl/embeddings/oleObject1.bin"));
    assert_eq!(object.preview(), Some("/xl/media/image1.png"));

    let sheet = text(&after, "/xl/worksheets/sheet1.xml");
    assert!(sheet.contains(
        "<mc:Choice Requires=\"x14\"><oleObject progId=\"Equation.DSMT4\" shapeId=\"1025\" r:id=\""
    ));
    assert!(sheet.contains("<objectPr defaultSize=\"0\" r:id=\""));
    assert!(sheet.contains("<mc:Fallback><oleObject progId=\"Equation.DSMT4\" shapeId=\"1025\""));
    assert!(sheet.contains("<drawing "));
    assert!(sheet.contains("<legacyDrawing "));
    let vml = text(&after, "/xl/drawings/vmlDrawing1.vml");
    assert!(vml.contains("id=\"_x0000_s1025\" type=\"#_x0000_t75\""));
    assert!(vml.contains("<x:ClientData ObjectType=\"Pict\">"));
    assert!(vml.contains("<o:idmap v:ext=\"edit\" data=\"1\"/>"));
    let drawing = text(&after, "/xl/drawings/drawing1.xml");
    assert!(drawing.contains("<a14:compatExt spid=\"_x0000_s1025\"/>"));
    assert!(drawing.contains("<xdr:cNvPr id=\"1025\" name=\"Object 1\" hidden=\"1\">"));
    let embedding = after
        .opc()
        .get_part(&PackURI::new("/xl/embeddings/oleObject1.bin").unwrap())
        .unwrap();
    assert_eq!(
        embedding.content_type(),
        "application/vnd.openxmlformats-officedocument.oleObject"
    );
    let image = after
        .opc()
        .get_part(&PackURI::new("/xl/media/image1.png").unwrap())
        .unwrap();
    assert_eq!(image.content_type(), "image/png");
    assert!(image.blob().starts_with(b"\x89PNG"));

    // The strict OLE-object owner reads the new graph back.
    let mut ole = after.edit_ole_objects(0usize).unwrap();
    let objects = ole.objects().unwrap().unwrap();
    assert_eq!(objects.objects.len(), 1);
    assert_eq!(
        objects.objects[0].program_id.as_deref(),
        Some("Equation.DSMT4")
    );
    // The new drawing shape is not a picture or a shape.
    assert!(after.drawing_elements(0).unwrap().iter().all(|element| {
        !matches!(
            element.kind,
            crate::drawing::edit::ElementKind::Picture | crate::drawing::edit::ElementKind::Shape
        )
    }));
    // The saved package opens again with the object.
    let reopened = Package::from_bytes(&after.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.equation_objects(0).unwrap().len(), 1);
}

#[test]
fn add_set_and_remove_keep_every_other_byte() {
    let package = Package::new().unwrap();
    let (one, first) = add(&package, 0, 0, Source::Latex("x^2"));
    let (two, second) = add(&one, 5, 3, Source::Mtef(&equation3()));
    assert_eq!(second, 1026);
    let objects = two.equation_objects(0).unwrap();
    assert_eq!(objects.len(), 2);
    assert_eq!(objects[1].prog_id(), "Equation.3");
    assert_eq!(objects[1].latex(&limits()).unwrap(), "2");

    // Set the second: the first object's parts keep their bytes.
    let before = parts(&two);
    let mut edit = two.edit_equation_objects().unwrap();
    edit.set(0, second, Source::Latex(r"\sqrt{y}"), &limits())
        .unwrap();
    let three = edit.commit().unwrap().into_package();
    let after = parts(&three);
    for name in ["/xl/embeddings/oleObject1.bin", "/xl/media/image1.png"] {
        assert_eq!(before[name], after[name], "{name}");
    }
    assert_ne!(
        before["/xl/embeddings/oleObject2.bin"].1,
        after["/xl/embeddings/oleObject2.bin"].1
    );
    assert_ne!(
        before["/xl/media/image2.png"].1,
        after["/xl/media/image2.png"].1
    );
    let objects = three.equation_objects(0).unwrap();
    assert_eq!(objects[1].prog_id(), "Equation.3");
    assert_eq!(objects[1].latex(&limits()).unwrap(), r"\sqrt{y}");
    let info = litchi_equation::read_mtef(objects[1].native().unwrap(), &limits()).unwrap();
    assert_eq!(info.info().unwrap().version, 3);
    assert_eq!(norm(&objects[0].latex(&limits()).unwrap()), "x^2");

    // MTEF of the other kind is refused.
    let mut edit = three.edit_equation_objects().unwrap();
    let mathtype = litchi_equation::latex_to_equation_native(
        "z",
        &litchi_equation::MtefTarget::MathType,
        &limits(),
    )
    .unwrap();
    let error = edit
        .set(0, second, Source::Mtef(mathtype.bytes()), &limits())
        .unwrap_err();
    assert!(matches!(error, crate::Error::Refused(_)), "{error}");
    edit.set(0, first, Source::Mtef(mathtype.bytes()), &limits())
        .unwrap();

    // Remove the first, then the second: the workbook is the new one again.
    let mut edit = three.edit_equation_objects().unwrap();
    edit.remove(0, first).unwrap();
    let four = edit.commit().unwrap().into_package();
    let objects = four.equation_objects(0).unwrap();
    assert_eq!(objects.len(), 1);
    assert_eq!(objects[0].shape_id(), second);
    let mut edit = four.edit_equation_objects().unwrap();
    edit.remove(0, second).unwrap();
    let five = edit.commit().unwrap().into_package();
    assert_eq!(parts(&five), parts(&package));
    // No content-type default or override of the removed parts survives.
    assert_eq!(five.to_bytes().unwrap(), package.to_bytes().unwrap());
}

#[test]
fn objects_join_and_leave_excel_embedded_objects_byte_for_byte() {
    for bytes in [POI, EXCEL] {
        let package = Package::from_bytes(bytes).unwrap();
        assert!(package.equation_objects(0).unwrap().is_empty());
        let before = parts(&package);
        let (after, shape) = add(&package, 30, 2, Source::Latex("E=mc^2"));
        let objects = after.equation_objects(0).unwrap();
        assert_eq!(objects.len(), 1);
        // The existing objects keep their entries; the new one follows.
        let mut ole = after.edit_ole_objects(0usize).unwrap();
        let count = ole.objects().unwrap().unwrap().objects.len();
        assert_eq!(count, if bytes == POI { 3 } else { 2 });
        // Parts the edit does not own keep their bytes.
        let changed = parts(&after);
        for (name, value) in &before {
            if !name.starts_with("/xl/worksheets/") && !name.starts_with("/xl/drawings/") {
                assert_eq!(&changed[name], value, "{name}");
            }
        }
        let mut edit = after.edit_equation_objects().unwrap();
        edit.remove(0, shape).unwrap();
        let removed = edit.commit().unwrap().into_package();
        assert_eq!(parts(&removed), before);
    }
}

#[test]
fn a_shared_emf_preview_is_left_to_the_other_object() {
    // Both POI objects share one EMF preview.
    let package = garbled(&[]);
    let objects = package.equation_objects(0).unwrap();
    assert_eq!(objects.len(), 2);
    assert!(objects.iter().all(|object| object.error().is_some()));
    assert!(objects[0].latex(&limits()).is_err());

    let before = parts(&package);
    let mut edit = package.edit_equation_objects().unwrap();
    edit.set(0, 1025, Source::Latex("a+b"), &limits()).unwrap();
    let set = edit.commit().unwrap().into_package();
    let objects = set.equation_objects(0).unwrap();
    assert_eq!(norm(&objects[0].latex(&limits()).unwrap()), "a+b");
    assert_eq!(objects[0].prog_id(), "Equation.DSMT4");
    assert_eq!(objects[0].preview(), Some("/xl/media/image1.png"));
    assert_eq!(objects[1].preview(), Some("/xl/media/image1.emf"));
    let after = parts(&set);
    assert_eq!(
        after["/xl/media/image1.emf"].1,
        before["/xl/media/image1.emf"].1
    );
    assert_eq!(
        after["/xl/embeddings/oleObject2.bin"],
        before["/xl/embeddings/oleObject2.bin"]
    );
    let vml = text(&set, "/xl/drawings/vmlDrawing1.vml");
    assert_eq!(vml.matches("o:relid=\"rId1\"").count(), 1);

    // Removing the other object releases the EMF.
    let mut edit = set.edit_equation_objects().unwrap();
    edit.remove(0, 1026).unwrap();
    let removed = edit.commit().unwrap().into_package();
    let names = parts(&removed);
    assert!(!names.contains_key("/xl/media/image1.emf"));
    assert!(!names.contains_key("/xl/embeddings/oleObject2.bin"));
    assert!(names.contains_key("/xl/media/image1.png"));
    assert_eq!(removed.equation_objects(0).unwrap().len(), 1);
}

/// The POI workbook with its two objects named equation objects (their
/// embeddings are not equation compound files), without the parts `drop`.
fn garbled(drop: &[&str]) -> Package {
    let package = Package::from_bytes(POI).unwrap();
    let mut opc = package.opc().clone();
    let sheet = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let xml = String::from_utf8(opc.get_part(&sheet).unwrap().blob().to_vec())
        .unwrap()
        .replace("progId=\"Package2\"", "progId=\"Equation.3\"")
        .replace("progId=\"Package\"", "progId=\"Equation.DSMT4\"");
    opc.get_part_mut(&sheet).unwrap().set_blob(xml.into_bytes());
    for name in drop {
        opc.remove_part(&PackURI::new(*name).unwrap());
    }
    Package::from_opc(opc).unwrap()
}

#[test]
fn missing_parts_are_per_object_errors_and_are_left_alone() {
    let package = garbled(&[
        "/xl/embeddings/oleObject1.bin",
        "/xl/drawings/vmlDrawing1.vml",
        "/xl/drawings/drawing1.xml",
    ]);
    let objects = package.equation_objects(0).unwrap();
    assert_eq!(objects.len(), 2);
    assert!(objects[0].error().is_some());
    assert!(objects[0].native().is_none());

    // The missing embedding cannot be replaced; the other object can.
    let mut edit = package.edit_equation_objects().unwrap();
    assert!(edit.set(0, 1025, Source::Latex("x"), &limits()).is_err());
    edit.set(0, 1026, Source::Latex("x"), &limits()).unwrap();
    edit.remove(0, 1025).unwrap();
    let after = edit.commit().unwrap().into_package();
    let objects = after.equation_objects(0).unwrap();
    assert_eq!(objects.len(), 1);
    assert_eq!(objects[0].prog_id(), "Equation.3");
    assert_eq!(objects[0].latex(&limits()).unwrap(), "x");
}

#[test]
fn hostile_input_is_refused_without_a_partial_edit() {
    let package = Package::new().unwrap();
    let mut edit = package.edit_equation_objects().unwrap();
    assert!(
        edit.add(0, 0, 0, Source::Latex(r"\frac{"), &limits())
            .is_err()
    );
    assert!(
        edit.add(0, 0, 0, Source::Mtef(b"garbage"), &limits())
            .is_err()
    );
    assert!(
        edit.add(0, 2_000_000, 0, Source::Latex("x"), &limits())
            .is_err()
    );
    assert!(edit.add(9, 0, 0, Source::Latex("x"), &limits()).is_err());
    let huge = "x".repeat(4 * 1024 * 1024);
    assert!(edit.add(0, 0, 0, Source::Latex(&huge), &limits()).is_err());
    assert!(edit.set(0, 1025, Source::Latex("x"), &limits()).is_err());
    assert!(edit.remove(0, 1025).is_err());
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(parts(commit.package()), parts(&package));
}

#[test]
fn patches_replay_and_invert() {
    let package = Package::new().unwrap();
    let mut edit = package.edit_equation_objects().unwrap();
    edit.add(0, 0, 0, Source::Latex("y"), &limits()).unwrap();
    let commit = edit.commit().unwrap();
    let replayed = package
        .apply_equation_objects(commit.patch())
        .unwrap()
        .into_package();
    assert_eq!(parts(&replayed), parts(commit.package()));
    let inverted = commit
        .package()
        .apply_equation_objects(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(parts(&inverted), parts(&package));
}

#[test]
fn excel_written_vml_shapes_are_found_by_id() {
    let package = Package::from_bytes(EXCEL).unwrap();
    let vml = text(&package, "/xl/drawings/vmlDrawing1.vml");
    let ids = vml
        .match_indices("<v:shape ")
        .map(|(at, _)| markup::vml_shape_id(&vml.as_bytes()[at..]))
        .collect::<Vec<_>>();
    assert_eq!(ids, vec![Some(1025), Some(1027)]);
}
