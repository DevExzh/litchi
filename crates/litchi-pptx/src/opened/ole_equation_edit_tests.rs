//! OLE equation objects: read, add, replace and remove, on PowerPoint's own
//! markup-compatibility form and on objects this crate writes.

use std::collections::BTreeMap;
use std::sync::Arc;

use litchi_equation::{Limits, MtefTarget};
use litchi_opc::PackURI;

use crate::shape::math::OleEquationSource as EquationSource;
use crate::{Error, Package, Result};

/// PowerPoint 2013: a MathType frame in the VML-choice form with a WMF
/// preview, a legacy VML drawing and an animation of the frame.
const FIXTURE: &[u8] =
    include_bytes!("../../../../test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf169088.pptx");
const SLIDE: &str = "/ppt/slides/slide1.xml";

fn part_name(name: &str) -> PackURI {
    PackURI::new(name).unwrap()
}

fn blob(package: &Package, name: &str) -> Vec<u8> {
    package
        .opc()
        .unwrap()
        .get_part(&part_name(name))
        .unwrap()
        .blob()
        .to_vec()
}

fn has_part(package: &Package, name: &str) -> bool {
    package.opc().unwrap().contains_part(&part_name(name))
}

fn content_type(package: &Package, name: &str) -> String {
    package
        .opc()
        .unwrap()
        .get_part(&part_name(name))
        .unwrap()
        .content_type()
        .to_owned()
}

fn parts(package: &Package) -> BTreeMap<String, Vec<u8>> {
    let opc = package.opc().unwrap();
    let names: Vec<_> = opc
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    names
        .into_iter()
        .map(|name| {
            let blob = opc.get_part(&name).unwrap().blob().to_vec();
            (name.to_string(), blob)
        })
        .collect()
}

/// The parts of `after` that are not byte-identical in `before` (changed
/// or new), and the parts of `before` that are gone.
fn changes(before: &Package, after: &Package) -> (Vec<String>, Vec<String>) {
    let (a, b) = (parts(before), parts(after));
    let changed = b
        .iter()
        .filter(|(name, value)| a.get(*name) != Some(value))
        .map(|(name, _)| name.clone())
        .collect();
    let gone = a
        .keys()
        .filter(|name| !b.contains_key(*name))
        .cloned()
        .collect();
    (changed, gone)
}

fn reopen(package: &Package) -> Package {
    Package::from_bytes(&package.to_bytes().unwrap()).unwrap()
}

fn ole(package: &Package) -> Vec<crate::shape::math::OleEquation> {
    let presentation = package.presentation().unwrap();
    presentation
        .slide(0)
        .unwrap()
        .unwrap()
        .ole_equations()
        .unwrap()
}

fn compact(latex: &str) -> String {
    latex.chars().filter(|ch| !ch.is_whitespace()).collect()
}

fn latex(equation: &crate::shape::math::OleEquation) -> String {
    compact(&equation.latex(&Limits::DEFAULT).unwrap())
}

fn find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

/// The slide from its root element on.
fn root(xml: &[u8]) -> &[u8] {
    &xml[find(xml, b"<p:sld").expect("a slide root")..]
}

fn contains(haystack: &[u8], needle: &str) -> bool {
    find(haystack, needle.as_bytes()).is_some()
}

/// The fixture with its slide XML passed through `change`.
fn with_slide(change: impl FnOnce(String) -> String) -> Package {
    let mut package = Package::from_bytes(FIXTURE).unwrap();
    let xml = String::from_utf8(blob(&package, SLIDE)).unwrap();
    Arc::make_mut(&mut package.opc)
        .get_part_mut(&part_name(SLIDE))
        .unwrap()
        .set_blob(change(xml).into_bytes());
    reopen(&package)
}

/// The fixture without its animation, which refuses the frame's removal.
fn without_timing() -> Package {
    with_slide(|xml| {
        let start = xml.find("<p:timing>").unwrap();
        let end = xml.find("</p:timing>").unwrap() + "</p:timing>".len();
        format!("{}{}", &xml[..start], &xml[end..])
    })
}

/// The members of a saved package.
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let archive = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    archive
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = archive.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

#[test]
fn powerpoints_equation_object_is_listed_eagerly_and_from_the_source() -> Result<()> {
    let package = Package::from_bytes(FIXTURE)?;
    let listed = ole(&package);
    assert_eq!(listed.len(), 1);
    let equation = &listed[0];
    assert_eq!(equation.shape_id(), Some(295_958));
    assert_eq!(equation.name(), Some("Object 22"));
    assert_eq!(equation.prog_id(), "Equation.DSMT4");
    assert_eq!(equation.offset(), Some((3_806_825, 5_440_363)));
    assert_eq!(equation.extent(), Some((1_101_725, 752_475)));
    assert_eq!(
        equation.embedding_part(),
        Some("/ppt/embeddings/oleObject1.bin")
    );
    assert_eq!(equation.preview_part(), Some("/ppt/media/image1.wmf"));
    assert!(equation.has_legacy_drawing());
    let native = equation.equation_native().unwrap();
    let read = litchi_equation::read_mtef(&native, &Limits::DEFAULT)?;
    assert_eq!(read.info().unwrap().version, 5);
    let eager_latex = latex(equation);
    assert!(!eager_latex.is_empty());

    let source = crate::presentation::source::SourceBackedPresentation::from_read_at(Arc::new(
        litchi_core::OwnedSource::new(FIXTURE.to_vec()),
    ))?;
    let listed = source.slide(0).unwrap().ole_equations()?;
    assert_eq!(listed.len(), 1);
    assert_eq!(listed[0].shape_id(), Some(295_958));
    assert_eq!(&*listed[0].equation_native().unwrap(), &*native);
    assert_eq!(latex(&listed[0]), eager_latex);
    // The Office Math reader lists none of them.
    let presentation = package.presentation()?;
    assert!(presentation.slide(0)?.unwrap().equations()?.is_empty());
    Ok(())
}

#[test]
fn an_unreadable_embedding_is_listed_with_its_reason() -> Result<()> {
    let mut package = Package::from_bytes(FIXTURE)?;
    Arc::make_mut(&mut package.opc)
        .get_part_mut(&part_name("/ppt/embeddings/oleObject1.bin"))?
        .set_blob(b"not a compound file".to_vec());
    let listed = ole(&reopen(&package));
    assert_eq!(listed.len(), 1);
    let error = listed[0].equation_native().unwrap_err();
    assert!(error.reason().contains("OLE equation object"), "{error}");
    assert!(listed[0].latex(&Limits::DEFAULT).is_err());

    // A relationship to a missing part, and one that is missing.
    let missing = with_slide(|xml| xml.replace(r#"r:id="rId3""#, r#"r:id="rId99""#));
    let listed = ole(&missing);
    assert_eq!(listed.len(), 1);
    assert_eq!(listed[0].embedding_part(), None);
    assert!(listed[0].equation_native().is_err());
    Ok(())
}

#[test]
fn a_new_mathtype_object_is_written_in_the_plain_form() -> Result<()> {
    let package = Package::from_bytes(FIXTURE)?;
    let mut edit = package.edit()?;
    let written = edit.add_ole_equation(
        0usize,
        EquationSource::Latex(r"\frac{a}{b}"),
        (914_400, 457_200),
        None,
    )?;
    assert_eq!(written.index, 1);
    assert!(written.losses.is_empty(), "{:?}", written.losses);
    assert_eq!(written.shape_id, 295_959);
    let after = reopen(&edit.commit()?.into_package());

    let listed = ole(&after);
    assert_eq!(listed.len(), 2);
    let added = &listed[1];
    assert_eq!(added.shape_id(), Some(295_959));
    assert_eq!(added.prog_id(), "Equation.DSMT4");
    assert!(!added.has_legacy_drawing());
    assert_eq!(latex(added), r"\frac{a}{b}");
    assert_eq!(latex(&listed[0]), latex(&ole(&package)[0]));
    let object = added.embedding_part().unwrap();
    assert!(object.starts_with("/ppt/embeddings/oleObject"));
    assert_eq!(
        content_type(&after, object),
        "application/vnd.openxmlformats-officedocument.oleObject"
    );
    let preview = added.preview_part().unwrap();
    assert_eq!(content_type(&after, preview), "image/png");
    assert!(blob(&after, preview).starts_with(b"\x89PNG"));
    // The frame is the equation's size at 18 points, and the object's
    // compound file carries the MathType class.
    let (cx, cy) = added.extent().unwrap();
    assert!(cx > 100_000 && cy > 100_000, "{cx} x {cy}");
    assert_eq!(added.offset(), Some((914_400, 457_200)));
    let bin = blob(&after, object);
    assert_eq!(
        litchi_equation::equation_object_native(&bin)?,
        &*added.equation_native().unwrap()
    );

    let xml = blob(&after, SLIDE);
    let start = find(&xml, br#"<p:graphicFrame xmlns:p="#).unwrap();
    let frame = std::str::from_utf8(&xml[start..]).unwrap();
    assert!(frame.contains(r#"<a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj name="Equation" r:id="#), "{frame}");
    assert!(
        frame.contains(&format!(
            r#"imgW="{cx}" imgH="{cy}" progId="Equation.DSMT4"><p:embed/><p:pic>"#
        )),
        "{frame}"
    );
    // Only the slide, its relationships and the two new parts change.
    let (changed, gone) = changes(&package, &after);
    assert!(gone.is_empty(), "{gone:?}");
    let mut expected = vec![object.to_owned(), preview.to_owned(), SLIDE.to_owned()];
    expected.sort();
    let mut changed = changed;
    changed.retain(|name| name != "/[Content_Types].xml");
    assert_eq!(changed, expected);
    // The slide before the new frame is unchanged (a commit drops the line
    // break PowerPoint writes after the XML declaration).
    let before = blob(&package, SLIDE);
    let (before, xml) = (root(&before), root(&xml));
    let tree_end = find(before, b"</p:spTree>").unwrap();
    assert_eq!(&xml[..tree_end], &before[..tree_end]);
    Ok(())
}

#[test]
fn raw_mtef_three_makes_an_equation_editor_object_that_keeps_its_kind() -> Result<()> {
    let v3 = litchi_equation::latex_to_equation_native(
        "x^{2}",
        &MtefTarget::EquationEditor3,
        &Limits::DEFAULT,
    )?;
    let package = Package::from_bytes(FIXTURE)?;
    let mut edit = package.edit()?;
    edit.add_ole_equation(0usize, EquationSource::Mtef(v3.bytes()), (0, 0), Some(18.0))?;
    // MTEF 5 cannot replace it; LaTeX is written as MTEF 3.
    let v5 =
        litchi_equation::latex_to_equation_native("y", &MtefTarget::MathType, &Limits::DEFAULT)?;
    assert!(matches!(
        edit.set_ole_equation(0usize, 1, EquationSource::Mtef(v5.bytes())),
        Err(Error::Refused(_))
    ));
    edit.set_ole_equation(0usize, 1, EquationSource::Latex("y+1"))?;
    let after = reopen(&edit.commit()?.into_package());
    let listed = ole(&after);
    assert_eq!(listed[1].prog_id(), "Equation.3");
    let native = listed[1].equation_native().unwrap();
    assert_eq!(native[28], 3);
    assert_eq!(latex(&listed[1]), "y+1");
    Ok(())
}

#[test]
fn powerpoints_object_is_replaced_in_the_plain_form_without_its_vml_shape() -> Result<()> {
    let package = Package::from_bytes(FIXTURE)?;
    let old = blob(&package, "/ppt/embeddings/oleObject1.bin");
    let mut edit = package.edit()?;
    let written = edit
        .set_ole_equation(0usize, 0, EquationSource::Latex("x^{2}+1"))?
        .expect("the equation changed");
    assert_eq!(written.shape_id, 295_958);
    assert_eq!(written.index, 0);
    let after = reopen(&edit.commit()?.into_package());

    let listed = ole(&after);
    assert_eq!(listed.len(), 1);
    let equation = &listed[0];
    assert_eq!(latex(equation), "x^{2}+1");
    assert_eq!(equation.prog_id(), "Equation.DSMT4");
    assert!(!equation.has_legacy_drawing());
    // The compound file keeps its name, class and other streams.
    assert_eq!(
        equation.embedding_part(),
        Some("/ppt/embeddings/oleObject1.bin")
    );
    let bin = blob(&after, "/ppt/embeddings/oleObject1.bin");
    assert_ne!(bin, old);
    let native = litchi_equation::equation_object_native(&bin)?;
    assert_eq!(native[28], 5);
    // The WMF preview and the VML drawing that showed the old equation
    // are gone; a new PNG preview takes their place.
    let preview = equation.preview_part().unwrap();
    assert!(preview.ends_with(".png"), "{preview}");
    for gone in [
        "/ppt/media/image1.wmf",
        "/ppt/drawings/vmlDrawing1.vml",
        "/ppt/drawings/_rels/vmlDrawing1.vml.rels",
    ] {
        assert!(!has_part(&after, gone), "{gone}");
    }
    let xml = blob(&after, SLIDE);
    assert!(!contains(&xml, "mc:AlternateContent"));
    assert!(!contains(&xml, "spid=\"_x0000"));
    // The frame keeps its position, its picture keeps its fill and line,
    // and the animation of the frame stays.
    assert_eq!(equation.offset(), Some((3_806_825, 5_440_363)));
    assert!(contains(&xml, r#"<a:srgbClr val="FFFF00"/>"#));
    assert!(contains(&xml, r#"<p:spTgt spid="295958"/>"#));
    let (cx, cy) = equation.extent().unwrap();
    assert!(contains(&xml, &format!(r#"imgW="{cx}" imgH="{cy}""#)));
    // The slide outside the frame is unchanged.
    let before = blob(&package, SLIDE);
    let (before, xml) = (root(&before).to_vec(), root(&xml).to_vec());
    let frame = find(&before, b"<p:graphicFrame>").unwrap();
    assert_eq!(&xml[..frame], &before[..frame]);
    let tail = find(&before, b"</p:graphicFrame>").unwrap();
    let after_tail = find(&xml, b"</p:graphicFrame>").unwrap();
    assert_eq!(&xml[after_tail..], &before[tail..]);
    let (changed, _) = changes(&package, &after);
    assert!(
        changed.iter().all(|name| [
            SLIDE,
            "/ppt/slides/_rels/slide1.xml.rels",
            "/ppt/embeddings/oleObject1.bin",
            preview,
            "/[Content_Types].xml"
        ]
        .contains(&name.as_str())),
        "{changed:?}"
    );
    Ok(())
}

#[test]
fn a_png_preview_only_one_object_shows_is_replaced_in_place() -> Result<()> {
    let package = Package::from_bytes(FIXTURE)?;
    let mut edit = package.edit()?;
    edit.add_ole_equation(0usize, EquationSource::Latex("a"), (0, 0), None)?;
    let added = reopen(&edit.commit()?.into_package());
    let listed = ole(&added);
    let preview = listed[1].preview_part().unwrap().to_owned();
    let (cx, cy) = listed[1].extent().unwrap();
    let before = blob(&added, &preview);

    let mut edit = added.edit()?;
    let omml = "<m:oMath><m:r><m:t>a</m:t></m:r><m:r><m:t>+</m:t></m:r><m:r><m:t>b</m:t></m:r><m:r><m:t>+</m:t></m:r><m:r><m:t>c</m:t></m:r></m:oMath>";
    edit.set_ole_equation(0usize, 1, EquationSource::Omml(omml))?;
    let after = reopen(&edit.commit()?.into_package());
    let listed = ole(&after);
    assert_eq!(latex(&listed[1]), "a+b+c");
    assert_eq!(listed[1].preview_part(), Some(preview.as_str()));
    assert_ne!(blob(&after, &preview), before);
    // The frame widened with the equation, at the same height.
    let (new_cx, new_cy) = listed[1].extent().unwrap();
    assert!(new_cx > cx, "{new_cx} <= {cx}");
    assert!(new_cy >= cy, "{new_cy} < {cy}");
    // The same content again changes nothing.
    let native = listed[1].equation_native().unwrap();
    let mut edit = after.edit()?;
    assert!(
        edit.set_ole_equation(0usize, 1, EquationSource::Mtef(&native))?
            .is_none()
    );
    Ok(())
}

#[test]
fn a_garbled_embedding_is_replaced_by_a_new_object() -> Result<()> {
    let mut package = Package::from_bytes(FIXTURE)?;
    Arc::make_mut(&mut package.opc)
        .get_part_mut(&part_name("/ppt/embeddings/oleObject1.bin"))?
        .set_blob(vec![0xD0; 4096]);
    let package = reopen(&package);
    let mut edit = package.edit()?;
    edit.set_ole_equation(0usize, 0, EquationSource::Latex("z"))?;
    let after = reopen(&edit.commit()?.into_package());
    let listed = ole(&after);
    assert_eq!(latex(&listed[0]), "z");
    // The ProgID names the kind of the new object.
    assert_eq!(listed[0].equation_native().unwrap()[28], 5);

    // An object without its part cannot be replaced, but can be removed.
    let missing = with_slide(|xml| {
        let start = xml.find("<p:timing>").unwrap();
        let end = xml.find("</p:timing>").unwrap() + "</p:timing>".len();
        format!("{}{}", &xml[..start], &xml[end..]).replace(r#"r:id="rId3""#, r#"r:id="rId99""#)
    });
    let mut edit = missing.edit()?;
    assert!(matches!(
        edit.set_ole_equation(0usize, 0, EquationSource::Latex("z")),
        Err(Error::Refused(_))
    ));
    assert!(matches!(
        edit.set_ole_equation(0usize, 1, EquationSource::Latex("z")),
        Err(Error::IndexOutOfBounds { .. })
    ));
    assert_eq!(edit.remove_ole_equation(0usize, 0)?, 295_958);
    Ok(())
}

#[test]
fn removing_an_animated_frame_is_refused() -> Result<()> {
    let package = Package::from_bytes(FIXTURE)?;
    let mut edit = package.edit()?;
    assert!(matches!(
        edit.remove_ole_equation(0usize, 0),
        Err(Error::Refused(_))
    ));
    // The refusal left the edit unchanged.
    let after = edit.commit()?.into_package();
    assert_eq!(blob(&after, SLIDE), blob(&package, SLIDE));
    assert!(has_part(&after, "/ppt/embeddings/oleObject1.bin"));
    Ok(())
}

/// A removed object in PowerPoint's markup-compatibility form leaves no
/// trace in the saved file: not its embedding, its WMF preview, its VML
/// shape and drawing, their relationships or their content types.
#[test]
fn a_removed_object_leaves_nothing_behind_in_the_saved_file() -> Result<()> {
    let package = without_timing();
    let native = ole(&package)[0].equation_native().unwrap().to_vec();
    let mut edit = package.edit()?;
    assert_eq!(edit.remove_ole_equation(0usize, 0)?, 295_958);
    let bytes = edit.commit()?.into_package().to_bytes()?;
    let saved = members(&bytes);
    for gone in [
        "ppt/embeddings/oleObject1.bin",
        "ppt/media/image1.wmf",
        "ppt/drawings/vmlDrawing1.vml",
        "ppt/drawings/_rels/vmlDrawing1.vml.rels",
    ] {
        assert!(!saved.contains_key(gone), "{gone}");
    }
    let types = &saved["[Content_Types].xml"];
    for extension in [
        r#"Extension="bin""#,
        r#"Extension="wmf""#,
        r#"Extension="vml""#,
    ] {
        assert!(!contains(types, extension), "{extension}");
    }
    let rels = &saved["ppt/slides/_rels/slide1.xml.rels"];
    for kind in ["/oleObject", "/image", "/vmlDrawing"] {
        assert!(!contains(rels, kind), "{kind}");
    }
    assert!(!contains(&saved["ppt/slides/slide1.xml"], "graphicFrame"));
    // No member holds the equation's MTEF any more.
    let mtef = &native[28..];
    assert!(saved.values().all(|member| find(member, mtef).is_none()));
    let reread = Package::from_bytes(&bytes)?;
    assert!(ole(&reread).is_empty());
    Ok(())
}

/// A VML drawing that holds another shape keeps it and loses only the
/// removed object's shape and the relationship only that shape used.
#[test]
fn a_shared_vml_drawing_keeps_its_other_shapes() -> Result<()> {
    let mut package = without_timing();
    let vml_name = "/ppt/drawings/vmlDrawing1.vml";
    let vml = String::from_utf8(blob(&package, vml_name)).unwrap();
    let other = r##"<v:shape id="_x0000_s1" type="#_x0000_t75" style='width:1pt;height:1pt'/>"##;
    let vml = vml.replace("</xml>", &format!("{other}</xml>"));
    Arc::make_mut(&mut package.opc)
        .get_part_mut(&part_name(vml_name))?
        .set_blob(vml.into_bytes());
    let package = reopen(&package);
    let mut edit = package.edit()?;
    edit.remove_ole_equation(0usize, 0)?;
    let after = reopen(&edit.commit()?.into_package());
    let kept = blob(&after, vml_name);
    assert!(contains(&kept, other));
    assert!(!contains(&kept, "_x0000_s295990"));
    assert!(!contains(&kept, "o:relid"));
    assert!(!has_part(&after, "/ppt/media/image1.wmf"));
    assert!(!has_part(&after, "/ppt/embeddings/oleObject1.bin"));
    let opc = after.opc().unwrap();
    let drawing = opc.get_part(&part_name(vml_name))?;
    assert_eq!(drawing.rels().iter().count(), 0);
    Ok(())
}

/// An object that names only its VML shape (PowerPoint 2007's form, no
/// `p:pic`) gets a PNG preview of its own, and its VML shape leaves.
#[test]
fn an_object_without_a_picture_gets_one() -> Result<()> {
    let package = with_slide(|xml| {
        let start = xml.find("<mc:AlternateContent").unwrap();
        let end = xml.find("</mc:AlternateContent>").unwrap() + "</mc:AlternateContent>".len();
        format!(
            r#"{}<p:oleObj spid="_x0000_s295990" name="Equation" r:id="rId3" imgW="647640" imgH="444240" progId="Equation.DSMT4"><p:embed/></p:oleObj>{}"#,
            &xml[..start],
            &xml[end..]
        )
    });
    let listed = ole(&package);
    assert!(listed[0].has_legacy_drawing());
    assert_eq!(listed[0].preview_part(), None);
    let mut edit = package.edit()?;
    edit.set_ole_equation(0usize, 0, EquationSource::Latex("n+1"))?;
    let after = reopen(&edit.commit()?.into_package());
    let listed = ole(&after);
    assert_eq!(latex(&listed[0]), "n+1");
    assert!(!listed[0].has_legacy_drawing());
    let preview = listed[0].preview_part().unwrap();
    assert!(blob(&after, preview).starts_with(b"\x89PNG"));
    assert!(!has_part(&after, "/ppt/drawings/vmlDrawing1.vml"));
    let xml = blob(&after, SLIDE);
    assert!(!contains(&xml, "spid=\"_x0000"));
    assert!(contains(&xml, "<p:embed/><p:pic xmlns:a="));
    Ok(())
}

/// The patch of a removal or a replacement replays on the source, and its
/// inverse restores every part and the content-types manifest.
#[test]
fn equation_object_patches_replay_and_invert() -> Result<()> {
    let package = without_timing();
    let source_bytes = package.to_bytes()?;
    for stage in [0, 1] {
        let mut edit = package.edit()?;
        if stage == 0 {
            edit.remove_ole_equation(0usize, 0)?;
        } else {
            edit.set_ole_equation(0usize, 0, EquationSource::Latex("q"))?;
        }
        let (changed, patch) = edit.commit()?.into_parts();
        let output = changed.to_bytes()?;
        let replayed = package.apply(&patch)?.into_package().to_bytes()?;
        assert!(members(&replayed) == members(&output), "stage {stage}");
        let restored = changed.apply(&patch.inverse())?.into_package().to_bytes()?;
        assert!(
            members(&restored) == members(&source_bytes),
            "stage {stage}"
        );
    }
    Ok(())
}

#[test]
fn hostile_input_is_refused_without_panicking() -> Result<()> {
    let package = Package::from_bytes(FIXTURE)?;
    let mut edit = package.edit()?;
    for source in [
        EquationSource::Latex(r"\frac{"),
        EquationSource::Omml("<m:oMath><m:r>"),
        EquationSource::Mtef(&[0xFF; 64]),
        EquationSource::Mtef(&[]),
    ] {
        assert!(edit.add_ole_equation(0usize, source, (0, 0), None).is_err());
        assert!(edit.set_ole_equation(0usize, 0, source).is_err());
    }
    let huge = "x+".repeat(2_000_000);
    assert!(
        edit.add_ole_equation(0usize, EquationSource::Latex(&huge), (0, 0), None)
            .is_err()
    );
    assert!(
        edit.add_ole_equation(0usize, EquationSource::Latex("x"), (i64::MAX, 0), None)
            .is_err()
    );
    assert!(
        edit.add_ole_equation(0usize, EquationSource::Latex("x"), (0, 0), Some(f64::NAN))
            .is_err()
    );
    assert!(matches!(
        edit.remove_ole_equation(0usize, 5),
        Err(Error::IndexOutOfBounds { .. })
    ));
    // Nothing was staged.
    let after = edit.commit()?.into_package();
    assert_eq!(parts(&after), parts(&package));
    Ok(())
}
