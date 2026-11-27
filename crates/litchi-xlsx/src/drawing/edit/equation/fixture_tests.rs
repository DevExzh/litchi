//! Worksheet equations: read, add, replace and remove, on an ONLYOFFICE
//! file and on hand-made drawings.

use std::sync::Arc;

use litchi_core::edit::RefusalKind;
use litchi_opc::PackURI;

use super::{EquationSource, Marker};
use crate::error::Error;
use crate::{Package, Result};

/// An ONLYOFFICE-written workbook whose first sheet's drawing holds an
/// equation inside a text box paragraph (see `test-data/poi/PROVENANCE.md`).
const ONLYOFFICE: &[u8] = include_bytes!(
    "../../../../../../test-data/poi/test-data/spreadsheet/LIBRE_OFFICE-128382-0-small-media.xlsx"
);
const IMAGE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

fn part_name(name: &str) -> PackURI {
    PackURI::new(name).unwrap()
}

fn blob(package: &Package, name: &str) -> Vec<u8> {
    package
        .opc()
        .get_part(&part_name(name))
        .unwrap()
        .blob()
        .to_vec()
}

fn text(package: &Package, name: &str) -> String {
    String::from_utf8(blob(package, name)).unwrap()
}

fn has_part(package: &Package, name: &str) -> bool {
    package.opc().get_part(&part_name(name)).is_ok()
}

/// Every part's name and bytes, sorted by name.
fn parts(package: &Package) -> Vec<(String, Vec<u8>)> {
    let opc = package.opc();
    let names: Vec<PackURI> = opc
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    let mut out: Vec<_> = names
        .into_iter()
        .map(|name| {
            let blob = opc.get_part(&name).unwrap().blob().to_vec();
            (name.to_string(), blob)
        })
        .collect();
    out.sort();
    out
}

/// The image relationships of a part: ID and target.
fn images(package: &Package, name: &str) -> Vec<(String, String)> {
    let part = package.opc().get_part(&part_name(name)).unwrap();
    let mut out: Vec<_> = part
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == IMAGE)
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.target_partname().unwrap().to_string(),
            )
        })
        .collect();
    out.sort();
    out
}

fn reopen(package: &Package) -> Package {
    Package::from_bytes(&package.to_bytes().unwrap()).unwrap()
}

fn compact(latex: &str) -> String {
    latex.chars().filter(|ch| !ch.is_whitespace()).collect()
}

fn latex_of(package: &Package, sheet: usize) -> Vec<String> {
    package
        .equations(sheet)
        .unwrap()
        .iter()
        .map(|equation| compact(&equation.latex().unwrap()))
        .collect()
}

fn refusal(error: &Error, kind: RefusalKind) -> bool {
    matches!(error, Error::Refused(refusal) if refusal.kind() == kind)
}

/// The parts that differ between two packages, by name.
fn changed(before: &Package, after: &Package) -> Vec<String> {
    let before = parts(before);
    let after = parts(after);
    let mut names: Vec<String> = before
        .iter()
        .chain(after.iter())
        .map(|(name, _)| name.clone())
        .collect();
    names.sort();
    names.dedup();
    names
        .into_iter()
        .filter(|name| {
            before.iter().find(|(n, _)| n == name) != after.iter().find(|(n, _)| n == name)
        })
        .collect()
}

#[test]
fn an_onlyoffice_equation_is_read_with_its_cell_and_shape() -> Result<()> {
    let package = Package::from_bytes(ONLYOFFICE)?;
    let found = package.equations(0)?;
    assert_eq!(found.len(), 1);
    let equation = &found[0];
    assert_eq!(equation.shape_id(), Some(6));
    assert_eq!(equation.name(), Some("TextBox 3"));
    assert_eq!(equation.cell().as_deref(), Some("A22"));
    assert!(
        equation.omml().starts_with("<m:oMathPara>"),
        "{}",
        equation.omml()
    );
    assert!(equation.text().starts_with("𝑒𝑥=1+"), "{}", equation.text());
    let latex = compact(&equation.latex()?);
    assert!(latex.starts_with(r"e^x=1+\frac{x}{1!}"), "{latex}");
    // The converter also reads it as linear text.
    assert!(
        litchi_equation::omml_to_unicode_math(equation.omml(), &litchi_equation::Limits::DEFAULT)?
            .contains("=1+")
    );
    // The other sheet has a drawing without equations.
    assert!(package.equations(1)?.is_empty());

    // The source-backed reader lists the same equations.
    let workbook = crate::workbook::SourceBackedWorkbook::from_read_at(Arc::new(
        litchi_core::OwnedSource::new(ONLYOFFICE.to_vec()),
    ))?;
    let sheet = workbook.sheet(0usize)?.unwrap();
    assert_eq!(sheet.equations()?, found);
    Ok(())
}

#[test]
fn an_equation_inside_a_paragraph_is_refused_for_edits() -> Result<()> {
    let package = Package::from_bytes(ONLYOFFICE)?;
    let mut edit = package.edit_drawings()?;
    let error = edit
        .set_equation(0, 0, EquationSource::Latex("y"))
        .unwrap_err();
    assert!(
        refusal(&error, RefusalKind::UnsupportedContent),
        "{error:?}"
    );
    let error = edit.remove_equation(0, 0).unwrap_err();
    assert!(
        refusal(&error, RefusalKind::UnsupportedContent),
        "{error:?}"
    );
    let error = edit
        .set_equation(0, 1, EquationSource::Latex("y"))
        .unwrap_err();
    assert!(refusal(&error, RefusalKind::NotFound), "{error:?}");
    // Nothing was staged.
    assert!(edit.commit()?.patch().is_noop());
    Ok(())
}

#[test]
fn a_new_equation_creates_the_drawing_and_round_trips() -> Result<()> {
    let package = Package::new()?;
    let mut edit = package.edit_drawings()?;
    let written = edit.add_equation(
        0,
        EquationSource::Latex(r"x = \frac{-b \pm \sqrt{b^2-4ac}}{2a}"),
        Marker::cell(1, 2),
        None,
    )?;
    assert_eq!((written.shape_id, written.index), (2, 0));
    assert!(written.losses.is_empty(), "{:?}", written.losses);
    let committed = edit.commit()?.into_package();
    let after = reopen(&committed);

    // The drawing part, its worksheet relationship and `drawing` element;
    // no picture.
    let sheet = text(&after, "/xl/worksheets/sheet1.xml");
    assert!(sheet.contains(r#"<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>"#), "{sheet}");
    let mut changes = changed(&package, &committed);
    changes.retain(|name| name != "/[Content_Types].xml");
    assert_eq!(
        changes,
        ["/xl/drawings/drawing1.xml", "/xl/worksheets/sheet1.xml"]
    );
    assert!(images(&after, "/xl/drawings/drawing1.xml").is_empty());

    // A two-cell anchor around the block, the block around the text box.
    let xml = text(&after, "/xl/drawings/drawing1.xml");
    assert!(
        xml.contains(
            "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to>"
        ),
        "{xml}"
    );
    assert!(
        xml.contains(r#"</xdr:to><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><xdr:sp macro="" textlink="">"#),
        "{xml}"
    );
    assert!(
        xml.ends_with("</mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"),
        "{xml}"
    );
    assert!(xml.contains(r#"<xdr:cNvSpPr txBox="1"/>"#), "{xml}");
    assert!(
        xml.contains(r#"<xdr:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>"#),
        "{xml}"
    );
    assert!(xml.contains(r#"<a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" rtlCol="0" anchor="t"><a:spAutoFit/></a:bodyPr>"#), "{xml}");
    assert!(xml.contains(r#"<a:p><a14:m><m:oMathPara"#), "{xml}");
    // The anchor declares the DrawingML prefix; the OMML does not again.
    assert_eq!(xml.matches(r#"xmlns:a="http"#).count(), 2, "{xml}");
    assert!(
        xml.contains(r#"<a:rPr lang="en-US" sz="1100" i="1"><a:latin typeface="Cambria Math""#),
        "{xml}"
    );
    // The Fallback's text box holds the linear form as upright text.
    let fallback = &xml[xml.find("<mc:Fallback>").unwrap()..xml.find("</mc:Fallback>").unwrap()];
    assert!(
        fallback.contains(r#"<a:p><a:r><a:rPr lang="en-US" sz="1100" b="0" i="0"><a:latin typeface="Cambria Math""#),
        "{fallback}"
    );
    assert!(
        fallback.contains("<a:t>x=(−b±√(b^2−4ac))/(2a)</a:t>"),
        "{fallback}"
    );
    assert!(
        !fallback.contains("a14:m") && !fallback.contains("blip"),
        "{fallback}"
    );

    let found = after.equations(0)?;
    assert_eq!(found.len(), 1);
    assert_eq!(found[0].cell().as_deref(), Some("B3"));
    assert_eq!(found[0].shape_id(), Some(2));
    assert!(matches!(
        found[0].placement(),
        Some(crate::drawing::edit::Placement::TwoCell { to, .. }) if to.col >= 1 && to.row >= 3
    ));
    assert_eq!(latex_of(&after, 0), [r"x=\frac{-b\pm\sqrt{b^{2}-4ac}}{2a}"]);
    // The anchor is neither a picture nor a shape of the drawing.
    let elements = after.drawing_elements(0)?;
    assert_eq!(elements.len(), 1);
    assert_eq!(elements[0].kind, crate::drawing::edit::ElementKind::Other);
    Ok(())
}

#[test]
fn objects_added_after_an_equation_take_new_ids() -> Result<()> {
    let package = Package::new()?;
    let mut edit = package.edit_drawings()?;
    let equation = edit.add_equation(0, EquationSource::Latex("x"), Marker::cell(0, 0), None)?;
    let placement = crate::drawing::edit::Placement::TwoCell {
        from: Marker::cell(4, 4),
        to: Marker::cell(6, 6),
    };
    edit.add_shape(0, &crate::drawing::edit::ShapeSpec::new(placement, "Shape"))?;
    let after = edit.commit()?.into_package();
    let shapes = after.drawing_elements(0)?;
    let shape = shapes
        .iter()
        .find(|element| element.kind == crate::drawing::edit::ElementKind::Shape)
        .unwrap();
    assert_eq!(equation.shape_id, 2);
    assert_eq!(shape.id, Some(3));
    let mut edit = after.edit_drawings()?;
    let second = edit.add_equation(0, EquationSource::Latex("y"), Marker::cell(0, 9), None)?;
    assert_eq!(second.shape_id, 4);
    Ok(())
}

#[test]
fn adding_to_an_existing_drawing_keeps_every_other_byte() -> Result<()> {
    let package = Package::from_bytes(ONLYOFFICE)?;
    let before = blob(&package, "/xl/drawings/drawing1.xml");
    let mut edit = package.edit_drawings()?;
    let written = edit.add_equation(
        0,
        EquationSource::Omml("<m:oMath><m:r><m:t>q</m:t></m:r></m:oMath>"),
        Marker::cell(3, 4),
        Some(14.0),
    )?;
    assert_eq!(written.index, 1);
    // One past the largest `cNvPr` ID of the drawing.
    assert_eq!(written.shape_id, 9);
    let committed = edit.commit()?.into_package();
    let after = reopen(&committed);
    let xml = blob(&after, "/xl/drawings/drawing1.xml");
    let close = before.len() - b"</xdr:wsDr>".len();
    assert_eq!(xml[..close], before[..close]);
    assert!(xml.ends_with(b"</xdr:twoCellAnchor></xdr:wsDr>"));
    assert_eq!(changed(&package, &committed), ["/xl/drawings/drawing1.xml"]);
    let found = after.equations(0)?;
    assert_eq!(found.len(), 2);
    // The drawing's root binds the math prefix, so the OMML stays as given.
    assert_eq!(
        found[1].omml(),
        "<m:oMath><m:r><m:t>q</m:t></m:r></m:oMath>"
    );
    assert!(
        String::from_utf8_lossy(&xml).contains(r#"sz="1400" b="0" i="0""#),
        "the text size is written on the fallback run"
    );
    // The first equation keeps its bytes.
    assert_eq!(found[0], package.equations(0)?[0]);
    Ok(())
}

/// The drawing with two equations from `add_equation`.
fn two_equations() -> Package {
    let base = Package::new().unwrap();
    let mut edit = base.edit_drawings().unwrap();
    edit.add_equation(0, EquationSource::Latex("a+b"), Marker::cell(0, 0), None)
        .unwrap();
    edit.add_equation(
        0,
        EquationSource::Latex(r"\sqrt{x}"),
        Marker::cell(0, 8),
        Some(20.0),
    )
    .unwrap();
    reopen(&edit.commit().unwrap().into_package())
}

/// The second anchor of a drawing's XML, to its end.
fn second_anchor(xml: &str) -> &str {
    let first = xml.find("<xdr:twoCellAnchor").unwrap();
    let at = xml[first + 1..].find("<xdr:twoCellAnchor").unwrap() + first + 1;
    &xml[at..]
}

#[test]
fn an_equation_is_replaced_and_only_its_markup_changes() -> Result<()> {
    let package = two_equations();
    let before = text(&package, "/xl/drawings/drawing1.xml");
    let untouched = package.equations(0)?[1].clone();

    let mut edit = package.edit_drawings()?;
    let written = edit
        .set_equation(0, 0, EquationSource::Latex(r"\frac{a}{b}"))?
        .expect("the equation changed");
    assert_eq!((written.shape_id, written.index), (2, 0));
    let committed = edit.commit()?.into_package();
    assert_eq!(changed(&package, &committed), ["/xl/drawings/drawing1.xml"]);
    let after = reopen(&committed);
    let xml = text(&after, "/xl/drawings/drawing1.xml");
    assert_eq!(latex_of(&after, 0), [r"\frac{a}{b}", r"\sqrt{x}"]);
    // The other equation's anchor keeps its bytes.
    assert_eq!(second_anchor(&xml), second_anchor(&before));
    assert_eq!(after.equations(0)?[1], untouched);
    // The Fallback follows the new equation, with the old run properties.
    let fallback = &xml[xml.find("<mc:Fallback>").unwrap()..xml.find("</mc:Fallback>").unwrap()];
    assert!(fallback.contains("<a:t>a/b</a:t>"), "{fallback}");
    assert!(fallback.contains(r#"sz="1100" b="0" i="0""#), "{fallback}");
    assert!(xml.contains(r#"sz="1100" i="1""#), "{xml}");

    // An equal equation is no change at all.
    let current = after.equations(0)?[0].omml().to_owned();
    let mut edit = after.edit_drawings()?;
    assert!(
        edit.set_equation(0, 0, EquationSource::Omml(&current))?
            .is_none()
    );
    assert!(edit.commit()?.patch().is_noop());
    Ok(())
}

#[test]
fn a_longer_replacement_resizes_the_text_box_and_its_anchor() -> Result<()> {
    let package = two_equations();
    let before = text(&package, "/xl/drawings/drawing1.xml");
    let frames = |xml: &str| -> Vec<(i64, i64)> {
        xml.match_indices("<a:ext ")
            .map(|(at, _)| {
                let tag = &xml[at..at + xml[at..].find("/>").unwrap()];
                let value = |name: &str| -> i64 {
                    let start = tag.find(&format!(r#"{name}=""#)).unwrap() + name.len() + 2;
                    tag[start..start + tag[start..].find('"').unwrap()]
                        .parse()
                        .unwrap()
                };
                (value("cx"), value("cy"))
            })
            .collect()
    };
    let to = |xml: &str| -> String {
        let at = xml.find("<xdr:to>").unwrap();
        xml[at..at + xml[at..].find("</xdr:to>").unwrap()].to_owned()
    };
    let mut edit = package.edit_drawings()?;
    edit.set_equation(
        0,
        0,
        EquationSource::Latex(r"\sum_{k=1}^{n} k^2 = \frac{n(n+1)(2n+1)}{6}"),
    )?
    .expect("changed");
    let after = reopen(&edit.commit()?.into_package());
    let xml = text(&after, "/xl/drawings/drawing1.xml");
    let (old, new) = (frames(&before), frames(&xml));
    // Both branches of the first text box grow alike; the second keeps its
    // frame.
    assert_eq!(new[0], new[1]);
    assert!(
        new[0].0 > old[0].0 && new[0].1 > old[0].1,
        "{old:?} {new:?}"
    );
    assert_eq!(new[2..], old[2..]);
    // The first anchor's end marker moves; the second anchor keeps its
    // bytes.
    assert_ne!(to(&xml), to(&before));
    assert_eq!(second_anchor(&xml), second_anchor(&before));
    let placement = after.equations(0)?[0].placement();
    let Some(crate::drawing::edit::Placement::TwoCell { from, to }) = placement else {
        panic!("{placement:?}");
    };
    // 64-pixel default columns: the box spans the measured width.
    let spanned = i64::from(to.col - from.col) * 64 * 9_525 + to.col_offset - from.col_offset;
    assert_eq!(spanned, new[0].0);
    Ok(())
}

#[test]
fn an_equation_is_removed_with_its_anchor() -> Result<()> {
    let package = two_equations();
    let before = text(&package, "/xl/drawings/drawing1.xml");
    let mut edit = package.edit_drawings()?;
    assert_eq!(edit.remove_equation(0, 0)?, 2);
    let committed = edit.commit()?.into_package();
    assert_eq!(changed(&package, &committed), ["/xl/drawings/drawing1.xml"]);
    let after = reopen(&committed);
    assert_eq!(latex_of(&after, 0), [r"\sqrt{x}"]);
    let xml = text(&after, "/xl/drawings/drawing1.xml");
    assert_eq!(xml.matches("<xdr:twoCellAnchor").count(), 1);
    assert!(xml.ends_with(second_anchor(&before)), "{xml}");

    // The last equation leaves an empty drawing part, which stays.
    let mut edit = after.edit_drawings()?;
    edit.remove_equation(0, 0)?;
    let empty = reopen(&edit.commit()?.into_package());
    assert!(empty.equations(0)?.is_empty());
    assert!(has_part(&empty, "/xl/drawings/drawing1.xml"));
    Ok(())
}

/// A drawing with an equation block around a whole anchor whose Fallback
/// shows a picture (as other producers write it), and an anchor-level block
/// whose text box mixes text and an equation, with a text Fallback.
fn hand_made() -> Package {
    let a14 = r#"xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main""#;
    let anchor = |row: u32, object: &str| {
        format!(
            r#"<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="1828800" cy="381000"/>{object}<xdr:clientData/></xdr:oneCellAnchor>"#
        )
    };
    let shape = |id: u32, fill: &str, paragraph: &str| {
        format!(
            r#"<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{id}" name="TextBox {}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="381000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill}</xdr:spPr><xdr:txBody><a:bodyPr wrap="none" rtlCol="0" anchor="t"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p>{paragraph}</a:p></xdr:txBody></xdr:sp>"#,
            id - 1
        )
    };
    let math = r#"<a14:m><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath><m:r><a:rPr lang="en-US" sz="1100" i="1"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr><m:t>𝑥</m:t></m:r></m:oMath></m:oMathPara></a14:m>"#;
    let text_run = r#"<a:r><a:rPr lang="en-US" sz="1100" b="0" i="0"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr><a:t>𝑥</a:t></a:r>"#;
    let picture = r#"<a:blipFill rotWithShape="1"><a:blip r:embed="rIdPicture"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>"#;
    let root = format!(
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice {a14} Requires="a14">{}</mc:Choice><mc:Fallback>{}</mc:Fallback></mc:AlternateContent>"#,
        anchor(1, &shape(2, "<a:noFill/>", math)),
        anchor(1, &shape(2, picture, "")),
    );
    let inner = anchor(
        6,
        &format!(
            r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice {a14} Requires="a14">{}</mc:Choice><mc:Fallback>{}</mc:Fallback></mc:AlternateContent>"#,
            shape(3, "", &format!("<a:r><a:t>Area: </a:t></a:r>{math}")),
            shape(3, "", &format!("<a:r><a:t>Area: </a:t></a:r>{text_run}")),
        ),
    );
    let drawing = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            "\n",
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{}{}</xdr:wsDr>"#
        ),
        root, inner
    );
    // Stage the drawing through a shape edit, then swap in the hand-made XML
    // and its picture.
    let package = Package::new().unwrap();
    let mut edit = package.edit_drawings().unwrap();
    let spec = crate::drawing::edit::ShapeSpec::new(
        crate::drawing::edit::Placement::TwoCell {
            from: Marker::cell(0, 0),
            to: Marker::cell(1, 1),
        },
        "Placeholder",
    );
    edit.add_shape(0, &spec).unwrap();
    let mut package = edit.commit().unwrap().into_package();
    let mut opc = package.opc().clone();
    opc.add_part(Box::new(litchi_opc::BlobPart::new(
        part_name("/xl/media/image1.png"),
        "image/png".to_owned(),
        PNG.to_vec(),
    )));
    let drawing_part = opc
        .get_part_mut(&part_name("/xl/drawings/drawing1.xml"))
        .unwrap();
    drawing_part.set_blob(drawing.into_bytes());
    drawing_part.rels_mut().add_relationship(
        IMAGE.to_owned(),
        "../media/image1.png".to_owned(),
        "rIdPicture".to_owned(),
        false,
    );
    package = package.derive(opc).unwrap();
    reopen(&package)
}

const PNG: &[u8] = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR\0\0\0\x01\0\0\0\x01\x08\x06\0\0\0\x1f\x15\xc4\x89\0\0\0\rIDATx\x9cc\xf8\x0f\0\0\x01\x01\0\x05\x18\xd8\xa8\x00\0\0\0IEND\xaeB`\x82";

#[test]
fn a_picture_fallback_keeps_its_kind() -> Result<()> {
    let package = hand_made();
    assert_eq!(latex_of(&package, 0), ["x", "x"]);
    let before = text(&package, "/xl/drawings/drawing1.xml");
    let mut edit = package.edit_drawings()?;
    edit.set_equation(0, 0, EquationSource::Latex(r"\alpha^2"))?
        .expect("changed");
    let after = reopen(&edit.commit()?.into_package());
    let xml = text(&after, "/xl/drawings/drawing1.xml");
    assert_eq!(latex_of(&after, 0), [r"\alpha^{2}", "x"]);
    // The picture is redrawn as a new part; the old one leaves.
    assert!(!has_part(&after, "/xl/media/image1.png"));
    let related = images(&after, "/xl/drawings/drawing1.xml");
    assert_eq!(related.len(), 1, "{related:?}");
    assert!(blob(&after, &related[0].1).starts_with(b"\x89PNG"));
    let fallback = &xml[xml.find("<mc:Fallback>").unwrap()..xml.find("</mc:Fallback>").unwrap()];
    assert!(
        fallback.contains(&format!(r#"r:embed="{}""#, related[0].0)),
        "{fallback}"
    );
    // The self-sizing text box takes the new size in both branches and
    // anchors; the second anchor keeps its bytes.
    let tail = |xml: &str| {
        xml[xml.find("<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6").unwrap()..].to_owned()
    };
    assert_eq!(tail(&xml), tail(&before));
    assert_eq!(xml.matches(r#"cx="1828800""#).count(), 3, "{xml}");

    // Removing it removes the block and its picture.
    let mut edit = after.edit_drawings()?;
    assert_eq!(edit.remove_equation(0, 0)?, 2);
    let removed = reopen(&edit.commit()?.into_package());
    assert!(images(&removed, "/xl/drawings/drawing1.xml").is_empty());
    assert!(!has_part(&removed, &related[0].1));
    assert_eq!(latex_of(&removed, 0), ["x"]);
    Ok(())
}

#[test]
fn a_text_fallback_beside_text_is_rewritten_in_linear_form() -> Result<()> {
    let package = hand_made();
    let mut edit = package.edit_drawings()?;
    edit.set_equation(0, 1, EquationSource::Latex(r"\pi r^2"))?
        .expect("changed");
    let after = reopen(&edit.commit()?.into_package());
    let xml = text(&after, "/xl/drawings/drawing1.xml");
    assert_eq!(latex_of(&after, 0), ["x", r"\pir^{2}"]);
    let last = &xml[xml.rfind("<mc:Fallback>").unwrap()..xml.rfind("</mc:Fallback>").unwrap()];
    // The Fallback is the Choice text box with the equation as linear text,
    // carrying the Choice's `a14` declaration.
    assert!(
        last.starts_with(r#"<mc:Fallback><xdr:sp xmlns:a14="#),
        "{last}"
    );
    assert!(last.contains("<a:t>Area: </a:t>"), "{last}");
    assert!(
        last.contains(r#"b="0" i="0""#) && last.contains("r^2</a:t>"),
        "{last}"
    );
    assert!(!last.contains("a14:m"), "{last}");
    Ok(())
}

#[test]
fn an_equation_beside_text_is_removed_and_its_block_unwrapped() -> Result<()> {
    let package = hand_made();
    let mut edit = package.edit_drawings()?;
    assert_eq!(edit.remove_equation(0, 1)?, 3);
    let after = reopen(&edit.commit()?.into_package());
    let xml = text(&after, "/xl/drawings/drawing1.xml");
    assert_eq!(latex_of(&after, 0), ["x"]);
    // The anchor holds the Choice shape again, without the block.
    assert!(
        xml.contains(r#"<xdr:ext cx="1828800" cy="381000"/><xdr:sp xmlns:a14="#),
        "{xml}"
    );
    assert!(
        xml.contains("<a:p><a:r><a:t>Area: </a:t></a:r></a:p>"),
        "{xml}"
    );
    assert_eq!(xml.matches("<mc:AlternateContent").count(), 1);
    // The picture of the other equation stays.
    assert_eq!(images(&after, "/xl/drawings/drawing1.xml").len(), 1);
    Ok(())
}

#[test]
fn bad_input_is_a_typed_error_and_stages_nothing() -> Result<()> {
    let package = hand_made();
    let mut edit = package.edit_drawings()?;
    assert!(matches!(
        edit.set_equation(0, 0, EquationSource::Latex(r"\frac{")),
        Err(Error::Equation(_))
    ));
    assert!(matches!(
        edit.set_equation(0, 0, EquationSource::Omml("<m:r/>")),
        Err(Error::Invalid(_))
    ));
    let error = edit
        .set_equation(
            0,
            0,
            EquationSource::Omml(
                r#"<m:oMath xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><m:r r:id="rId1"><m:t>x</m:t></m:r></m:oMath>"#,
            ),
        )
        .unwrap_err();
    assert!(error.to_string().contains("relationships"), "{error}");
    assert!(matches!(
        edit.add_equation(0, EquationSource::Latex("x"), Marker::cell(0, 0), Some(0.5)),
        Err(Error::Invalid(_))
    ));
    assert!(
        edit.add_equation(0, EquationSource::Latex("x"), Marker::cell(20_000, 0), None)
            .is_err()
    );
    assert!(
        edit.add_equation(9, EquationSource::Latex("x"), Marker::cell(0, 0), None)
            .is_err()
    );
    assert!(edit.commit()?.patch().is_noop());
    Ok(())
}

#[test]
fn hostile_latex_does_not_panic() -> Result<()> {
    let package = Package::new()?;
    let deep = format!("{}x{}", r"\frac{".repeat(5_000), "}{y}".repeat(5_000));
    let mut edit = package.edit_drawings()?;
    assert!(
        edit.add_equation(0, EquationSource::Latex(&deep), Marker::cell(0, 0), None)
            .is_err()
    );
    let wide = "x+".repeat(2_000) + "x";
    // A long equation is drawn coarser, not refused.
    edit.add_equation(0, EquationSource::Latex(&wide), Marker::cell(0, 0), None)?;
    Ok(())
}
