//! Equation shapes: add, replace and remove, on PowerPoint's own files and
//! on hand-made slides.

use std::sync::Arc;

use litchi_core::Position;
use litchi_opc::PackURI;

use crate::shape::math::{EquationSource, equations};
use crate::{Error, Package, Result};

const HANDOUT: &[u8] = include_bytes!(
    "../../../../test-data/microsoft-office/powerpoint/powerpoint-fonts-math-handout.pptx"
);
const IMPORT_MCE: &[u8] =
    include_bytes!("../../../../test-data/libreoffice-core/oox/qa/unit/data/import-mce.pptx");
const IMAGE_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

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

fn parts(package: &Package) -> Vec<(String, Vec<u8>)> {
    let opc = package.opc().unwrap();
    let names: Vec<_> = opc
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

fn image_targets(package: &Package, slide: &str) -> Vec<(String, String)> {
    let opc = package.opc().unwrap();
    let part = opc.get_part(&part_name(slide)).unwrap();
    let mut out: Vec<_> = part
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == IMAGE_RELATIONSHIP)
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

/// The package after a save and reopen.
fn reopen(package: &Package) -> Package {
    Package::from_bytes(&package.to_bytes().unwrap()).unwrap()
}

fn compact(latex: &str) -> String {
    latex.chars().filter(|ch| !ch.is_whitespace()).collect()
}

fn latex_of(xml: &[u8]) -> Vec<String> {
    equations(xml)
        .unwrap()
        .iter()
        .map(|equation| compact(&equation.latex().unwrap()))
        .collect()
}

/// The root element of a slide: a commit drops the line break PowerPoint
/// writes after the XML declaration.
fn root(xml: &[u8]) -> &[u8] {
    &xml[find(xml, b"<p:sld").expect("a slide root")..]
}

/// `before` with `range` replaced by `with`.
fn spliced(before: &[u8], range: std::ops::Range<usize>, with: &[u8]) -> Vec<u8> {
    let mut out = before[..range.start].to_vec();
    out.extend_from_slice(with);
    out.extend_from_slice(&before[range.end..]);
    out
}

fn find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

/// The package with `slide`'s XML replaced by `xml`.
fn with_slide(source: &[u8], slide: &str, xml: &str) -> Package {
    let mut package = Package::from_bytes(source).unwrap();
    Arc::make_mut(&mut package.opc)
        .get_part_mut(&part_name(slide))
        .unwrap()
        .set_blob(xml.as_bytes().to_vec());
    reopen(&package)
}

#[test]
fn a_powerpoint_equation_is_replaced_and_only_its_markup_changes() -> Result<()> {
    let package = Package::from_bytes(HANDOUT)?;
    let before = blob(&package, "/ppt/slides/slide2.xml");
    let site = crate::shape::math::sites::scan(&before)?.0[0].clone();
    assert!(has_part(&package, "/ppt/media/image1.png"));

    let mut edit = package.edit()?;
    let written = edit
        .set_equation(1usize, 0, EquationSource::Latex(r"\frac{a}{b}"))?
        .expect("the equation changed");
    assert_eq!(written.shape_id, 3);
    assert!(written.losses.is_empty(), "{:?}", written.losses);
    let after = reopen(&edit.commit()?.into_package());
    let xml = blob(&after, "/ppt/slides/slide2.xml");

    // The equation reads back through the slide reader and the converter.
    assert_eq!(latex_of(&xml), [r"\frac{a}{b}"]);
    let omml = equations(&xml)?[0].omml().to_owned();
    assert!(omml.starts_with("<m:oMath "), "{omml}");
    // New runs carry the replaced equation's first run properties.
    assert!(
        omml.contains(r#"<a:rPr lang="en-US" i="1" smtClean="0"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr>"#),
        "{omml}"
    );

    // Only the OMML, the picture reference and the picture's fill rectangle
    // changed.
    let new_site = crate::shape::math::sites::scan(&xml)?.0[0].clone();
    let rewritten = spliced(
        &before,
        site.content.clone(),
        &xml[new_site.content.clone()],
    );
    let old_embed = find(&rewritten, br#"r:embed="rId2""#).expect("old picture reference");
    let new_embed = image_targets(&after, "/ppt/slides/slide2.xml")[0].0.clone();
    let rewritten = spliced(
        &rewritten,
        old_embed + 9..old_embed + 13,
        new_embed.as_bytes(),
    );
    let old_fill: &[u8] = br#"<a:stretch><a:fillRect l="-1086"/></a:stretch>"#;
    let old_fill_at = find(&rewritten, old_fill).expect("old fill");
    let new_fill_at = find(&xml, b"<a:stretch>").expect("new fill");
    let new_fill_end = new_fill_at + find(&xml[new_fill_at..], b"</a:stretch>").unwrap() + 12;
    let new_fill = &xml[new_fill_at..new_fill_end];
    let expected = spliced(
        &rewritten,
        old_fill_at..old_fill_at + old_fill.len(),
        new_fill,
    );
    assert_eq!(root(&xml), root(&expected));
    // The picture covers the layout placeholder's frame (the slide's
    // placeholder has none of its own), so it keeps the equation's
    // proportions when a reader stretches it over the frame.
    assert_eq!(new_fill, br#"<a:stretch><a:fillRect/></a:stretch>"#);
    let png = blob(
        &after,
        &image_targets(&after, "/ppt/slides/slide2.xml")[0].1,
    );
    let pixels = |at: usize| f64::from(u32::from_be_bytes(png[at..at + 4].try_into().unwrap()));
    let ratio = (pixels(16) / pixels(20)) / (10_515_600.0 / 4_351_338.0);
    assert!((ratio - 1.0).abs() < 0.01, "{ratio}");

    // The old picture is gone, the new one is a PNG of the equation, and
    // every other part is byte-identical.
    let targets = image_targets(&after, "/ppt/slides/slide2.xml");
    assert_eq!(targets.len(), 1, "{targets:?}");
    assert!(!has_part(&after, "/ppt/media/image1.png"));
    let png = blob(&after, &targets[0].1);
    assert!(png.starts_with(b"\x89PNG\r\n\x1a\n"));
    let untouched = |package: &Package| -> Vec<(String, Vec<u8>)> {
        parts(package)
            .into_iter()
            .filter(|(name, _)| {
                !name.starts_with("/ppt/media/")
                    && name != "/ppt/slides/slide2.xml"
                    && name != "/docProps/app.xml"
            })
            .collect()
    };
    assert_eq!(untouched(&after), untouched(&Package::from_bytes(HANDOUT)?));
    assert_eq!(
        blob(&after, "/ppt/media/image2.png"),
        blob(&Package::from_bytes(HANDOUT)?, "/ppt/media/image2.png")
    );
    Ok(())
}

#[test]
fn raw_omml_is_written_as_given_and_an_equal_equation_is_a_no_op() -> Result<()> {
    let package = Package::from_bytes(HANDOUT)?;
    let before = blob(&package, "/ppt/slides/slide3.xml");
    let current = equations(&before)?[0].omml().to_owned();
    let mut edit = package.edit()?;
    assert_eq!(
        edit.set_equation(2usize, 0, EquationSource::Omml(&current))?,
        None
    );
    // Undeclared conventional prefixes are declared on the root.
    let omml = "<m:oMath><m:r><a:rPr lang=\"en-US\" i=\"1\"/><m:t>y</m:t></m:r></m:oMath>";
    edit.set_equation(2usize, 0, EquationSource::Omml(omml))?;
    let after = reopen(&edit.commit()?.into_package());
    let xml = blob(&after, "/ppt/slides/slide3.xml");
    let written = equations(&xml)?[0].omml().to_owned();
    assert_eq!(
        written,
        "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><a:rPr lang=\"en-US\" i=\"1\"/><m:t>y</m:t></m:r></m:oMath>"
    );
    assert_eq!(latex_of(&xml), ["y"]);
    Ok(())
}

#[test]
fn an_autofit_text_box_takes_the_new_size_in_both_branches() -> Result<()> {
    let package = Package::from_bytes(IMPORT_MCE)?;
    let before = blob(&package, "/ppt/slides/slide1.xml");
    let mut edit = package.edit()?;
    edit.set_equation(
        0usize,
        0,
        EquationSource::Latex(r"\sum_{k=1}^{n} k^2 = \frac{n(n+1)(2n+1)}{6}"),
    )?;
    let after = reopen(&edit.commit()?.into_package());
    let xml = blob(&after, "/ppt/slides/slide1.xml");
    let text = std::str::from_utf8(&xml).unwrap();
    // The display paragraph and its justification stay.
    let omml = equations(&xml)?[0].omml().to_owned();
    assert!(omml.starts_with("<m:oMathPara "), "{omml}");
    assert!(omml.contains(r#"<m:jc m:val="centerGroup"/>"#), "{omml}");
    // Both frames have the same, new extent; the offset is kept.
    let block = &text[text.find("<mc:AlternateContent").unwrap()..];
    let extents: Vec<_> = block
        .match_indices("<a:ext cx=")
        .map(|(at, _)| &block[at..at + block[at..].find("/>").unwrap()])
        .collect();
    assert_eq!(extents.len(), 2, "{extents:?}");
    assert_eq!(extents[0], extents[1]);
    assert_ne!(extents[0], r#"<a:ext cx="802912" cy="618311""#);
    assert_eq!(
        text.matches(r#"<a:off x="4140613" y="3291840"/>"#).count(),
        2
    );
    assert!(
        text.contains("<a:stretch><a:fillRect/></a:stretch>"),
        "{text}"
    );
    // Outside the block, the slide is unchanged.
    let before = root(&before).to_vec();
    let xml = root(&xml).to_vec();
    let start = find(&before, b"<mc:AlternateContent").unwrap();
    assert_eq!(xml[..start], before[..start]);
    let tail = before.len() - find(&before, b"</mc:AlternateContent>").unwrap();
    assert_eq!(xml[xml.len() - tail..], before[before.len() - tail..]);
    Ok(())
}

#[test]
fn a_powerpoint_equation_shape_is_removed_with_its_picture() -> Result<()> {
    let package = Package::from_bytes(HANDOUT)?;
    let before = blob(&package, "/ppt/slides/slide3.xml");
    let mut edit = package.edit()?;
    assert_eq!(edit.remove_equation(2usize, 0)?, 3);
    let after = reopen(&edit.commit()?.into_package());
    let xml = blob(&after, "/ppt/slides/slide3.xml");
    let start = find(&before, b"<mc:AlternateContent").unwrap();
    let end = find(&before, b"</mc:AlternateContent>").unwrap() + "</mc:AlternateContent>".len();
    assert_eq!(root(&xml), root(&spliced(&before, start..end, b"")));
    assert!(equations(&xml)?.is_empty());
    assert!(image_targets(&after, "/ppt/slides/slide3.xml").is_empty());
    assert!(!has_part(&after, "/ppt/media/image2.png"));
    assert!(has_part(&after, "/ppt/media/image1.png"));
    assert!(has_part(&after, "/ppt/media/image3.png"));
    // The other equation slides are untouched.
    for slide in ["/ppt/slides/slide2.xml", "/ppt/slides/slide4.xml"] {
        assert_eq!(
            blob(&after, slide),
            blob(&Package::from_bytes(HANDOUT)?, slide)
        );
    }
    Ok(())
}

#[test]
fn a_new_equation_is_added_in_powerpoints_form_and_round_trips() -> Result<()> {
    let package = Package::from_bytes(HANDOUT)?;
    let mut edit = package.edit()?;
    let first = edit.add_equation(
        0usize,
        EquationSource::Latex(r"E = mc^2"),
        (914_400, 914_400),
        None,
    )?;
    // A second equation on the same, now markup-compatible, slide.
    let second = edit.add_equation(
        0usize,
        EquationSource::Latex(r"\sqrt{x}"),
        (914_400, 2_743_200),
        Some(28.0),
    )?;
    assert!(second.shape_id > first.shape_id, "{first:?} {second:?}");
    let after = reopen(&edit.commit()?.into_package());
    let xml = blob(&after, "/ppt/slides/slide1.xml");
    let text = std::str::from_utf8(&xml).unwrap();
    assert_eq!(latex_of(&xml), ["E=m{c}^{2}", r"\sqrt{x}"]);
    let found = equations(&xml)?;
    assert_eq!(found[0].shape_id(), Some(first.shape_id));
    assert_eq!(found[1].shape_id(), Some(second.shape_id));
    for marker in [
        r#"<mc:Choice Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id=""#,
        r#"<p:cNvSpPr txBox="1"/>"#,
        r#"<a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr>"#,
        r#"<a:p><a14:m><m:oMathPara "#,
        r#"<m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>"#,
        r#"<a:rPr lang="en-US" i="1"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr>"#,
        r#"<a:rPr lang="en-US" sz="2800" i="1">"#,
        r#"noTextEdit="1"/></p:cNvSpPr>"#,
        r#"<a:blipFill rotWithShape="1"><a:blip r:embed="#,
        "<a:r><a:rPr lang=\"en-US\"><a:noFill/></a:rPr><a:t>\u{a0}</a:t></a:r>",
    ] {
        assert!(text.contains(marker), "{marker} in {text}");
    }
    // The scene, which reads the fallbacks, sees two picture-filled shapes.
    let scene = crate::shape::Scene::read(&xml, crate::shape::Limits::DEFAULT)?;
    assert!(scene.iter().any(|shape| shape.id() == Some(first.shape_id)));
    assert!(
        scene
            .iter()
            .any(|shape| shape.id() == Some(second.shape_id))
    );
    let targets = image_targets(&after, "/ppt/slides/slide1.xml");
    assert_eq!(targets.len(), 2, "{targets:?}");
    for (_, target) in &targets {
        assert!(blob(&after, target).starts_with(b"\x89PNG"));
        assert_eq!(
            after.opc()?.get_part(&part_name(target))?.content_type(),
            "image/png"
        );
    }
    // Each shape has the same frame in both branches.
    let blocks: Vec<_> = text.split("<mc:AlternateContent").skip(1).collect();
    assert_eq!(blocks.len(), 2);
    for block in blocks {
        let extents: Vec<_> = block
            .match_indices("<a:ext cx=")
            .map(|(at, _)| &block[at..at + block[at..].find("/>").unwrap()])
            .collect();
        assert_eq!(extents.len(), 2, "{extents:?}");
        assert_eq!(extents[0], extents[1]);
    }

    // Replace, then remove both: no picture is left behind.
    let mut edit = after.edit()?;
    edit.set_equation(0usize, 1, EquationSource::Latex(r"\sqrt[3]{x}"))?;
    let after = reopen(&edit.commit()?.into_package());
    assert_eq!(
        latex_of(&blob(&after, "/ppt/slides/slide1.xml")),
        ["E=m{c}^{2}", r"\sqrt[3]{x}"]
    );
    let mut edit = after.edit()?;
    edit.remove_equation(0usize, 1)?;
    edit.remove_equation(0usize, 0)?;
    let after = reopen(&edit.commit()?.into_package());
    assert!(equations(&blob(&after, "/ppt/slides/slide1.xml"))?.is_empty());
    assert!(image_targets(&after, "/ppt/slides/slide1.xml").is_empty());
    let media: Vec<_> = parts(&after)
        .into_iter()
        .map(|(name, _)| name)
        .filter(|name| name.starts_with("/ppt/media/"))
        .collect();
    assert_eq!(
        media,
        [
            "/ppt/media/image1.png",
            "/ppt/media/image2.png",
            "/ppt/media/image3.png"
        ]
    );
    Ok(())
}

#[test]
fn a_new_presentation_takes_equations() -> Result<()> {
    let source = Package::new()?;
    let mut edit = source.edit()?;
    edit.insert_slide(Position::new(0), "Blank")?;
    let written = edit.add_equation(
        0usize,
        EquationSource::Omml("<m:oMath><m:r><m:t>z</m:t></m:r></m:oMath>"),
        (0, 0),
        None,
    )?;
    let after = reopen(&edit.commit()?.into_package());
    let presentation = after.presentation()?;
    let slide = presentation.slide(0)?.expect("slide");
    let found = slide.equations()?;
    assert_eq!(found.len(), 1);
    assert_eq!(found[0].shape_id(), Some(written.shape_id));
    assert_eq!(found[0].latex().unwrap(), "z");
    Ok(())
}

fn mixed_slide(body: &str) -> String {
    format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
            r#"<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"#
        ),
        body
    )
}

const MATH: &str = r#"<a14:m><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath></a14:m>"#;

fn block(paragraphs: &str) -> String {
    format!(
        concat!(
            r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">"#,
            r#"<mc:Choice Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id="5" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{}</p:txBody></p:sp></mc:Choice>"#,
            r#"<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="5" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>"#
        ),
        paragraphs
    )
}

fn refusal(error: &Error) -> bool {
    matches!(error, Error::Refused(_))
}

/// The slide XML of slide 2 of `package`.
fn slide2(package: &Package) -> String {
    String::from_utf8(blob(package, "/ppt/slides/slide2.xml")).unwrap()
}

/// The Fallback branch of the first block of `xml`.
fn fallback(xml: &str) -> &str {
    let start = xml.find("<mc:Fallback").expect("a Fallback");
    &xml[start..start + xml[start..].find("</mc:Fallback>").expect("its end")]
}

#[test]
fn an_equation_outside_a_block_is_refused() -> Result<()> {
    let body = format!(
        r#"<p:sp xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><p:nvSpPr><p:cNvPr id="5" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>{MATH}</a:p></p:txBody></p:sp>"#
    );
    let package = with_slide(HANDOUT, "/ppt/slides/slide2.xml", &mixed_slide(&body));
    let mut edit = package.edit()?;
    let error = edit
        .set_equation(1usize, 0, EquationSource::Latex("y"))
        .unwrap_err();
    assert!(refusal(&error), "{error:?}");
    let error = edit.remove_equation(1usize, 0).unwrap_err();
    assert!(refusal(&error), "{error:?}");
    assert!(edit.commit()?.patch().is_noop());
    Ok(())
}

#[test]
fn an_equation_beside_text_is_replaced_with_a_text_fallback() -> Result<()> {
    let text = format!(
        "<a:p><a:r><a:rPr lang=\"en-US\"/><a:t>where </a:t></a:r>{MATH}<a:r><a:t> is small</a:t></a:r></a:p>"
    );
    let package = with_slide(
        HANDOUT,
        "/ppt/slides/slide2.xml",
        &mixed_slide(&block(&text)),
    );
    let mut edit = package.edit()?;
    let written = edit
        .set_equation(1usize, 0, EquationSource::Latex(r"\frac{a}{b}"))?
        .expect("changed");
    assert_eq!((written.shape_id, written.index), (5, 0));
    let after = reopen(&edit.commit()?.into_package());
    let xml = slide2(&after);
    assert_eq!(latex_of(xml.as_bytes()), [r"\frac{a}{b}"]);
    // The Fallback is the shape's text with the equation in linear form,
    // and the old picture is gone.
    let fallback = fallback(&xml);
    assert!(fallback.contains("<a:t>where </a:t>"), "{fallback}");
    assert!(fallback.contains("<a:t>a/b</a:t>"), "{fallback}");
    assert!(fallback.contains("<a:t> is small</a:t>"), "{fallback}");
    assert!(
        !fallback.contains("blipFill") && !fallback.contains("a14:m"),
        "{fallback}"
    );
    assert!(image_targets(&after, "/ppt/slides/slide2.xml").is_empty());
    assert!(!has_part(&after, "/ppt/media/image1.png"));
    // The scene, which reads the Fallback, sees the text.
    let scene = crate::shape::Scene::read(xml.as_bytes(), crate::shape::Limits::DEFAULT)?;
    let shape = scene.iter().find(|shape| shape.id() == Some(5)).unwrap();
    assert_eq!(shape.common().text(), Some("where a/b is small"));
    Ok(())
}

#[test]
fn equations_of_a_mixed_shape_are_removed_one_by_one() -> Result<()> {
    let text = format!("<a:p><a:r><a:t>Let </a:t></a:r>{MATH}</a:p><a:p>{MATH}</a:p>");
    let package = with_slide(
        HANDOUT,
        "/ppt/slides/slide2.xml",
        &mixed_slide(&block(&text)),
    );
    let mut edit = package.edit()?;
    assert_eq!(edit.remove_equation(1usize, 1)?, 5);
    let after = reopen(&edit.commit()?.into_package());
    let xml = slide2(&after);
    assert_eq!(latex_of(xml.as_bytes()), ["x"]);
    assert!(fallback(&xml).contains("<a:t>Let </a:t>"));
    // The last equation leaves the plain shape behind.
    let mut edit = after.edit()?;
    edit.remove_equation(1usize, 0)?;
    let after = reopen(&edit.commit()?.into_package());
    let xml = slide2(&after);
    assert!(!xml.contains("AlternateContent"), "{xml}");
    assert!(xml.contains(r#"<p:cNvPr id="5" name="T"/>"#), "{xml}");
    assert!(xml.contains("<a:t>Let </a:t>"), "{xml}");
    assert!(equations(xml.as_bytes())?.is_empty());
    Ok(())
}

#[test]
fn an_equation_is_added_to_a_shapes_text() -> Result<()> {
    // A plain text box gets wrapped in PowerPoint's block.
    let plain = r#"<p:sp><p:nvSpPr><p:cNvPr id="7" name="Body"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Area: </a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#;
    let package = with_slide(HANDOUT, "/ppt/slides/slide2.xml", &mixed_slide(plain));
    let mut edit = package.edit()?;
    let written = edit.add_equation_to_shape(1usize, 7, EquationSource::Latex(r"\pi r^2"))?;
    assert_eq!((written.shape_id, written.index), (7, 0));
    let after = reopen(&edit.commit()?.into_package());
    let xml = slide2(&after);
    assert_eq!(latex_of(xml.as_bytes()), [r"\pir^{2}"]);
    assert!(
        xml.contains(r#"<mc:Choice Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id="7" name="Body"/>"#),
        "{xml}"
    );
    // The equation sits before the paragraph's end properties.
    assert!(xml.contains("</a14:m><a:endParaRPr"), "{xml}");
    let fallback = fallback(&xml);
    assert!(fallback.contains("<a:t>Area: </a:t>"), "{fallback}");
    assert!(fallback.contains(">πr^2</a:t>"), "{fallback}");
    // A second one joins the same block.
    let mut edit = after.edit()?;
    let written = edit.add_equation_to_shape(1usize, 7, EquationSource::Latex("x"))?;
    assert_eq!(written.index, 1);
    let after = reopen(&edit.commit()?.into_package());
    let xml = slide2(&after);
    assert_eq!(xml.matches("<mc:AlternateContent").count(), 1);
    assert_eq!(equations(xml.as_bytes())?.len(), 2);
    // Missing shapes and shapes without text are refused.
    let mut edit = after.edit()?;
    assert!(
        edit.add_equation_to_shape(1usize, 99, EquationSource::Latex("x"))
            .is_err()
    );
    Ok(())
}

#[test]
fn an_equation_with_an_empty_fallback_is_removed() -> Result<()> {
    // LibreOffice writes an empty Fallback for its equation shapes.
    let body = format!(
        concat!(
            r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">"#,
            r#"<mc:Choice Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id="5" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>{}</a:p></p:txBody></p:sp></mc:Choice>"#,
            r#"<mc:Fallback/></mc:AlternateContent>"#
        ),
        MATH
    );
    let package = with_slide(HANDOUT, "/ppt/slides/slide2.xml", &mixed_slide(&body));
    let mut edit = package.edit()?;
    // A replacement gets a text Fallback, as the shape has no picture.
    edit.set_equation(1usize, 0, EquationSource::Latex("y"))?;
    let staged = reopen(&edit.commit()?.into_package());
    let xml = slide2(&staged);
    assert!(fallback(&xml).contains(">y</a:t>"), "{xml}");
    let mut edit = package.edit()?;
    assert_eq!(edit.remove_equation(1usize, 0)?, 5);
    let after = reopen(&edit.commit()?.into_package());
    assert!(!slide2(&after).contains("AlternateContent"));
    Ok(())
}

#[test]
fn untouched_equations_keep_their_space_only_runs() -> Result<()> {
    // A space-only math run in one equation, a space-only text run in
    // another shape: an edit of the slide keeps both.
    let spaced = r#"<a14:m><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>where</m:t></m:r><m:r><m:t> </m:t></m:r><m:r><m:t>a</m:t></m:r></m:oMath></a14:m>"#;
    let text = r#"<p:sp><p:nvSpPr><p:cNvPr id="9" name="U"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>a</a:t></a:r><a:r><a:t> </a:t></a:r><a:r><a:t>b</a:t></a:r></a:p></p:txBody></p:sp>"#;
    let body = format!("{}{text}", block(&format!("<a:p>{spaced}</a:p>")));
    let package = with_slide(HANDOUT, "/ppt/slides/slide2.xml", &mixed_slide(&body));
    let mut edit = package.edit()?;
    edit.add_equation(1usize, EquationSource::Latex("x"), (0, 0), None)?;
    let after = reopen(&edit.commit()?.into_package());
    let xml = slide2(&after);
    assert!(xml.contains("<m:r><m:t> </m:t></m:r>"), "{xml}");
    assert!(xml.contains("<a:r><a:t> </a:t></a:r>"), "{xml}");
    Ok(())
}

#[test]
fn caller_omml_cannot_refer_to_relationships() -> Result<()> {
    let package = Package::from_bytes(HANDOUT)?;
    let mut edit = package.edit()?;
    for omml in [
        r#"<m:oMath><m:r><a:rPr><a:hlinkClick r:id="rId77"/></a:rPr><m:t>x</m:t></m:r></m:oMath>"#,
        r#"<m:oMath xmlns:q="http://purl.oclc.org/ooxml/officeDocument/relationships"><m:r><a:rPr><a:hlinkClick q:id="rId1"/></a:rPr><m:t>x</m:t></m:r></m:oMath>"#,
    ] {
        assert!(
            edit.set_equation(1usize, 0, EquationSource::Omml(omml))
                .is_err(),
            "{omml}"
        );
        assert!(
            edit.add_equation(0usize, EquationSource::Omml(omml), (0, 0), None)
                .is_err(),
            "{omml}"
        );
    }
    assert!(edit.commit()?.patch().is_noop());
    Ok(())
}

#[test]
fn a_long_equation_is_drawn_coarser_not_refused() -> Result<()> {
    let sum = (1..=40)
        .map(|k| format!("x_{{{k}}}^{{2}}"))
        .collect::<Vec<_>>()
        .join("+");
    let package = Package::from_bytes(HANDOUT)?;
    let mut edit = package.edit()?;
    edit.add_equation(0usize, EquationSource::Latex(&sum), (0, 0), None)?;
    let after = reopen(&edit.commit()?.into_package());
    assert_eq!(equations(&blob(&after, "/ppt/slides/slide1.xml"))?.len(), 1);
    Ok(())
}

#[test]
fn a_picture_follows_the_paragraph_alignment() -> Result<()> {
    // A centred inline equation in a framed, bottom-anchored box: the
    // picture's opaque pixels sit in the middle and at the bottom.
    let body = format!(
        concat!(
            r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">"#,
            r#"<mc:Choice Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id="5" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="3000000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="b"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>{}</a:p></p:txBody></p:sp></mc:Choice>"#,
            r#"<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="5" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="3000000"/></a:xfrm><a:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>"#
        ),
        MATH
    );
    let package = with_slide(HANDOUT, "/ppt/slides/slide2.xml", &mixed_slide(&body));
    let mut edit = package.edit()?;
    edit.set_equation(1usize, 0, EquationSource::Latex("y"))?;
    let after = reopen(&edit.commit()?.into_package());
    let target = image_targets(&after, "/ppt/slides/slide2.xml")[0].1.clone();
    let png = blob(&after, &target);
    let (width, height, alpha) = opaque_box(&png);
    let ((left, top), (right, bottom)) = alpha;
    let centre = f64::from(left + right) / 2.0 / f64::from(width);
    assert!((centre - 0.5).abs() < 0.05, "{centre}");
    assert!(
        f64::from(bottom) > f64::from(height) * 0.9,
        "{bottom} of {height}"
    );
    assert!(
        f64::from(top) > f64::from(height) * 0.5,
        "{top} of {height}"
    );
    Ok(())
}

/// The size of a PNG and the box of its non-transparent pixels.
fn opaque_box(png: &[u8]) -> (u32, u32, ((u32, u32), (u32, u32))) {
    let picture = tiny_skia::Pixmap::decode_png(png).unwrap();
    let (width, height) = (picture.width(), picture.height());
    let (mut left, mut top, mut right, mut bottom) = (width, height, 0, 0);
    for (index, pixel) in picture.pixels().iter().enumerate() {
        if pixel.alpha() > 0 {
            let index = u32::try_from(index).unwrap();
            let (x, y) = (index % width, index / width);
            left = left.min(x);
            right = right.max(x);
            top = top.min(y);
            bottom = bottom.max(y);
        }
    }
    (width, height, ((left, top), (right, bottom)))
}

#[test]
fn bad_input_is_a_typed_error_and_stages_nothing() -> Result<()> {
    let package = Package::from_bytes(HANDOUT)?;
    let mut edit = package.edit()?;
    assert!(matches!(
        edit.set_equation(1usize, 5, EquationSource::Latex("y")),
        Err(Error::IndexOutOfBounds { index: 5, len: 1 })
    ));
    assert!(matches!(
        edit.set_equation(0usize, 0, EquationSource::Latex("y")),
        Err(Error::IndexOutOfBounds { index: 0, len: 0 })
    ));
    assert!(matches!(
        edit.set_equation(1usize, 0, EquationSource::Latex(r"\frac{")),
        Err(Error::Equation(litchi_equation::Error::Malformed { .. }))
    ));
    for omml in [
        "<m:oMath><m:r>",
        "<w:p/>",
        "<m:oMath/><m:oMath/>",
        "<m:oMath><q:r/></m:oMath>",
        "text",
        r#"<x:oMath xmlns:x="urn:other"/>"#,
    ] {
        assert!(
            edit.set_equation(1usize, 0, EquationSource::Omml(omml))
                .is_err(),
            "{omml}"
        );
        assert!(
            edit.add_equation(0usize, EquationSource::Omml(omml), (0, 0), None)
                .is_err(),
            "{omml}"
        );
    }
    for size in [0.0, f64::NAN, 5000.0] {
        assert!(
            edit.add_equation(0usize, EquationSource::Latex("x"), (0, 0), Some(size))
                .is_err()
        );
    }
    assert!(
        edit.add_equation(0usize, EquationSource::Latex("x"), (i64::MAX, 0), None)
            .is_err()
    );
    // Hostile LaTeX: deep nesting and a picture beyond the pixel limit.
    let deep = format!("{}x{}", "{".repeat(10_000), "}".repeat(10_000));
    assert!(
        edit.set_equation(1usize, 0, EquationSource::Latex(&deep))
            .is_err()
    );
    assert!(edit.commit()?.patch().is_noop());
    Ok(())
}

#[test]
fn an_unwrapped_shape_keeps_the_declarations_it_uses() -> Result<()> {
    // LibreOffice declares a14 on the Choice; the shape's properties use it.
    let body = format!(
        concat!(
            r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">"#,
            r#"<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id="3" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#,
            r#"<p:spPr><a:extLst><a:ext uri="{{X}}"><a14:hiddenFill><a:noFill/></a14:hiddenFill></a:ext></a:extLst></p:spPr>"#,
            r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Let </a:t></a:r>{}</a:p></p:txBody></p:sp></mc:Choice>"#,
            r#"<mc:Fallback/></mc:AlternateContent>"#
        ),
        MATH
    );
    let package = with_slide(HANDOUT, "/ppt/slides/slide2.xml", &mixed_slide(&body));
    let mut edit = package.edit()?;
    // Another edit in the same transaction must survive.
    edit.add_equation(0usize, EquationSource::Latex("z"), (0, 0), None)?;
    assert_eq!(edit.remove_equation(1usize, 0)?, 3);
    let after = reopen(&edit.commit()?.into_package());
    let xml = slide2(&after);
    assert!(!xml.contains("AlternateContent"), "{xml}");
    assert!(
        xml.contains(r#"<p:sp xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main""#),
        "{xml}"
    );
    assert!(xml.contains("<a14:hiddenFill>"), "{xml}");
    assert_eq!(equations(&blob(&after, "/ppt/slides/slide1.xml"))?.len(), 1);
    Ok(())
}

#[test]
fn an_added_equation_takes_the_paragraph_text_size_and_fills_an_empty_paragraph() -> Result<()> {
    let sized = r#"<p:sp><p:nvSpPr><p:cNvPr id="7" name="A"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2400"/><a:t>Area </a:t></a:r></a:p></p:txBody></p:sp>"#;
    let empty = r#"<p:sp><p:nvSpPr><p:cNvPr id="8" name="B"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>first</a:t></a:r></a:p><a:p/></p:txBody></p:sp>"#;
    let package = with_slide(
        HANDOUT,
        "/ppt/slides/slide2.xml",
        &mixed_slide(&format!("{sized}{empty}")),
    );
    let mut edit = package.edit()?;
    edit.add_equation_to_shape(1usize, 7, EquationSource::Latex("x"))?;
    edit.add_equation_to_shape(1usize, 8, EquationSource::Latex("y"))?;
    let after = reopen(&edit.commit()?.into_package());
    let xml = slide2(&after);
    let found = equations(xml.as_bytes())?;
    assert_eq!(found.len(), 2);
    assert!(
        found[0].omml().contains(r#"sz="2400""#),
        "{}",
        found[0].omml()
    );
    // The math run's size reaches the text fallback too.
    let fallback = fallback(&xml);
    assert!(fallback.contains(r#"sz="2400""#), "{fallback}");
    // The equation went into the empty last paragraph, not the first.
    assert!(
        xml.contains("<a:t>first</a:t></a:r></a:p><a:p><a14:m"),
        "{xml}"
    );
    Ok(())
}

#[test]
fn an_equation_is_not_added_to_a_shape_without_text() -> Result<()> {
    // A table frame (and any shape that is not a text shape) is refused,
    // not reported missing; so is a shape without a text body.
    let frame = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid><a:gridCol w="10"/></a:tblGrid><a:tr h="10"><a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
    let bare = r#"<p:sp><p:nvSpPr><p:cNvPr id="10" name="R"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>"#;
    let package = with_slide(
        HANDOUT,
        "/ppt/slides/slide2.xml",
        &mixed_slide(&format!("{frame}{bare}")),
    );
    let mut edit = package.edit()?;
    for id in [9, 10, 99] {
        let error = edit
            .add_equation_to_shape(1usize, id, EquationSource::Latex("x"))
            .unwrap_err();
        assert!(refusal(&error), "{id}: {error:?}");
    }
    assert!(edit.commit()?.patch().is_noop());
    Ok(())
}
