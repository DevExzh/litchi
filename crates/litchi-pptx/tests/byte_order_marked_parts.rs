#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "differential assertions intentionally panic on fixture failures"
)]

//! Change 0765: a part that begins with a UTF-8 byte-order mark reads, edits
//! and saves exactly like the same part without one.
//!
//! quick-xml drops a leading mark before its first event without counting it
//! in its positions, so every span this crate took from a reader was three
//! bytes early in a marked part. Each test below runs one scenario twice —
//! on a package and on its twin whose XML members carry a mark — and
//! requires identical observations and an output that differs only by marks:
//! every member of the marked output is the unmarked output's member, or that
//! member behind one mark, and a mark appears only on a member whose input
//! had one. Members an edit never touched keep their mark byte for byte.

use std::collections::{BTreeMap, BTreeSet};

use litchi_pptx::Package;
use litchi_pptx::shape::Scene;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const BOM: &[u8] = b"\xEF\xBB\xBF";
const SHAPE_TAGS: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf103477.pptx"
);
const MARKED_MASTER: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/poi/test-data/slideshow/bug65551.pptx"
);

fn is_xml_member(name: &str) -> bool {
    name.ends_with(".xml") || name.ends_with(".rels")
}

fn members(archive: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let reader = ArchiveReader::new(archive).unwrap();
    let names: Vec<String> = reader.file_names().map(str::to_owned).collect();
    names
        .into_iter()
        .map(|name| {
            let bytes = reader.read(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

/// Indent every start tag that directly follows a tag: whitespace-only text
/// between elements, never inside an element that holds only text.
fn indent(xml: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(xml.len() + xml.len() / 8);
    for (index, byte) in xml.iter().enumerate() {
        if *byte == b'<'
            && index > 0
            && xml[index - 1] == b'>'
            && xml
                .get(index + 1)
                .is_some_and(|next| next.is_ascii_alphabetic())
        {
            output.extend_from_slice(b"\n    ");
        }
        output.push(*byte);
    }
    output
}

/// Rebuild `archive` with every member `mark` selects behind one mark and
/// every member that already starts with one without it, XML members
/// indented when `indented`; returns the archive and its marked members.
fn rebuild(
    archive: &[u8],
    mark: &dyn Fn(&str) -> bool,
    indented: bool,
) -> (Vec<u8>, BTreeSet<String>) {
    let mut writer = StreamingArchiveWriter::new();
    let mut marked = BTreeSet::new();
    for (name, bytes) in members(archive) {
        let body = bytes.strip_prefix(BOM).unwrap_or(&bytes);
        let body = if indented && is_xml_member(&name) {
            indent(body)
        } else {
            body.to_vec()
        };
        let bytes = if mark(&name) {
            marked.insert(name.clone());
            [BOM, body.as_slice()].concat()
        } else {
            body
        };
        writer.write_stored(&name, &bytes).unwrap();
    }
    (writer.finish_to_bytes().unwrap(), marked)
}

/// Require the marked output to equal the unmarked output except for marks.
/// Returns the members that kept a mark.
fn assert_equal_but_marks(
    marked: &[u8],
    plain: &[u8],
    marked_input: &BTreeSet<String>,
) -> BTreeSet<String> {
    let marked = members(marked);
    let plain = members(plain);
    assert_eq!(
        marked.keys().collect::<Vec<_>>(),
        plain.keys().collect::<Vec<_>>(),
        "member lists differ"
    );
    let mut kept = BTreeSet::new();
    for (name, bytes) in &marked {
        let unmarked = &plain[name];
        assert!(
            !unmarked.starts_with(BOM),
            "{name}: unmarked output gained a mark"
        );
        if let Some(rest) = bytes.strip_prefix(BOM) {
            assert!(
                marked_input.contains(name),
                "{name}: a mark appeared on a member whose input had none"
            );
            assert_eq!(rest, unmarked.as_slice(), "{name}: differs beyond its mark");
            kept.insert(name.clone());
        } else {
            assert_eq!(bytes, unmarked, "{name}: differs");
        }
    }
    kept
}

/// Run `scenario` on the plain and the marked twin of `archive`, compact and
/// indented, compare the observations and the saved packages, and return the
/// members that kept a mark (compact run) with the marked input's members.
fn differential(
    archive: &[u8],
    mark: impl Fn(&str) -> bool,
    scenario: impl Fn(&mut Package) -> String,
) -> (BTreeSet<String>, BTreeSet<String>) {
    let mut result = (BTreeSet::new(), BTreeSet::new());
    for indented in [true, false] {
        let (plain, _) = rebuild(archive, &|_| false, indented);
        let (marked, marked_input) = rebuild(archive, &mark, indented);
        assert!(!marked_input.is_empty(), "the scenario must mark a member");
        let mut plain_package = Package::from_vec(plain).unwrap();
        let mut marked_package = Package::from_vec(marked).unwrap();
        let plain_observed = scenario(&mut plain_package);
        let marked_observed = scenario(&mut marked_package);
        assert_eq!(marked_observed, plain_observed, "indented: {indented}");
        let plain_out = plain_package.to_bytes().unwrap();
        let marked_out = marked_package.to_bytes().unwrap();
        let kept = assert_equal_but_marks(&marked_out, &plain_out, &marked_input);
        result = (kept, marked_input);
    }
    result
}

fn text_deck() -> Vec<u8> {
    let mut package = Package::new().unwrap();
    {
        let presentation = package.presentation_mut().unwrap();
        for slide_index in 0..3_i64 {
            let slide = presentation.add_slide().unwrap();
            for shape in 0..3_i64 {
                slide.add_text_box(
                    &format!("slide {slide_index} shape {shape}"),
                    10 + shape * 1_000,
                    20,
                    300,
                    400,
                );
            }
        }
    }
    package.to_bytes().unwrap()
}

fn slide_texts(package: &Package) -> String {
    let presentation = package.presentation().unwrap();
    let mut observed = String::new();
    for slide in presentation.slides().unwrap() {
        let scene = slide.shapes().unwrap();
        for index in 0..scene.len() {
            let shape = scene.at(index).unwrap();
            observed.push_str(&format!(
                "{index}:{:?}:{:?};",
                shape.name(),
                shape.common().text()
            ));
        }
        observed.push('|');
    }
    observed
}

#[test]
fn a_fully_marked_deck_reads_like_the_unmarked_deck() {
    let deck = text_deck();
    let (kept, marked) = differential(&deck, is_xml_member, |package| slide_texts(package));
    // Saving an unedited package preserves every member, marks included.
    assert_eq!(kept, marked);
}

#[test]
fn scene_spans_are_byte_offsets_into_the_marked_slide() {
    let deck = text_deck();
    let slide = members(&deck)["ppt/slides/slide1.xml"].clone();
    assert!(!slide.starts_with(BOM));
    let marked = [BOM, slide.as_slice()].concat();
    let plain_scene = Scene::read(&slide).unwrap();
    let marked_scene = Scene::read(&marked).unwrap();
    assert_eq!(marked_scene.len(), plain_scene.len());
    assert!(!marked_scene.is_rewritten());
    assert_eq!(marked_scene.xml(), marked.as_slice());
    for index in 0..plain_scene.len() {
        let plain_shape = plain_scene.at(index).unwrap().common();
        let marked_shape = marked_scene.at(index).unwrap().common();
        let plain_span = plain_shape.span().unwrap();
        let marked_span = marked_shape.span().unwrap();
        // Spans address the raw part bytes, mark included.
        assert_eq!(marked_span.start(), plain_span.start() + 3);
        assert_eq!(marked_span.len(), plain_span.len());
        let range = marked_span.start() as usize..marked_span.end().unwrap() as usize;
        assert_eq!(&marked[range], marked_shape.xml().unwrap());
        assert_eq!(marked_shape.xml().unwrap(), plain_shape.xml().unwrap());
        assert_eq!(
            marked_shape.self_contained_xml().unwrap(),
            plain_shape.self_contained_xml().unwrap()
        );
        assert_eq!(marked_shape.text(), plain_shape.text());
    }
}

#[test]
fn opened_transaction_edits_of_marked_slides_match_unmarked_edits() {
    let deck = text_deck();
    let (kept, _) = differential(&deck, is_xml_member, |package| {
        let mut edit = package.opened_presentation_transaction().unwrap();
        assert!(
            edit.set_shape_text(0_usize, 1_usize, "edited & <ok>")
                .unwrap()
        );
        assert!(edit.remove_shape(1_usize, 0_usize).is_ok());
        edit.add_text_box(2_usize, "added", (5, 6, 7, 8)).unwrap();
        assert!(edit.move_slide(0, 2).unwrap());
        let commit = edit.commit().unwrap();
        package.apply_opened_presentation_commit(commit).unwrap();
        slide_texts(package)
    });
    // The edited slides are spliced and compacted from their marked sources,
    // so they keep the producer's mark like every untouched member.
    assert!(kept.contains("ppt/slides/slide1.xml"), "{kept:?}");
    assert!(kept.contains("ppt/slides/slide2.xml"), "{kept:?}");
    assert!(kept.contains("ppt/slides/slide3.xml"), "{kept:?}");
    assert!(kept.contains("ppt/presentation.xml"), "{kept:?}");
}

#[test]
fn batched_text_edits_of_a_marked_slide_match_unmarked_edits() {
    let deck = text_deck();
    let (kept, _) = differential(
        &deck,
        |name| name == "ppt/slides/slide2.xml",
        |package| {
            let mut edit = package.opened_presentation_transaction().unwrap();
            let replacements = [
                litchi_pptx::opened::ShapeTextReplacement::at(0, "first"),
                litchi_pptx::opened::ShapeTextReplacement::at(2, "third"),
            ];
            assert_eq!(edit.set_shape_texts(1_usize, &replacements).unwrap(), 2);
            let commit = edit.commit().unwrap();
            package.apply_opened_presentation_commit(commit).unwrap();
            slide_texts(package)
        },
    );
    assert_eq!(kept, BTreeSet::from(["ppt/slides/slide2.xml".to_owned()]));
}

#[test]
fn real_shape_tag_deck_marked_everywhere_reads_and_edits_like_the_original() {
    let original = std::fs::read(SHAPE_TAGS).unwrap();
    differential(&original, is_xml_member, |package| {
        let mut observed = slide_texts(package);
        let presentation = package.presentation().unwrap();
        let slides = presentation.slides().unwrap();
        let slide = &slides[0];
        let scene = slide.shapes().unwrap();
        for index in 0..scene.len() {
            observed.push_str(&format!("{:?};", slide.shape_tags(index).unwrap()));
        }
        drop(slides);
        let mut replacement = litchi_pptx::tag::List::new();
        replacement
            .add(litchi_pptx::tag::Tag::new("Reviewer", "Ada").unwrap())
            .unwrap();
        let old = package
            .put_shape_tags(0_usize, "Objekt 2", replacement)
            .unwrap();
        observed.push_str(&format!("{old:?}"));
        observed
    });
}

#[test]
fn a_naturally_marked_master_reads_and_saves_like_its_unmarked_twin() {
    let original = std::fs::read(MARKED_MASTER).unwrap();
    assert!(
        members(&original)["ppt/slideMasters/slideMaster1.xml"].starts_with(BOM),
        "the fixture's master carries a producer mark"
    );
    let (kept, marked) = differential(
        &original,
        |name| name == "ppt/slideMasters/slideMaster1.xml",
        |package| slide_texts(package),
    );
    assert_eq!(kept, marked);
}

#[test]
fn layout_placeholder_and_font_edits_of_marked_parts_match_unmarked_edits() {
    use litchi_pptx::master_layout::{PlaceholderKind, PlaceholderSpec, SlideLayoutKind};
    let deck = text_deck();
    differential(&deck, is_xml_member, |package| {
        let master = litchi_opc::PackURI::new("/ppt/slideMasters/slideMaster1.xml").unwrap();
        let first = package
            .add_slide_layout(&master, SlideLayoutKind::Blank, "First added", &[])
            .unwrap();
        let second = package
            .add_slide_layout(
                &master,
                SlideLayoutKind::Title,
                "Second added",
                &[PlaceholderSpec::new(PlaceholderKind::CenteredTitle).with_text("title")],
            )
            .unwrap();
        package
            .store_placeholder_shape(
                &second.part_name,
                &PlaceholderSpec::new(PlaceholderKind::CenteredTitle).with_text("retitled"),
            )
            .unwrap();
        package.remove_slide_layout(&first.part_name).unwrap();
        package.validate_master_layout_graph().unwrap();
        let mut fonts = litchi_pptx::font::Fonts::new();
        fonts
            .add(litchi_pptx::font::Font::new("Example Sans").unwrap())
            .unwrap();
        let put = package.put_fonts(fonts).map_err(|error| error.to_string());
        let removed = package
            .remove_fonts()
            .map(|fonts| fonts.map(|fonts| fonts.len()))
            .map_err(|error| error.to_string());
        format!("{put:?}|{removed:?}|{}", slide_texts(package))
    });
}

#[test]
fn theme_scheme_replacement_of_a_marked_theme_matches_the_unmarked_one() {
    // Base: an indented marked theme lost three indentation bytes and kept
    // the old scheme's tail (`me>`) as text inside `a:themeElements`, which
    // neither the read-back nor the publication audit refuses.
    use litchi_pptx::shape::theme::{self, Color, Slot};
    let deck = text_deck();
    let (kept, _) = differential(&deck, is_xml_member, |package| {
        package
            .edit_opc(|opc| {
                let name = "/ppt/theme/theme1.xml";
                let current = theme::load(opc, name)?;
                let colors = current
                    .colors
                    .clone()
                    .with(Slot::Accent1, Color::rgb("123456").unwrap());
                theme::put_colors(opc, name, &colors)?;
                let fonts = current.fonts.clone();
                theme::put_fonts(opc, name, &fonts)?;
                Ok(format!("{:?}", theme::load(opc, name)?))
            })
            .unwrap()
    });
    assert!(kept.contains("ppt/theme/theme1.xml"), "{kept:?}");
}
