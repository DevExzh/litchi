#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Plumbing for the leniency carrier: every open path produces a handle whose
//! tolerance report exists, and a conformant fixture leaves it empty.

use litchi_pptx::{Package, SourceBackedPresentation};

const SAMPLE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/ooxml/pptx/sample.pptx"
);
const SHAPES: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/shapes.pptx");

fn assert_clean(report: &litchi_pptx::ToleranceReport) {
    assert!(report.is_clean());
    assert!(report.defects().is_empty());
    assert_eq!(report.unrecorded(), 0);
    assert_eq!(report.total(), 0);
}

#[test]
fn package_open_of_a_conformant_fixture_reports_no_repairs() {
    let package = Package::open(SAMPLE).unwrap();
    assert_clean(package.tolerance_report());
}

#[test]
fn every_package_open_path_carries_an_empty_report_for_a_conformant_fixture() {
    assert_clean(Package::from_bytes(SHAPES).unwrap().tolerance_report());
    assert_clean(
        Package::from_vec(SHAPES.to_vec())
            .unwrap()
            .tolerance_report(),
    );
    assert_clean(Package::from_reader(SHAPES).unwrap().tolerance_report());
    assert_clean(Package::new().unwrap().tolerance_report());
}

#[test]
fn source_backed_open_of_a_conformant_fixture_reports_no_repairs() {
    let presentation = SourceBackedPresentation::open(SAMPLE).unwrap();
    assert_clean(presentation.tolerance_report());
}

#[test]
fn lazy_slide_access_leaves_the_open_report_clean() {
    let package = Package::from_bytes(SHAPES).unwrap();
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    assert!(!slides.is_empty());
    for slide in &slides {
        assert!(!slide.shapes().unwrap().is_empty());
    }
    assert_clean(package.tolerance_report());
}

// ---------------------------------------------------------------------------
// Audit finding 6.1: stale timing targets and duplicate shape IDs.
// ---------------------------------------------------------------------------

use litchi_opc::PackURI;
use litchi_pptx::animations::{Sequence, parse_package_slide_with_tolerance};
use litchi_pptx::{Defect, TimingDefect};

const SLIDE1: &str = "/ppt/slides/slide1.xml";

/// One click-triggered entrance effect on `spid`, nested the way
/// `PowerPoint` writes main-sequence click groups.
fn click_effect(id: &mut u32, spid: &str, group: &str) -> String {
    let [outer, inner, effect, set] = [0, 1, 2, 3].map(|offset| *id + offset);
    *id += 4;
    format!(
        concat!(
            r#"<p:par><p:cTn id="{outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>"#,
            r#"<p:par><p:cTn id="{inner}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>"#,
            r#"<p:par><p:cTn id="{effect}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="{group}" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>"#,
            r#"<p:set><p:cBhvr><p:cTn id="{set}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>"#,
            r"</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>",
        ),
        outer = outer,
        inner = inner,
        effect = effect,
        set = set,
        group = group,
        spid = spid,
    )
}

/// A `p:timing` element animating `targets` in click order, each with a
/// `p:bldP` build entry, the shape `PowerPoint` saves.
fn timing(targets: &[&str]) -> String {
    let mut id = 3;
    let effects: String = targets
        .iter()
        .map(|spid| click_effect(&mut id, spid, "0"))
        .collect();
    let builds: String = targets
        .iter()
        .map(|spid| format!(r#"<p:bldP spid="{spid}" grpId="0"/>"#))
        .collect();
    format!(
        concat!(
            r#"<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>"#,
            r#"<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>{effects}</p:childTnLst></p:cTn>"#,
            r#"<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>"#,
            r#"<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>"#,
            r"</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>{builds}</p:bldLst></p:timing>",
        ),
        effects = effects,
        builds = builds,
    )
}

/// The `shapes.pptx` first slide, with `timing` spliced in where
/// `PowerPoint` writes it and `edit` applied to the slide XML.
fn slide_xml(timing: &str, edit: impl FnOnce(String) -> String) -> Vec<u8> {
    let package = Package::from_bytes(SHAPES).unwrap();
    let part = PackURI::new(SLIDE1).unwrap();
    let original = package
        .opc()
        .unwrap()
        .get_part(&part)
        .unwrap()
        .blob()
        .to_vec();
    let original = String::from_utf8(original).unwrap();
    let xml = original.replace("</p:clrMapOvr>", &format!("</p:clrMapOvr>{timing}"));
    assert_ne!(xml, original, "fixture splice point is missing");
    edit(xml).into_bytes()
}

/// Write `slide` over the first slide and round-trip the package through
/// its serialized bytes, so the result is an ordinary file opened by default.
fn deck_with_slide(slide: Vec<u8>) -> Package {
    Package::from_vec(deck_bytes_with_part(SLIDE1, slide)).unwrap()
}

/// The serialized `shapes.pptx` deck with part `name` replaced by `blob`.
fn deck_bytes_with_part(name: &str, blob: Vec<u8>) -> Vec<u8> {
    let mut package = Package::from_bytes(SHAPES).unwrap();
    let part = PackURI::new(name).unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&part)?.set_blob(blob);
            Ok(())
        })
        .unwrap();
    package.to_bytes().unwrap()
}

/// The XML of part `name` in `shapes.pptx`.
fn original_part(name: &str) -> String {
    let package = Package::from_bytes(SHAPES).unwrap();
    let part = PackURI::new(name).unwrap();
    String::from_utf8(
        package
            .opc()
            .unwrap()
            .get_part(&part)
            .unwrap()
            .blob()
            .to_vec(),
    )
    .unwrap()
}

/// `original_part(name)` with `edit` applied, asserting the edit changed it.
fn patched_part(name: &str, edit: impl FnOnce(&str) -> String) -> Vec<u8> {
    let original = original_part(name);
    let patched = edit(&original);
    assert_ne!(
        patched, original,
        "fixture patch point is missing in {name}"
    );
    patched.into_bytes()
}

fn first_slide_animations(package: &Package) -> litchi_pptx::Result<Sequence> {
    let presentation = package.presentation()?;
    let slides = presentation.slides()?;
    slides[0].animations()
}

fn located(report: &litchi_pptx::ToleranceReport) -> Vec<(Defect, u32, Option<&str>)> {
    report
        .defects()
        .iter()
        .map(|tolerated| (tolerated.defect, tolerated.context, tolerated.part()))
        .collect()
}

#[test]
fn conformant_timing_tree_opens_unchanged_with_an_empty_report() {
    let xml = slide_xml(&timing(&["4", "7"]), |xml| xml);
    let package = deck_with_slide(xml.clone());
    let sequence = first_slide_animations(&package).unwrap();

    assert_clean(sequence.tolerance_report());
    assert_clean(package.tolerance_report());
    let targets: Vec<u32> = sequence.animations.iter().map(|a| a.shape_id).collect();
    assert_eq!(targets, [4, 7]);
    assert_eq!(sequence.paragraph_builds.len(), 2);
    // A conformant tree keeps its verbatim timing subtree for round-trip.
    assert!(sequence.preserved_timing_xml().is_some());
    // The strict verification form accepts the same slide.
    assert!(Sequence::parse_slide_xml_strict(&xml).is_ok());
}

#[test]
fn stale_shape_target_opens_by_default_and_is_recorded() {
    // Shape 99 was deleted after it was animated; PowerPoint leaves the
    // effect and its build entry behind and plays the rest.
    let xml = slide_xml(&timing(&["4", "99", "7"]), |xml| xml);
    let package = deck_with_slide(xml.clone());
    let sequence = first_slide_animations(&package).unwrap();

    let targets: Vec<(u32, u32)> = sequence
        .animations
        .iter()
        .map(|a| (a.shape_id, a.order))
        .collect();
    assert_eq!(targets, [(4, 1), (7, 2)]);
    let builds: Vec<u32> = sequence
        .paragraph_builds
        .iter()
        .map(|b| b.shape_id)
        .collect();
    assert_eq!(builds, [4, 7]);
    // The verbatim subtree still names shape 99, so it is not preserved.
    assert!(sequence.preserved_timing_xml().is_none());

    let orphan = Defect::Timing(TimingDefect::OrphanedTarget);
    assert_eq!(
        located(sequence.tolerance_report()),
        [(orphan, 99, Some(SLIDE1)), (orphan, 99, Some(SLIDE1))]
    );
    assert!(!orphan.repair().is_empty());
    // Lazy repairs do not leak into the open-time report.
    assert_clean(package.tolerance_report());

    // The repaired model regenerates timing XML that parses cleanly.
    let regenerated = sequence.to_xml();
    assert!(!regenerated.contains(r#"spid="99""#));
    let reparsed = Sequence::parse_timing_xml(&regenerated).unwrap();
    assert_eq!(reparsed.animations.len(), 2);

    // Strict verification still rejects the stale reference.
    let error = Sequence::parse_slide_xml_strict(&xml).unwrap_err();
    assert!(error.to_string().contains("99"), "{error}");
}

#[test]
fn stale_build_entry_alone_is_dropped_with_its_group() {
    // Only the build entry survived the deletion, in a group of its own.
    let xml = slide_xml(&timing(&["4"]), |xml| {
        xml.replace("</p:bldLst>", r#"<p:bldP spid="42" grpId="5"/></p:bldLst>"#)
    });
    let package = deck_with_slide(xml);
    let sequence = first_slide_animations(&package).unwrap();

    assert_eq!(sequence.animations.len(), 1);
    assert_eq!(sequence.paragraph_builds.len(), 1);
    assert_eq!(
        located(sequence.tolerance_report()),
        [(
            Defect::Timing(TimingDefect::OrphanedTarget),
            42,
            Some(SLIDE1)
        )]
    );
}

#[test]
fn duplicate_shape_id_opens_by_default_first_wins_and_is_recorded() {
    // The connector reuses the text box's ID, as copy/paste leaves it.
    let xml = slide_xml(&timing(&["4", "7"]), |xml| {
        xml.replace(r#"<p:cNvPr id="6" "#, r#"<p:cNvPr id="4" "#)
    });
    let package = deck_with_slide(xml.clone());
    let sequence = first_slide_animations(&package).unwrap();

    let targets: Vec<u32> = sequence.animations.iter().map(|a| a.shape_id).collect();
    assert_eq!(targets, [4, 7]);
    assert_eq!(
        located(sequence.tolerance_report()),
        [(
            Defect::Timing(TimingDefect::DuplicateShapeId),
            4,
            Some(SLIDE1)
        )]
    );
    assert!(Sequence::parse_slide_xml_strict(&xml).is_err());
}

#[test]
fn per_slide_reports_aggregate_into_one_deck_report() {
    let xml = slide_xml(&timing(&["4", "99"]), |xml| {
        xml.replace(r#"<p:cNvPr id="6" "#, r#"<p:cNvPr id="7" "#)
    });
    let package = deck_with_slide(xml);
    let part = PackURI::new(SLIDE1).unwrap();
    let mut deck = litchi_pptx::ToleranceReport::default();
    let sequence =
        parse_package_slide_with_tolerance(package.opc().unwrap(), &part, &mut deck).unwrap();

    assert_eq!(deck, *sequence.tolerance_report());
    assert_eq!(
        located(&deck),
        [
            (
                Defect::Timing(TimingDefect::DuplicateShapeId),
                7,
                Some(SLIDE1)
            ),
            (
                Defect::Timing(TimingDefect::OrphanedTarget),
                99,
                Some(SLIDE1)
            ),
            (
                Defect::Timing(TimingDefect::OrphanedTarget),
                99,
                Some(SLIDE1)
            ),
        ]
    );

    // A bare-XML parse has no part to name.
    let bare =
        Sequence::parse_slide_xml(package.opc().unwrap().get_part(&part).unwrap().blob()).unwrap();
    assert!(
        bare.tolerance_report()
            .defects()
            .iter()
            .all(|d| d.part().is_none())
    );
    assert_eq!(bare.tolerance_report().total(), 3);
}

#[test]
fn structural_timing_corruption_stays_fatal() {
    // Unbalanced timing XML is still a hard error, stale target or not.
    let broken = slide_xml(&timing(&["4", "99"]), |xml| xml.replace("</p:bldLst>", ""));
    assert!(Sequence::parse_slide_xml(&broken).is_err());
    // A malformed spid value is grammar, not a stale reference.
    let malformed = slide_xml(&timing(&["4", "nope"]), |xml| xml);
    assert!(Sequence::parse_slide_xml(&malformed).is_err());
    // A duplicate build shape/group pair is grammar too.
    let duplicate_pair = slide_xml(&timing(&["4", "4"]), |xml| xml);
    assert!(Sequence::parse_slide_xml(&duplicate_pair).is_err());
}

// ---------------------------------------------------------------------------
// Audit finding 6.2: one malformed transform degrades only its own shape.
// ---------------------------------------------------------------------------

use litchi_pptx::ShapeDefect;

fn first_slide_geometry(
    package: &Package,
) -> (Vec<(Option<u32>, bool)>, litchi_pptx::ToleranceReport) {
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let scene = slides[0].shapes().unwrap();
    let geometry = scene
        .iter()
        .map(|shape| (shape.id(), shape.bounds().is_some()))
        .collect();
    (geometry, scene.tolerance_report().clone())
}

#[test]
fn conformant_transforms_report_no_shape_repairs() {
    let package = Package::from_bytes(SHAPES).unwrap();
    let (geometry, report) = first_slide_geometry(&package);
    assert!(geometry.iter().all(|(_, bounded)| *bounded));
    assert_clean(&report);
}

#[test]
fn malformed_transform_degrades_only_its_shape_and_is_recorded() {
    // A missing `y` on the text box and a negative width on the connector,
    // as third-party converters emit for flipped shapes.
    let slide = patched_part(SLIDE1, |xml| {
        xml.replace(
            r#"<a:off x="914400" y="914400"/>"#,
            r#"<a:off x="914400"/>"#,
        )
        .replace(
            r#"<a:ext cx="2133600" cy="0"/>"#,
            r#"<a:ext cx="-2133600" cy="0"/>"#,
        )
    });
    let package = deck_with_slide(slide);
    let (geometry, report) = first_slide_geometry(&package);

    assert_eq!(
        geometry.len(),
        Package::from_bytes(SHAPES)
            .map(|p| first_slide_geometry(&p).0.len())
            .unwrap()
    );
    let degraded: Vec<Option<u32>> = geometry
        .iter()
        .filter(|(_, bounded)| !bounded)
        .map(|(id, _)| *id)
        .collect();
    assert_eq!(degraded, [Some(4), Some(6)]);
    assert!(geometry.contains(&(Some(7), true)));

    let malformed = Defect::Shape(ShapeDefect::MalformedTransform);
    assert_eq!(
        located(&report),
        [(malformed, 4, Some(SLIDE1)), (malformed, 6, Some(SLIDE1))]
    );
    assert!(!malformed.repair().is_empty());
    assert_clean(package.tolerance_report());
}

#[test]
fn non_numeric_transform_coordinate_is_degraded_not_fatal() {
    let slide = patched_part(SLIDE1, |xml| {
        xml.replace(
            r#"<a:off x="595961" y="2674961"/>"#,
            r#"<a:off x="12pt" y="2674961"/>"#,
        )
    });
    let package = deck_with_slide(slide);
    let (_, report) = first_slide_geometry(&package);
    assert_eq!(
        located(&report),
        [(
            Defect::Shape(ShapeDefect::MalformedTransform),
            7,
            Some(SLIDE1)
        )]
    );
}

#[test]
fn ill_formed_shape_xml_stays_fatal() {
    let slide = patched_part(SLIDE1, |xml| {
        xml.replace(
            r#"<a:off x="914400" y="914400"/>"#,
            r#"<a:off x="914400" y="914400">"#,
        )
    });
    assert!(litchi_pptx::shape::Scene::read(&slide).is_err());
}
