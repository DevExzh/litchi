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

// ---------------------------------------------------------------------------
// Audit finding 6.4: placeholder `p:ext` owner hygiene.
// ---------------------------------------------------------------------------

const SLIDE2: &str = "/ppt/slides/slide2.xml";
const P232: &str = "http://schemas.microsoft.com/office/powerpoint/2023/02/main";
const P232_OWNER: &str = "urn:litchi:pptx:p232:phTypeExt";

fn second_slide_placeholders(
    package: &Package,
) -> (
    Vec<(
        Option<u32>,
        Option<litchi_pptx::shape::PlaceholderTypeExtension>,
    )>,
    litchi_pptx::ToleranceReport,
) {
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let scene = slides[1].shapes().unwrap();
    let placeholders = scene
        .placeholders()
        .map(|shape| (shape.id(), shape.placeholder().unwrap().type_extension()))
        .collect();
    (placeholders, scene.tolerance_report().clone())
}

#[test]
fn placeholder_extension_hygiene_is_tolerated_per_placeholder_and_recorded() {
    let slide = patched_part(SLIDE2, |xml| {
        xml.replace(
            r#"<p:ph type="ctrTitle"/>"#,
            &format!(
                r#"<p:ph type="ctrTitle"><p:extLst><p:ext uri="{P232_OWNER}"><x:phTypeExt xmlns:x="{P232}"><x:type><x:cameo/><x:unknown/></x:type></x:phTypeExt></p:ext></p:extLst></p:ph>"#
            ),
        )
        .replace(
            r#"<p:ph type="subTitle" idx="1"/>"#,
            r#"<p:ph type="subTitle" idx="1"><p:extLst><p:ext><x:a xmlns:x="urn:x"/><x:b xmlns:x="urn:x"/></p:ext></p:extLst></p:ph>"#,
        )
    });
    let package = Package::from_vec(deck_bytes_with_part(SLIDE2, slide)).unwrap();
    let (placeholders, report) = second_slide_placeholders(&package);

    assert_eq!(placeholders, [(Some(3), None), (Some(4), None)]);
    assert_eq!(
        located(&report),
        [
            (
                Defect::Shape(ShapeDefect::PlaceholderTypeExtension),
                3,
                Some(SLIDE2)
            ),
            (
                Defect::Shape(ShapeDefect::PlaceholderExtensionOwner),
                4,
                Some(SLIDE2)
            ),
        ]
    );
}

#[test]
fn conformant_placeholder_extensions_report_nothing() {
    let slide = patched_part(SLIDE2, |xml| {
        xml.replace(
            r#"<p:ph type="ctrTitle"/>"#,
            &format!(
                r#"<p:ph type="ctrTitle"><p:extLst><p:ext uri="{P232_OWNER}"><x:phTypeExt xmlns:x="{P232}"><x:type><x:cameo/></x:type></x:phTypeExt></p:ext><p:ext uri="urn:future"/></p:extLst></p:ph>"#
            ),
        )
    });
    let package = Package::from_vec(deck_bytes_with_part(SLIDE2, slide)).unwrap();
    let (placeholders, report) = second_slide_placeholders(&package);
    assert_eq!(
        placeholders,
        [
            (
                Some(3),
                Some(litchi_pptx::shape::PlaceholderTypeExtension::Cameo)
            ),
            (Some(4), None)
        ]
    );
    assert_clean(&report);
}

// ---------------------------------------------------------------------------
// Audit findings 6.5/6.6: transition attribute domains.
// ---------------------------------------------------------------------------

use litchi_pptx::TransitionDefect;
use litchi_pptx::transition::Kind as TransitionKind;

/// The first slide of `shapes.pptx` with `transition` spliced in where
/// `PowerPoint` writes it.
fn deck_with_transition(transition: &str) -> Vec<u8> {
    let slide = patched_part(SLIDE1, |xml| {
        xml.replace("</p:clrMapOvr>", &format!("</p:clrMapOvr>{transition}"))
    });
    deck_bytes_with_part(SLIDE1, slide)
}

/// The first slide's transition, read through the source-backed snapshot
/// that names its part.
fn first_slide_transition(deck: Vec<u8>) -> litchi_pptx::transition::Transition {
    let editor = litchi_pptx::SourceBackedPresentationEditor::from_read_at(std::sync::Arc::new(
        litchi_core::source::OwnedSource::new(deck),
    ))
    .unwrap();
    editor
        .slide_snapshot(0)
        .unwrap()
        .transition()
        .unwrap()
        .expect("the slide has a transition")
}

#[test]
fn unsupported_wheel_spoke_count_is_kept_raw_and_recorded() {
    let transition = first_slide_transition(deck_with_transition(
        r#"<p:transition spd="slow"><p:wheel spokes="6"/></p:transition>"#,
    ));
    let TransitionKind::Raw(raw) = transition.kind() else {
        panic!("expected a raw effect, got {:?}", transition.kind());
    };
    assert!(raw.xml().contains(r#"spokes="6""#));
    assert_eq!(transition.speed(), litchi_pptx::transition::Speed::Slow);
    assert_eq!(
        located(transition.tolerance_report()),
        [(
            Defect::Transition(TransitionDefect::SpokeCount),
            6,
            Some(SLIDE1)
        )]
    );
}

#[test]
fn conformant_transition_reports_nothing() {
    let transition = first_slide_transition(deck_with_transition(
        r#"<p:transition spd="fast" advTm="2000"><p:wheel spokes="8"/></p:transition>"#,
    ));
    assert!(matches!(transition.kind(), TransitionKind::Wheel(_)));
    assert_clean(transition.tolerance_report());
}

#[test]
fn out_of_domain_transition_tokens_and_times_are_repaired_and_recorded() {
    let transition = first_slide_transition(deck_with_transition(
        r#"<p:transition spd="warp" advTm="3000000000"><p:push dir="horz"/></p:transition>"#,
    ));
    assert_eq!(transition.speed(), litchi_pptx::transition::Speed::Medium);
    assert_eq!(
        transition.after().map(litchi_pptx::transition::Ms::get),
        Some(litchi_pptx::transition::MAX_MS)
    );
    let TransitionKind::Raw(raw) = transition.kind() else {
        panic!("expected a raw effect, got {:?}", transition.kind());
    };
    assert!(raw.xml().contains(r#"dir="horz""#));
    assert_eq!(
        located(transition.tolerance_report()),
        [
            (
                Defect::Transition(TransitionDefect::AttributeToken),
                0,
                Some(SLIDE1)
            ),
            (
                Defect::Transition(TransitionDefect::TimeClamped),
                3_000_000_000,
                Some(SLIDE1)
            ),
            (
                Defect::Transition(TransitionDefect::EffectToken),
                0,
                Some(SLIDE1)
            ),
        ]
    );
}

// ---------------------------------------------------------------------------
// Audit finding 6.7: presentation slide-list uniqueness and child hygiene.
// ---------------------------------------------------------------------------

use litchi_pptx::PresentationDefect;

const PRESENTATION: &str = "/ppt/presentation.xml";

fn deck_with_slide_list(edit: impl FnOnce(&str) -> String) -> Vec<u8> {
    deck_bytes_with_part(PRESENTATION, patched_part(PRESENTATION, edit))
}

#[test]
fn slide_list_defects_are_repaired_first_wins_and_recorded_at_open() {
    let deck = deck_with_slide_list(|xml| {
        xml.replace("<p:sldIdLst>", r#"<p:sldIdLst foo="1">"#)
            .replace(
                r#"<p:sldId id="256" r:id="rId2"/>"#,
                concat!(
                    r#"<p:sldId id="256" r:id="rId2"/>"#,
                    r#"<p:sldId id="256" r:id="rId6"/>"#,
                    r#"<p:sldId id="300" r:id="rId2"/>"#,
                    r#"<x:future xmlns:x="urn:x"><p:sldId id="301" r:id="rId9"/></x:future>"#,
                    "junk",
                ),
            )
    });

    let content = Defect::Presentation(PresentationDefect::SlideListContent);
    let expected = [
        (content, 0, Some(PRESENTATION)),
        (
            Defect::Presentation(PresentationDefect::DuplicateSlideId),
            256,
            Some(PRESENTATION),
        ),
        (
            Defect::Presentation(PresentationDefect::DuplicateSlideRelationship),
            300,
            Some(PRESENTATION),
        ),
        (content, 0, Some(PRESENTATION)),
        (content, 0, Some(PRESENTATION)),
    ];

    let package = Package::from_vec(deck.clone()).unwrap();
    assert_eq!(located(package.tolerance_report()), expected);
    let presentation = package.presentation().unwrap();
    let ids: Vec<u32> = presentation
        .slide_references()
        .unwrap()
        .iter()
        .map(litchi_pptx::parts::SlideReference::id)
        .collect();
    assert_eq!(ids, [256, 258, 257, 259, 260, 261]);
    assert_eq!(presentation.slides().unwrap().len(), 6);

    let source = SourceBackedPresentation::from_read_at(std::sync::Arc::new(
        litchi_core::source::OwnedSource::new(deck.clone()),
    ))
    .unwrap();
    assert_eq!(located(source.tolerance_report()), expected);

    // The lossless editor keeps the strict slide-list contract.
    assert!(
        litchi_pptx::SourceBackedPresentationEditor::from_read_at(std::sync::Arc::new(
            litchi_core::source::OwnedSource::new(deck),
        ))
        .is_err()
    );
}

#[test]
fn conformant_slide_list_records_nothing_and_invalid_entries_stay_fatal() {
    let package = Package::from_bytes(SHAPES).unwrap();
    assert_clean(package.tolerance_report());

    // An entry whose own value is outside ST_SlideId is grammar, not a
    // uniqueness repair: the slide list still fails, lazily, as before.
    let invalid = deck_with_slide_list(|xml| {
        xml.replace(
            r#"<p:sldId id="256" r:id="rId2"/>"#,
            r#"<p:sldId id="255" r:id="rId2"/>"#,
        )
    });
    let package = Package::from_vec(invalid).unwrap();
    assert_clean(package.tolerance_report());
    assert!(
        package
            .presentation()
            .and_then(|presentation| presentation.slide_references())
            .is_err()
    );
}

// ---------------------------------------------------------------------------
// Audit finding 6.3: processing instructions are skipped, recorded once.
// ---------------------------------------------------------------------------

use litchi_pptx::MarkupDefect;

#[test]
fn processing_instructions_are_skipped_and_recorded_once_per_part() {
    let slide = patched_part(SLIDE1, |xml| {
        xml.replace("<p:spTree>", "<p:spTree><?publisher stage=\"2\"?>")
            .replace(
                "<a:t>Learning PPTX</a:t>",
                "<a:t>Learning <?mark?>PPTX</a:t>",
            )
            .replace(
                "</p:clrMapOvr>",
                r#"</p:clrMapOvr><p:transition><p:fade><?fx?></p:fade></p:transition>"#,
            )
    });
    let package = deck_with_slide(slide);
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();

    let scene = slides[0].shapes().unwrap();
    assert_eq!(
        located(scene.tolerance_report()),
        [(
            Defect::Markup(MarkupDefect::ProcessingInstruction),
            0,
            Some(SLIDE1)
        )]
    );
    assert!(slides[0].text().unwrap().contains("Learning PPTX"));
    assert!(slides[0].animations().is_ok());

    let blob = package
        .opc()
        .unwrap()
        .get_part(&PackURI::new(SLIDE1).unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let transition = litchi_pptx::transition::read(&blob).unwrap().unwrap();
    assert_eq!(transition.kind(), &TransitionKind::Fade { black: None });

    // DOCTYPE stays rejected.
    let doctype =
        String::from_utf8(blob)
            .unwrap()
            .replacen("<p:sld ", "<!DOCTYPE p:sld><p:sld ", 1);
    assert!(litchi_pptx::shape::Scene::read(doctype.as_bytes()).is_err());
}

// ---------------------------------------------------------------------------
// Audit finding 6.8: unparseable shape ID / placeholder index.
// ---------------------------------------------------------------------------

#[test]
fn unparseable_shape_id_and_placeholder_index_degrade_and_are_recorded() {
    let slide = patched_part(SLIDE2, |xml| {
        xml.replace(r#"<p:cNvPr id="3" "#, r#"<p:cNvPr id="" "#)
            .replace(
                r#"<p:ph type="subTitle" idx="1"/>"#,
                r#"<p:ph type="subTitle" idx="x1"/>"#,
            )
    });
    let package = Package::from_vec(deck_bytes_with_part(SLIDE2, slide)).unwrap();
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let scene = slides[1].shapes().unwrap();
    let placeholders: Vec<_> = scene
        .placeholders()
        .map(|shape| (shape.id(), shape.placeholder().unwrap().index()))
        .collect();
    assert_eq!(placeholders, [(None, 0), (Some(4), 0)]);
    assert_eq!(
        located(scene.tolerance_report()),
        [
            (Defect::Shape(ShapeDefect::InvalidShapeId), 0, Some(SLIDE2)),
            (
                Defect::Shape(ShapeDefect::InvalidPlaceholderIndex),
                4,
                Some(SLIDE2)
            ),
        ]
    );
}

// ---------------------------------------------------------------------------
// Packaging and metadata reports folded into the open-time report.
// ---------------------------------------------------------------------------

/// `shapes.pptx` with archive member `name` rewritten by `edit`, every other
/// member copied byte for byte in its original order.
fn shapes_with_member(name: &str, edit: impl FnOnce(&str) -> String) -> Vec<u8> {
    archive_with_member(SHAPES, name, edit)
}

/// `archive` with member `name` rewritten by `edit`, every other member
/// copied byte for byte in its original order.
fn archive_with_member(archive: &[u8], name: &str, edit: impl FnOnce(&str) -> String) -> Vec<u8> {
    use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
    let reader = ArchiveReader::new(archive).unwrap();
    let names: Vec<String> = reader.file_names().map(str::to_owned).collect();
    let mut edit = Some(edit);
    let mut writer = StreamingArchiveWriter::new();
    for member in names {
        let mut bytes = reader.read(&member).unwrap();
        if member == name {
            let original = String::from_utf8(bytes).unwrap();
            let patched = (edit.take().unwrap())(&original);
            assert_ne!(
                patched, original,
                "fixture patch point is missing in {name}"
            );
            bytes = patched.into_bytes();
        }
        writer.write_stored(&member, &bytes).unwrap();
    }
    assert!(edit.is_none(), "archive member {name} is missing");
    writer.finish_to_bytes().unwrap()
}

#[test]
fn packaging_repairs_are_folded_into_the_open_report() {
    let deck = shapes_with_member("ppt/_rels/presentation.xml.rels", |xml| {
        xml.replace(
            "</Relationships>",
            r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide6.xml"/></Relationships>"#,
        )
    });
    let duplicate = Defect::Package(litchi_opc::leniency::Defect::Relationship(
        litchi_opc::leniency::RelationshipDefect::DuplicateId,
    ));

    let package = Package::from_vec(deck.clone()).unwrap();
    let report = package.tolerance_report();
    assert_eq!(report.defects().len(), 1, "{report:?}");
    let tolerated = &report.defects()[0];
    assert_eq!(tolerated.defect, duplicate);
    assert_eq!(tolerated.part(), Some("/ppt/presentation.xml"));
    assert_eq!(tolerated.item(), Some("rId2"));
    assert_eq!(tolerated.value(), Some("slides/slide6.xml"));
    assert!(!tolerated.defect.repair().is_empty());
    // The package still opens and lists its slides through the kept rId2.
    assert_eq!(package.presentation().unwrap().slides().unwrap().len(), 6);

    let source = SourceBackedPresentation::from_read_at(std::sync::Arc::new(
        litchi_core::source::OwnedSource::new(deck),
    ))
    .unwrap();
    assert_eq!(source.tolerance_report().defects(), report.defects());
}

#[test]
fn core_properties_repairs_are_folded_into_the_open_report() {
    let deck = shapes_with_member("docProps/core.xml", |xml| {
        xml.replace(
            "<dc:creator>",
            "<dc:title>Second title</dc:title><dc:creator>",
        )
    });
    let duplicate = Defect::Properties(litchi_ooxml_common::leniency::Defect::CoreProperties(
        litchi_ooxml_common::leniency::CorePropertiesDefect::DuplicateProperty,
    ));

    let package = Package::from_vec(deck.clone()).unwrap();
    let defects = package.tolerance_report().defects();
    assert_eq!(defects.len(), 1, "{defects:?}");
    assert_eq!(defects[0].defect, duplicate);
    assert_eq!(defects[0].part(), Some("/docProps/core.xml"));

    // Source-backed open reads only the presentation root; the metadata
    // read is where its report is collected.
    let source = SourceBackedPresentation::from_read_at(std::sync::Arc::new(
        litchi_core::source::OwnedSource::new(deck),
    ))
    .unwrap();
    assert_clean(source.tolerance_report());
    let mut report = litchi_pptx::ToleranceReport::default();
    let properties = source
        .properties_with_tolerance(&mut report)
        .unwrap()
        .unwrap();
    assert_eq!(properties.title.as_deref(), Some("PowerPoint Presentation"));
    assert_eq!(report.defects().len(), 1);
    assert_eq!(report.defects()[0].defect, duplicate);

    // The conformant fixture folds nothing.
    let source = SourceBackedPresentation::from_read_at(std::sync::Arc::new(
        litchi_core::source::OwnedSource::new(SHAPES.to_vec()),
    ))
    .unwrap();
    let mut report = litchi_pptx::ToleranceReport::default();
    source.properties_with_tolerance(&mut report).unwrap();
    assert_clean(&report);
}

#[test]
fn ill_formed_package_manifests_stay_fatal() {
    let deck = shapes_with_member("ppt/_rels/presentation.xml.rels", |xml| {
        xml.replacen("<Relationship ", "<Relationship<", 1)
    });
    assert!(Package::from_vec(deck).is_err());
}

// ---------------------------------------------------------------------------
// Audit finding 6.12: legacy comment index 0 and duplicate (authorId, idx).
// ---------------------------------------------------------------------------

use litchi_pptx::CommentDefect;

const COMMENTED: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/slideshow/45545_Comment.pptx");
const COMMENT1: &str = "/ppt/comments/comment1.xml";
const COMMENT2: &str = "/ppt/comments/comment2.xml";

/// `45545_Comment.pptx` with a second comment of index 0 in `comment1.xml`
/// (slide 1) and the `comment2.xml` comment (a later slide) renumbered to
/// repeat `comment1.xml`'s `(0, 1)` key.
fn deck_with_comment_defects() -> Vec<u8> {
    let deck = archive_with_member(COMMENTED, "ppt/comments/comment1.xml", |xml| {
        xml.replace(
            "</p:cmLst>",
            r#"<p:cm authorId="0" idx="0"><p:pos x="1" y="2"/><p:text>zero</p:text></p:cm></p:cmLst>"#,
        )
    });
    archive_with_member(&deck, "ppt/comments/comment2.xml", |xml| {
        xml.replace(r#"idx="2""#, r#"idx="1""#)
    })
}

fn located_items(
    report: &litchi_pptx::ToleranceReport,
) -> Vec<(Defect, u32, Option<&str>, Option<&str>)> {
    report
        .defects()
        .iter()
        .map(|tolerated| {
            (
                tolerated.defect,
                tolerated.context,
                tolerated.item(),
                tolerated.part(),
            )
        })
        .collect()
}

#[test]
fn comment_index_zero_and_duplicate_keys_are_kept_and_recorded() {
    use litchi_pptx::comments::{
        load_presentation_comments, load_presentation_comments_strict,
        load_presentation_comments_with_tolerance,
    };
    let package = Package::from_vec(deck_with_comment_defects()).unwrap();
    let opc = package.opc().unwrap();

    let mut report = litchi_pptx::ToleranceReport::default();
    let comments = load_presentation_comments_with_tolerance(opc, &mut report)
        .unwrap()
        .unwrap();
    let keys: Vec<Vec<(u32, u32, &str)>> = comments
        .slides
        .iter()
        .map(|slide| {
            slide
                .comments
                .iter()
                .map(|comment| (comment.author_id, comment.index, comment.text.as_str()))
                .collect()
        })
        .collect();
    assert_eq!(
        keys,
        [
            vec![(0, 1, "testdoc"), (0, 0, "zero")],
            vec![(0, 1, "test phrase")],
        ]
    );
    assert_eq!(
        located_items(&report),
        [
            (
                Defect::Comment(CommentDefect::ZeroIndex),
                0,
                Some("0"),
                Some(COMMENT1)
            ),
            (
                Defect::Comment(CommentDefect::DuplicateKey),
                0,
                Some("1"),
                Some(COMMENT2)
            ),
        ]
    );
    assert_eq!(load_presentation_comments(opc).unwrap(), Some(comments));

    // The edit functions load strictly and refuse to rewrite the graph.
    assert!(load_presentation_comments_strict(opc).is_err());
    let mut editable = opc.clone();
    let before = editable
        .get_part(&PackURI::new(COMMENT1).unwrap())
        .unwrap()
        .blob()
        .to_vec();
    assert!(
        litchi_pptx::comments::remove_presentation_comment(
            &mut editable,
            "/ppt/slides/slide1.xml",
            0,
            1,
            litchi_pptx::Conformance::Transitional,
        )
        .is_err()
    );
    assert_eq!(
        editable
            .get_part(&PackURI::new(COMMENT1).unwrap())
            .unwrap()
            .blob(),
        before.as_slice()
    );

    // A key repeated within one list is kept after the first, too.
    let xml = concat!(
        r#"<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
        r#"<p:cm authorId="3" idx="4"><p:pos x="0" y="0"/><p:text>a</p:text></p:cm>"#,
        r#"<p:cm authorId="3" idx="4"><p:pos x="0" y="0"/><p:text>b</p:text></p:cm></p:cmLst>"#,
    );
    let mut report = litchi_pptx::ToleranceReport::default();
    let list =
        litchi_pptx::comments::parse_slide_comments_with_tolerance(xml.as_bytes(), &mut report)
            .unwrap();
    assert_eq!(list.len(), 2);
    assert_eq!(
        located_items(&report),
        [(
            Defect::Comment(CommentDefect::DuplicateKey),
            3,
            Some("4"),
            None
        )]
    );
    assert!(litchi_pptx::comments::parse_slide_comments_strict(xml.as_bytes()).is_err());
}

#[test]
fn conformant_comments_record_nothing_and_malformed_comments_stay_fatal() {
    use litchi_pptx::comments::load_presentation_comments_with_tolerance;
    let package = Package::from_bytes(COMMENTED).unwrap();
    let mut report = litchi_pptx::ToleranceReport::default();
    let comments = load_presentation_comments_with_tolerance(package.opc().unwrap(), &mut report)
        .unwrap()
        .unwrap();
    assert_eq!(comments.slides.len(), 2);
    assert_clean(&report);
    assert!(
        litchi_pptx::comments::load_presentation_comments_strict(package.opc().unwrap())
            .unwrap()
            .is_some()
    );

    // A non-numeric index is not an index-0/duplicate repair.
    let deck = archive_with_member(COMMENTED, "ppt/comments/comment1.xml", |xml| {
        xml.replace(r#"idx="1""#, r#"idx="one""#)
    });
    let package = Package::from_vec(deck).unwrap();
    assert!(litchi_pptx::comments::load_presentation_comments(package.opc().unwrap()).is_err());
}

// ---------------------------------------------------------------------------
// Audit finding 6.9: foreign `t` elements and character data outside `a:t`.
// ---------------------------------------------------------------------------

#[test]
fn foreign_text_elements_are_skipped_and_recorded_by_text_extraction() {
    let original_text = {
        let package = Package::from_bytes(SHAPES).unwrap();
        let presentation = package.presentation().unwrap();
        let slides = presentation.slides().unwrap();
        let mut report = litchi_pptx::ToleranceReport::default();
        let text = slides[0].text_with_tolerance(&mut report).unwrap();
        assert_clean(&report);
        text
    };
    let slide = patched_part(SLIDE1, |xml| {
        xml.replace(
            "</p:sld>",
            concat!(
                r#"<p:extLst><p:ext uri="{00000000-0000-0000-0000-000000000001}">"#,
                r#"<x:note xmlns:x="urn:example:x"><x:t>hidden <x:t>deeper</x:t></x:t>"#,
                r#"a &amp; b<![CDATA[<raw>]]><x:t/></x:note>"#,
                r#"</p:ext></p:extLst></p:sld>"#,
            ),
        )
    });
    let deck = deck_bytes_with_part(SLIDE1, slide);
    let package = Package::from_vec(deck.clone()).unwrap();
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();

    let mut report = litchi_pptx::ToleranceReport::default();
    let text = slides[0].text_with_tolerance(&mut report).unwrap();
    assert_eq!(text, original_text);
    assert_eq!(slides[0].text().unwrap(), original_text);
    let foreign = (
        Defect::Markup(MarkupDefect::ForeignTextElement),
        0,
        Some(SLIDE1),
    );
    assert_eq!(located(&report), [foreign, foreign]);

    let source = SourceBackedPresentation::from_read_at(std::sync::Arc::new(
        litchi_core::source::OwnedSource::new(deck),
    ))
    .unwrap();
    let slide = source.slides().next().unwrap();
    let mut report = litchi_pptx::ToleranceReport::default();
    assert_eq!(
        slide.text_with_tolerance(&mut report).unwrap(),
        original_text
    );
    assert_eq!(located(&report), [foreign, foreign]);

    // Nested DrawingML text stays a grammar error.
    let nested = patched_part(SLIDE1, |xml| {
        xml.replace(
            "<a:t>Learning PPTX</a:t>",
            "<a:t>Learning <a:t>PPTX</a:t></a:t>",
        )
    });
    let package = Package::from_vec(deck_bytes_with_part(SLIDE1, nested)).unwrap();
    let presentation = package.presentation().unwrap();
    assert!(presentation.slides().unwrap()[0].text().is_err());
}
