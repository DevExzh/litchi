//! Semantic regressions using unmodified native Microsoft PowerPoint files.

use litchi_pptx::Package;

fn native(name: &str) -> Package {
    Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(name),
    )
    .unwrap()
}

#[test]
fn native_equations_keep_shape_owners_formula_structure_and_fallback_parts() {
    let package = native("powerpoint-fonts-math-handout.pptx");
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    assert_eq!(slides.len(), 4);
    assert!(slides[0].equations().unwrap().is_empty());
    // Expected formulae follow the stored OMML; the inspection prose uses
    // typographic superscripts, whereas m:t contains ordinary digits.
    let expected = [
        (r"𝑥=\frac{−𝑏±\sqrt{𝑏^{2}−4𝑎𝑐}}{2𝑎}", "𝑥=−𝑏±𝑏2−4𝑎𝑐2𝑎"),
        (r"\sum_{𝑘=1}^{4} 𝑘^{2}=30", "𝑘=14𝑘2=30"),
        // The current formula projection documents text-only fallback for
        // matrices; this assertion does not claim row/column semantics.
        ("𝐴=1234", "𝐴=1234"),
    ];
    let reopened = Package::from_bytes(&package.to_bytes().unwrap()).unwrap();
    let reopened_presentation = reopened.presentation().unwrap();
    let reopened_slides = reopened_presentation.slides().unwrap();
    for (index, (formula, text)) in expected.into_iter().enumerate() {
        let slide = &slides[index + 1];
        let equations = slide.equations().unwrap();
        assert_eq!(equations.len(), 1);
        assert_eq!(equations[0].shape_id(), Some(3));
        assert_eq!(equations[0].formula(), formula);
        assert_eq!(equations[0].text(), text);
        assert_eq!(reopened_slides[index + 1].equations().unwrap(), equations);
        assert_eq!(
            reopened_slides[index + 1].part().part().blob(),
            slide.part().part().blob()
        );
        let relationship = slide.part().part().rels().get("rId2").unwrap();
        assert!(!relationship.is_external());
        assert_eq!(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::IMAGE
        );
        let target = relationship.target_partname().unwrap();
        assert_eq!(
            target.as_str(),
            format!("/ppt/media/image{}.png", index + 1)
        );
        let original_image = package.opc().unwrap().get_part(&target).unwrap();
        let reopened_image = reopened.opc().unwrap().get_part(&target).unwrap();
        assert_eq!(original_image.content_type(), "image/png");
        assert!(original_image.blob().starts_with(b"\x89PNG\r\n\x1a\n"));
        assert_eq!(original_image.blob(), reopened_image.blob());
    }
    assert!(package.math_properties().unwrap().is_none());
}

#[test]
fn native_transition_projection_keeps_push_morph_and_saved_timing() {
    use litchi_pptx::transition::{self, Kind, Morph, Ms, Side, Speed};
    let package = native("powerpoint-master-motion.pptx");
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let transitions: Vec<_> = slides
        .iter()
        .map(|slide| {
            transition::read(slide.part().part().blob(), transition::Limits::DEFAULT).unwrap()
        })
        .collect();
    assert_eq!(transitions.len(), 7);
    let push = transitions[1].as_ref().unwrap();
    assert_eq!(push.kind(), &Kind::Push(Side::Up));
    assert_eq!(push.speed(), Speed::Medium);
    assert_eq!(push.after().map(Ms::get), Some(3000));
    assert!(push.click());
    // The native UI reported 0.75 seconds, but this part saved only spd=med.
    assert_eq!(push.duration(), None);
    let morph = transitions[2].as_ref().unwrap();
    assert_eq!(morph.kind(), &Kind::Morph(Morph::ByObject));
    assert_eq!(morph.speed(), Speed::Slow);
    assert_eq!(morph.duration().map(Ms::get), Some(1500));
    assert_eq!(morph.after(), None);
    assert!(morph.click());
    for index in [0, 3, 4, 5, 6] {
        assert!(transitions[index].is_none());
    }
    for transition in [push, morph] {
        let fragment = transition::write(transition, transition::Limits::DEFAULT).unwrap();
        let xml = format!(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">{fragment}</p:sld>"#
        );
        let reopened = transition::read(xml.as_bytes(), transition::Limits::DEFAULT)
            .unwrap()
            .unwrap();
        assert!(reopened.same_semantics(transition));
    }
}

#[test]
fn native_master_theme_keeps_ion_palette_and_font_owners() {
    let package = native("powerpoint-master-motion.pptx");
    let presentation = package.presentation().unwrap();
    let masters = presentation.slide_masters().unwrap();
    assert_eq!(masters.len(), 1);
    let master = &masters[0];
    let theme = master.theme().unwrap().unwrap();
    assert_eq!(theme.name, "Ion");
    assert_eq!(
        theme.major_font.as_ref().unwrap().typeface,
        "Century Gothic"
    );
    assert_eq!(
        theme.minor_font.as_ref().unwrap().typeface,
        "Century Gothic"
    );
    assert_eq!(
        theme
            .colors
            .iter()
            .map(|color| (
                color.name.as_str(),
                color.rgb.as_deref(),
                color.system_color.as_deref()
            ))
            .collect::<Vec<_>>(),
        [
            ("dk1", Some("000000"), Some("windowText")),
            ("lt1", Some("FFFFFF"), Some("window")),
            ("dk2", Some("1E5155"), None),
            ("lt2", Some("EBEBEB"), None),
            ("accent1", Some("B01513"), None),
            ("accent2", Some("EA6312"), None),
            ("accent3", Some("E6B729"), None),
            ("accent4", Some("6AAC90"), None),
            ("accent5", Some("54849A"), None),
            ("accent6", Some("9E5E9B"), None),
            ("hlink", Some("58C1BA"), None),
            ("folHlink", Some("9DFFCB"), None),
        ]
    );
    let relationship = master
        .part()
        .part()
        .rels()
        .iter()
        .find(|r| r.reltype() == litchi_opc::constants::relationship_type::THEME)
        .unwrap();
    assert!(!relationship.is_external());
    let target = relationship.target_partname().unwrap();
    assert_eq!(target.as_str(), "/ppt/theme/theme1.xml");
    let part = package.opc().unwrap().get_part(&target).unwrap();
    assert_eq!(part.blob().len(), 7010);
    assert_eq!(
        sha256(part.blob()),
        "7be9ced965f7ce8f26f6d9776e719e8df09e4bdf0a874f2a256da54e778a3750"
    );
}

#[test]
fn native_handout_placeholders_expose_saved_text_visibility_and_background() {
    use litchi_pptx::presentation_properties::metadata::handout::{Limits, Master};
    let package = native("powerpoint-fonts-math-handout.pptx");
    let opc = package.opc().unwrap();
    let part = opc
        .get_part(&litchi_opc::PackURI::new("/ppt/handoutMasters/handoutMaster1.xml").unwrap())
        .unwrap();
    let parsed =
        Master::parse_xml(std::str::from_utf8(part.blob()).unwrap(), Limits::DEFAULT).unwrap();
    assert_eq!(
        parsed.header_footer.header_text.as_deref(),
        Some("HANDOUT_HEADER_009")
    );
    assert_eq!(
        parsed.header_footer.footer_text.as_deref(),
        Some("HANDOUT_FOOTER_009")
    );
    assert!(parsed.header_footer.show_header);
    assert!(parsed.header_footer.show_footer);
    assert!(parsed.header_footer.show_slide_number);
    assert!(!parsed.header_footer.show_date_time);
    assert!(!parsed.header_footer.auto_date);
    assert_eq!(parsed.header_footer.date_time_text, None);
    assert_eq!(parsed.background_color.as_deref(), Some("00B0F0"));
    // The print dialog was canceled; no six-slide print preference was saved.
    assert!(
        litchi_pptx::presentation_properties::load_from_package(
            opc,
            litchi_pptx::presentation_properties::Limits::default()
        )
        .unwrap()
        .unwrap()
        .print
        .is_none()
    );
    let xml = parsed.to_xml(Limits::DEFAULT).unwrap();
    let reopened = Master::parse_xml(&xml, Limits::DEFAULT).unwrap();
    assert_eq!(
        reopened.header_footer.header_text,
        parsed.header_footer.header_text
    );
    assert_eq!(
        reopened.header_footer.footer_text,
        parsed.header_footer.footer_text
    );
    assert_eq!(reopened.background_color, parsed.background_color);
    assert!(!reopened.header_footer.show_date_time);
}

#[test]
fn native_powerpoint_font_payload_and_face_metadata_remain_inert() {
    use litchi_pptx::font::{self, Style};
    let package = native("powerpoint-fonts-math-handout.pptx");
    let fonts = font::load(package.opc().unwrap()).unwrap().unwrap();
    assert_eq!(fonts.len(), 1);
    let font = fonts.get("Cambria Math").unwrap();
    assert_eq!(font.name(), "Cambria Math");
    assert_eq!(font.faces().len(), 1);
    let face = &font.faces()[0];
    assert_eq!(face.style(), Style::Regular);
    assert_eq!(face.data().bytes().len(), 129261);
    assert_eq!(
        sha256(face.data().bytes()),
        "4e9c7d3b3dc757b3f45eb6f7a902046ebabc9b5868c707aeaa4f1e6c42cccfc9"
    );
    let original = package.to_bytes().unwrap();
    let reopened = Package::from_bytes(&original).unwrap();
    assert_eq!(
        font::load(reopened.opc().unwrap())
            .unwrap()
            .unwrap()
            .get("Cambria Math")
            .unwrap()
            .faces()[0]
            .data()
            .bytes(),
        face.data().bytes()
    );
}

#[test]
fn native_powerpoint_slide_master_and_layout_root_properties_are_readable() {
    let package = native("powerpoint-master-motion.pptx");
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    assert_eq!(slides.len(), 7);
    for (index, slide) in slides.iter().enumerate() {
        assert_eq!(slide.is_hidden().unwrap(), index == 3);
    }
    let masters = presentation.slide_masters().unwrap();
    assert_eq!(masters.len(), 1);
    assert!(!masters[0].is_preserved().unwrap());
    let layouts = presentation.slide_layouts().unwrap();
    assert_eq!(layouts.len(), 17);
    let expected = [
        Some("title"),
        Some("obj"),
        Some("secHead"),
        Some("twoObj"),
        Some("twoTxTwoObj"),
        Some("titleOnly"),
        Some("blank"),
        Some("objTx"),
        Some("picTx"),
        None,
        None,
        None,
        None,
        None,
        None,
        Some("vertTx"),
        Some("vertTitleAndTx"),
    ];
    for (layout, kind) in layouts.iter().zip(expected) {
        assert_eq!(layout.kind().unwrap().as_deref(), kind);
    }
    assert_eq!(layouts[0].name().unwrap(), "Fixture Title Layout 006");
}

#[test]
fn native_powerpoint_fly_in_preserves_its_shared_target_and_duration() {
    use litchi_pptx::animations::{Direction, Duration, Effect, Trigger};
    let package = native("powerpoint-master-motion.pptx");
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let sequence = slides[1].animations().unwrap();
    assert_eq!(sequence.animations.len(), 1);
    let animation = &sequence.animations[0];
    assert_eq!(animation.shape_id, 2);
    assert_eq!(animation.effect, Effect::FlyIn);
    assert_eq!(animation.direction, Some(Direction::Left));
    assert_eq!(animation.trigger, Trigger::AfterPrevious);
    assert_eq!(animation.duration, Duration::Finite(1250));
    assert_eq!(sequence.paragraph_builds.len(), 1);
    assert!(sequence.timing_tree().is_some());
    assert!(sequence.preserved_timing_xml().is_some());
    let xml = sequence
        .to_xml(litchi_pptx::animations::Limits::DEFAULT)
        .unwrap();
    assert_eq!(Some(xml.as_str()), sequence.preserved_timing_xml());
    assert_eq!(
        litchi_pptx::animations::Sequence::parse_timing_xml(
            &xml,
            litchi_pptx::animations::Limits::DEFAULT
        )
        .unwrap(),
        sequence
    );
}

#[test]
fn native_powerpoint_variants_read_every_declared_slide_and_keep_package_members() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/powerpoint");
    for (name, count) in [
        ("powerpoint-text-table.pptx", 3),
        ("powerpoint-visual-objects.pptx", 5),
        ("powerpoint-visual-objects.potx", 5),
        ("powerpoint-visual-objects.ppsx", 5),
        ("powerpoint-navigation-review.pptx", 6),
        ("powerpoint-media-objects.pptx", 4),
        ("powerpoint-passive-macro.pptm", 1),
        ("powerpoint-passive-macro.potm", 1),
        ("powerpoint-passive-macro.ppsm", 1),
        ("powerpoint-master-motion.pptx", 7),
        ("powerpoint-modern-media.pptx", 8),
        ("powerpoint-fonts-math-handout.pptx", 4),
        ("powerpoint-ink-stroke.pptx", 1),
        ("powerpoint-ink-replay.pptx", 1),
    ] {
        let source = std::fs::read(root.join(name)).unwrap();
        let package = Package::from_bytes(&source).unwrap();
        let presentation = package.presentation().unwrap();
        assert_eq!(presentation.slide_count().unwrap(), count, "{name}");
        let slides = presentation.slides().unwrap();
        assert_eq!(slides.len(), count);
        for slide in slides {
            assert!(slide.shape_count().unwrap() > 0, "{name}");
            assert!(!slide.text().unwrap().is_empty(), "{name}");
            let _ = slide.is_hidden().unwrap();
        }
        assert_members_equal(&package.to_bytes().unwrap(), &source);
    }
}

fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
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

fn assert_members_equal(left: &[u8], right: &[u8]) {
    let left = members(left);
    let right = members(right);
    assert_eq!(
        left.keys().collect::<Vec<_>>(),
        right.keys().collect::<Vec<_>>()
    );
    for (name, bytes) in left {
        assert!(bytes == right[&name], "changed package member {name}");
    }
}

#[test]
fn native_animation_accepts_separate_same_shape_behaviors_but_not_duplicate_target_children() {
    use litchi_pptx::animations::{Limits, Sequence};
    let package = native("powerpoint-master-motion.pptx");
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let raw = std::str::from_utf8(slides[1].part().part().blob()).unwrap();
    let target = r#"<p:spTgt spid="2"/>"#;
    assert_eq!(raw.matches(target).count(), 3);
    let duplicate = raw.replacen(target, &format!("{target}{target}"), 1);
    assert!(
        matches!(Sequence::parse_slide_xml(duplicate.as_bytes(),Limits::DEFAULT),Err(litchi_pptx::Error::Invalid(message)) if message=="animation has multiple shape targets")
    );
    let conflicting = raw.replacen(target, r#"<p:spTgt spid="3"/>"#, 1);
    assert!(
        matches!(Sequence::parse_slide_xml(conflicting.as_bytes(),Limits::DEFAULT),Err(litchi_pptx::Error::Refused(refusal)) if refusal.kind()==litchi_core::edit::RefusalKind::UnsupportedContent && refusal.operation()=="parse_animation")
    );
    let explicit = raw.replace(r#"presetID="2""#, r#"dur="700" presetID="2""#);
    assert_eq!(
        Sequence::parse_slide_xml(explicit.as_bytes(), Limits::DEFAULT)
            .unwrap()
            .animations[0]
            .duration,
        litchi_pptx::animations::Duration::Finite(700)
    );
}

#[test]
fn native_powerpoint_actions_include_text_hover_and_resolved_internal_slide_targets() {
    use litchi_pptx::actions::{self, Jump, Kind, Trigger};
    let package = native("powerpoint-navigation-review.pptx");
    let opc = package.opc().unwrap();
    let slide = opc
        .get_part(&litchi_opc::PackURI::new("/ppt/slides/slide2.xml").unwrap())
        .unwrap();
    let settings =
        actions::load_slide_action_settings(opc, 1, slide, &mut actions::Limits::default())
            .unwrap();
    assert_eq!(settings.len(), 4);
    assert_eq!(
        settings[0].target().unwrap().external_target(),
        Some("https://example.com/litchi-ppt-fixture")
    );
    assert_eq!(settings[0].tooltip(), Some("EXTERNAL_TOOLTIP_003"));
    assert_eq!(settings[1].kind(), Kind::SlideJump);
    assert_eq!(
        settings[1].target().unwrap().part_name().unwrap().as_str(),
        "/ppt/slides/slide5.xml"
    );
    assert_eq!(
        settings[2].target().unwrap().external_target(),
        Some("mailto:fixture@example.com?subject=PPT_MAIL_SUBJECT_003")
    );
    assert_eq!(settings[3].kind(), Kind::SlideShowJump(Jump::NextSlide));
    let slide = opc
        .get_part(&litchi_opc::PackURI::new("/ppt/slides/slide6.xml").unwrap())
        .unwrap();
    let hover = actions::load_slide_action_settings(opc, 5, slide, &mut actions::Limits::default())
        .unwrap();
    assert_eq!(hover.len(), 1);
    assert_eq!(hover[0].trigger(), Trigger::Hover);
    assert_eq!(hover[0].kind(), Kind::SlideShowJump(Jump::FirstSlide));
}

#[test]
fn native_powerpoint_zooms_keep_summary_slide_and_section_targets() {
    use litchi_pptx::shape::zoom::Zoom;
    let package = native("powerpoint-master-motion.pptx");
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let owner = slides[0].zooms().unwrap();
    assert_eq!(owner.len(), 1);
    let Zoom::Summary(summary) = owner.get(0).unwrap() else {
        panic!("summary zoom")
    };
    assert_eq!(summary.items().len(), 2);
    assert_eq!(
        summary.items()[0].section_id(),
        "{A6EF3FF9-2FC5-7E48-898D-FC1F42C6E0BB}"
    );
    assert_eq!(
        summary.items()[1].section_id(),
        "{21A303EB-537E-B048-B900-3B79F2FC6319}"
    );
    assert_eq!(owner.to_xml().unwrap(), slides[0].part().part().blob());
    let owner = slides[1].zooms().unwrap();
    let Zoom::Slide(zoom) = owner.get(0).unwrap() else {
        panic!("slide zoom")
    };
    assert_eq!(zoom.slide_id(), 260);
    assert!(zoom.properties().return_to_parent());
    let owner = slides[2].zooms().unwrap();
    let Zoom::Section(zoom) = owner.get(0).unwrap() else {
        panic!("section zoom")
    };
    assert_eq!(zoom.section_id(), "{21A303EB-537E-B048-B900-3B79F2FC6319}");
}

fn sha256(bytes: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

#[test]
fn native_powerpoint_model3d_keeps_the_exact_glb_and_raster_preview() {
    let package = native("powerpoint-modern-media.pptx");
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let models = slides[5].model3ds().unwrap();
    assert_eq!(models.len(), 1);
    let model = &models[0];
    let data = model.asset().embedded_data().unwrap().as_slice();
    assert_eq!(data.len(), 1112);
    assert_eq!(&data[..4], b"glTF");
    assert_eq!(
        sha256(data),
        "c38cdec1804bd5c2aef0b8ae55186d9dabbb7831b83d074583d3af2f292ecddb"
    );
    assert!(model.scene().has_raster());
    let preview = model.preview().unwrap();
    assert_eq!(preview.content_type(), Some("image/png"));
    assert_eq!(
        sha256(preview.data().unwrap().as_slice()),
        "fd6b8bff032f6006be7aa80e5eb3149ccfc6c999a06591132e52809b3693c91e"
    );
}

#[test]
fn native_powerpoint_chartex_preserves_waterfall_subtotals_and_sunburst_values() {
    use litchi_pptx::chart::extension::{Dimension, SeriesLayout};
    let package = native("powerpoint-modern-media.pptx");
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    for (slide, layout, expected) in [
        (
            6,
            SeriesLayout::Waterfall,
            vec![100, 25, -10, 5, 120, -20, 30, 130],
        ),
        (7, SeriesLayout::Sunburst, (1..=16).collect()),
    ] {
        let parts = slides[slide].chart_extensions().unwrap();
        assert_eq!(parts.len(), 1);
        let document = parts[0].parse_in_package(package.opc().unwrap()).unwrap();
        let info = document.info();
        assert_eq!(info.series.len(), 1);
        assert_eq!(info.series[0].layout, layout);
        if slide == 6 {
            assert_eq!(
                info.series[0].layout_properties.as_ref().unwrap().subtotals,
                [0, 4, 7]
            );
        }
        let numbers = info
            .data_sets
            .iter()
            .flat_map(|data| &data.dimensions)
            .filter_map(|dimension| match dimension {
                Dimension::Numeric { levels, .. } => Some(levels),
                _ => None,
            })
            .flatten()
            .flat_map(|level| &level.points)
            .map(|point| point.value.parse::<i32>().unwrap())
            .collect::<Vec<_>>();
        assert_eq!(numbers, expected);
        assert!(document.external_data_target().is_some());
    }
}

#[test]
fn native_chartex_subtotal_attributes_reject_ambiguous_or_duplicate_indices() {
    let package = native("powerpoint-modern-media.pptx");
    let uri = litchi_opc::PackURI::new("/ppt/charts/chartEx1.xml").unwrap();
    let part = package.opc().unwrap().get_part(&uri).unwrap();
    let source = std::str::from_utf8(part.blob()).unwrap();
    let needle = r#"<cx:idx val="4" />"#;
    assert_eq!(source.matches(needle).count(), 1);
    for replacement in [
        r#"<cx:idx val="4">4</cx:idx>"#,
        r#"<cx:idx val="0"/>"#,
        r#"<cx:idx val="4294967296"/>"#,
        r#"<cx:idx xmlns:q="urn:wrong" q:val="4"/>"#,
    ] {
        let changed = source.replace(needle, replacement);
        let part = litchi_opc::BlobPart::new(
            uri.clone(),
            part.content_type().into(),
            changed.into_bytes(),
        );
        assert!(
            litchi_pptx::chart::extension::Part::from_part(&part)
                .unwrap()
                .parse()
                .is_err(),
            "{replacement}"
        );
    }
}

#[test]
fn native_model3d_accepts_the_historical_mime_alias_without_changing_payloads() {
    let source = native("powerpoint-modern-media.pptx").to_bytes().unwrap();
    let mut opc =
        litchi_opc::OpcPackage::from_bytes(&source, litchi_opc::SourceLimits::default()).unwrap();
    let uri = litchi_opc::PackURI::new("/ppt/media/model3d1.glb").unwrap();
    let bytes = opc.get_part(&uri).unwrap().blob().to_vec();
    opc.add_part(Box::new(litchi_opc::BlobPart::new(
        uri,
        "model/gltf-binary".into(),
        bytes.clone(),
    )));
    let package = Package::from_opc_package(opc).unwrap();
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let models = slides[5].model3ds().unwrap();
    assert_eq!(models.len(), 1);
    assert_eq!(models[0].asset().embedded_data().unwrap().as_slice(), bytes);
}

#[test]
fn native_powerpoint_video_uses_the_embedded_office_resource_and_reads_caption_cues() {
    use litchi_pptx::presentation_properties::metadata::tracks;
    let package = native("powerpoint-modern-media.pptx");
    let opc = package.opc().unwrap();
    let slide = litchi_opc::PackURI::new("/ppt/slides/slide2.xml").unwrap();
    let media = litchi_pptx::media_parts::load(opc, &slide).unwrap();
    assert_eq!(media.pictures.len(), 1);
    let picture = &media.pictures[0];
    assert_eq!(picture.shape_id, 4);
    assert_eq!(picture.kind, litchi_pptx::media_parts::Kind::Video);
    assert_eq!(picture.relationship_id, "rId1");
    let extension = picture.office_extension.as_ref().unwrap();
    assert_eq!(extension.embed_relationship_id.as_deref(), Some("rId2"));
    assert_eq!(
        extension.trim.as_ref().unwrap().start().duration().unwrap(),
        std::time::Duration::from_millis(50)
    );
    assert_eq!(
        extension
            .fade
            .as_ref()
            .unwrap()
            .fade_in()
            .duration()
            .unwrap(),
        std::time::Duration::from_millis(250)
    );
    assert_eq!(
        extension
            .fade
            .as_ref()
            .unwrap()
            .fade_out()
            .duration()
            .unwrap(),
        std::time::Duration::from_millis(500)
    );
    let resource = picture.resource.as_ref().unwrap();
    assert_eq!(resource.part_name, "/ppt/media/media1.mov");
    assert_eq!(
        sha256(resource.data.as_slice()),
        "e4b92ec504d372e2330b37d4dcd0441d2990096f41f880343795c45508a0c997"
    );
    let tracks = tracks::load(opc, tracks::Limits::DEFAULT).unwrap();
    assert_eq!(tracks.len(), 1);
    let tracks::Target::Internal { part_name, track } = &tracks[0].target else {
        panic!("embedded captions")
    };
    assert_eq!(part_name, "/ppt/media/track1.vtt");
    let cues = track
        .blocks
        .iter()
        .filter_map(|block| match block {
            tracks::Block::Cue(cue) => Some(cue),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(cues.len(), 2);
    for (index, cue) in cues.iter().enumerate() {
        assert_eq!(cue.start_milliseconds, index as u64 * 1000);
        assert_eq!(cue.end_milliseconds, (index as u64 + 1) * 1000);
        assert_eq!(
            cue.payload,
            [if index == 0 {
                "CAPTION_FIRST_007"
            } else {
                "CAPTION_SECOND_007"
            }]
        );
    }
    let metadata = tracks::load_media(
        opc,
        &tracks::MediaKey {
            slide_part_name: slide.as_str().into(),
            shape_id: 4,
        },
        tracks::Limits::DEFAULT,
    )
    .unwrap()
    .unwrap();
    let info = metadata.metadata().tracks_info.as_ref().unwrap();
    assert_eq!(info.display_location, tracks::DisplayLocation::Media);
    assert_eq!(info.captions[0].label, "English Fixture 007");
    assert_eq!(info.captions[0].language.as_deref(), Some("en-us"));
}

#[test]
fn native_powerpoint_svg_without_a_raster_fallback_is_read_as_the_primary_image() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/powerpoint/powerpoint-modern-media.pptx");
    let source = litchi_pptx::SourceBackedPresentation::open(path).unwrap();
    let slide = source.slide(4).unwrap();
    let images = slide.images().unwrap();
    assert_eq!(images.len(), 1);
    assert_eq!(images[0].relationship_id(), "rId2");
    assert_eq!(images[0].svg().unwrap().relationship_id(), "rId2");
    let image = slide.read_image(0).unwrap();
    assert_eq!(image.bytes().len(), 418);
    assert_eq!(
        sha256(image.bytes()),
        "91bdc6ffe9ddcd4ddda5c1e4872f0003b6b0864a0523c19945a4917da6858a95"
    );
    assert_eq!(slide.read_svg_image(0).unwrap().bytes(), image.bytes());
}

#[test]
fn native_vector_only_svg_still_rejects_missing_or_spoofed_relationship_owners() {
    let source = native("powerpoint-modern-media.pptx").to_bytes().unwrap();
    for case in 0..2 {
        let mut opc =
            litchi_opc::OpcPackage::from_bytes(&source, litchi_opc::SourceLimits::default())
                .unwrap();
        let uri = litchi_opc::PackURI::new("/ppt/slides/slide5.xml").unwrap();
        let part = opc.get_part_mut(&uri).unwrap();
        let raw = std::str::from_utf8(part.blob()).unwrap();
        let changed = if case == 0 {
            raw.replace(r#"r:embed="rId2""#, "")
        } else {
            raw.replace(
                "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
                "urn:spoofed",
            )
        };
        assert_ne!(changed, raw);
        part.set_blob(changed.into_bytes());
        let bytes = litchi_opc::PackageWriter::to_bytes(&opc).unwrap();
        let source =
            litchi_pptx::SourceBackedPresentation::from_reader(std::io::Cursor::new(bytes))
                .unwrap();
        assert!(source.slide(4).unwrap().images().is_err(), "case {case}");
    }
}

#[test]
fn native_video_fallback_requires_an_internal_correctly_typed_office_media_edge() {
    let source = native("powerpoint-modern-media.pptx").to_bytes().unwrap();
    let uri = litchi_opc::PackURI::new("/ppt/slides/slide2.xml").unwrap();
    for external in [false, true] {
        let mut opc =
            litchi_opc::OpcPackage::from_bytes(&source, litchi_opc::SourceLimits::default())
                .unwrap();
        let part = opc.get_part_mut(&uri).unwrap();
        part.rels_mut().remove("rId2");
        part.rels_mut().add_relationship(
            if external {
                litchi_opc::constants::relationship_type::MEDIA
            } else {
                litchi_opc::constants::relationship_type::IMAGE
            }
            .into(),
            if external {
                "https://example.invalid/video.mov"
            } else {
                "../media/media1.mov"
            }
            .into(),
            "rId2".into(),
            external,
        );
        assert!(matches!(
            litchi_pptx::media_parts::load(&opc, &uri),
            Err(litchi_pptx::Error::Invalid(_))
        ));
    }
}

#[test]
fn native_powerpoint_sections_shows_notes_and_comment_replies_keep_their_owners() {
    use litchi_pptx::presentation_properties::metadata::structure;
    for (name, section_ids, note_slide) in [
        (
            "powerpoint-navigation-review.pptx",
            vec![vec![256, 257, 258, 259], vec![260, 261]],
            4,
        ),
        (
            "powerpoint-master-motion.pptx",
            vec![vec![262], vec![256, 257, 258, 259], vec![260, 261]],
            5,
        ),
    ] {
        let package = native(name);
        let opc = package.opc().unwrap();
        let graph = structure::load(
            opc,
            structure::Limits::from_source_limits(opc.read_limits()),
        )
        .unwrap();
        assert_eq!(
            graph
                .sections
                .sections()
                .iter()
                .map(|section| section.slide_ids.clone())
                .collect::<Vec<_>>(),
            section_ids
        );
        assert_eq!(graph.custom_shows.shows.len(), 2);
        assert_eq!(graph.custom_shows.shows[0].name, "SHOW_ALPHA_003");
        assert_eq!(graph.custom_shows.shows[0].slide_ids, [256, 257, 260]);
        assert_eq!(graph.custom_shows.shows[1].slide_ids, [261, 259, 256]);
        let notes = package.notes().unwrap().unwrap();
        assert_eq!(notes.slides().len(), 1);
        assert_eq!(
            notes.slides()[0].owner(),
            format!("/ppt/slides/slide{note_slide}.xml")
        );
        assert_eq!(
            notes.slides()[0]
                .text(litchi_pptx::notes::Limits::DEFAULT)
                .unwrap()
                .as_deref(),
            Some(
                format!("PPT_NOTES_003: native navigation and review note.\n{note_slide}").as_str()
            )
        );
        let comments = litchi_pptx::modern_comments::load_modern_comment_graph(opc).unwrap();
        let authors = comments.authors.unwrap().authors.authors;
        assert_eq!(authors.len(), 1);
        assert_eq!(authors[0].name, "Ryker Zhu");
        assert_eq!(authors[0].initials.as_deref(), Some("RZ"));
        assert_eq!(comments.comments.len(), 1);
        assert_eq!(
            comments.comments[0].slide_part_name,
            format!("/ppt/slides/slide{note_slide}.xml")
        );
        let roots = &comments.comments[0].comments.comments;
        assert_eq!(roots.len(), 1);
        assert_eq!(roots[0].author_id, authors[0].id);
        assert_eq!(roots[0].replies.len(), 1);
        assert!(
            std::str::from_utf8(roots[0].text_body_xml.as_ref().unwrap())
                .unwrap()
                .contains("PPT_COMMENT_ROOT_003")
        );
        assert!(
            std::str::from_utf8(roots[0].replies[0].text_body_xml.as_ref().unwrap())
                .unwrap()
                .contains("PPT_COMMENT_REPLY_003")
        );
    }
}

#[test]
fn native_powerpoint_zoom_edit_keeps_graphic_frame_fallback_and_unrelated_members() {
    use litchi_pptx::shape::zoom::Zoom;
    let package = native("powerpoint-master-motion.pptx");
    let before = package.to_bytes().unwrap();
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let mut owner = slides[1].zooms().unwrap();
    let mut value = owner.get(0).unwrap().clone();
    let Zoom::Slide(zoom) = &mut value else {
        panic!("slide zoom")
    };
    let fallback = zoom.fallback_xml().to_vec();
    let blip = zoom.properties().blip_fill_xml().to_vec();
    let geometry = zoom.properties().shape_properties_xml().to_vec();
    zoom.properties_mut().set_return_to_parent(false);
    owner.replace(0, value).unwrap();
    let mut edit = package.edit_zooms().unwrap();
    edit.set(1, owner).unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let view = reopened.presentation().unwrap();
    let slides = view.slides().unwrap();
    let owner = slides[1].zooms().unwrap();
    let Zoom::Slide(zoom) = owner.get(0).unwrap() else {
        panic!("slide zoom")
    };
    assert_eq!(zoom.slide_id(), 260);
    assert!(!zoom.properties().return_to_parent());
    assert_eq!(zoom.fallback_xml(), fallback);
    assert_eq!(zoom.properties().blip_fill_xml(), blip);
    assert_eq!(zoom.properties().shape_properties_xml(), geometry);
    let a = members(&before);
    let b = members(&after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in a {
        if name == "ppt/slides/slide2.xml" {
            let source = std::str::from_utf8(&bytes).unwrap();
            let changed = std::str::from_utf8(&b[&name]).unwrap();
            let frame_start = source.find("<p:graphicFrame").unwrap();
            let payload_start = frame_start + source[frame_start..].find("<pslz:sldZm").unwrap();
            assert!(changed.contains(&source[frame_start..payload_start]));
        } else {
            assert!(b[&name] == bytes, "changed {name}");
        }
    }
    assert!(
        commit
            .package()
            .apply_zooms(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == before
    );
}

#[test]
fn native_powerpoint_model3d_asset_edit_keeps_preview_and_scene_children() {
    use litchi_pptx::model3d::{Asset, Data};
    let package = native("powerpoint-modern-media.pptx");
    let before = package.to_bytes().unwrap();
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let original = slides[5].model3ds().unwrap().remove(0);
    // Native MCE active/fallback shapes share a name. Use the exact source
    // index supplied by the typed model instead of an ambiguous name.
    let shape = original.shape().index();
    let preview = original
        .preview()
        .unwrap()
        .data()
        .unwrap()
        .as_slice()
        .to_vec();
    let mut replacement = original.clone();
    let glb = include_bytes!("../../../test-data/microsoft-office/assets/colored-triangle-007.glb")
        .to_vec();
    replacement.set_asset(Asset::embedded(Data::new(glb.clone()).unwrap()));
    let mut edit = package.edit_model3d().unwrap();
    edit.set(5, shape, replacement).unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let view = reopened.presentation().unwrap();
    let slides = view.slides().unwrap();
    let models = slides[5].model3ds().unwrap();
    assert_eq!(models.len(), 1);
    assert_eq!(models[0].asset().embedded_data().unwrap().as_slice(), glb);
    assert_eq!(
        models[0].preview().unwrap().data().unwrap().as_slice(),
        preview
    );
    let a = members(&before);
    let b = members(&after);
    for (name, bytes) in &a {
        if [
            "[Content_Types].xml",
            "ppt/slides/slide6.xml",
            "ppt/slides/_rels/slide6.xml.rels",
            "ppt/media/model3d1.glb",
        ]
        .contains(&name.as_str())
        {
            continue;
        }
        assert!(b.get(name) == Some(bytes), "changed {name}");
    }
    let before_xml = std::str::from_utf8(&a["ppt/slides/slide6.xml"]).unwrap();
    let after_xml = std::str::from_utf8(&b["ppt/slides/slide6.xml"]).unwrap();
    for local in [
        "camera",
        "trans",
        "raster",
        "objViewport",
        "ambientLight",
        "ptLight",
    ] {
        let start = format!("<am3d:{local}");
        let end = format!("</am3d:{local}>");
        if let Some(begin) = before_xml.find(&start) {
            let tag_end = begin + before_xml[begin..].find('>').unwrap() + 1;
            let close = if before_xml[begin..tag_end].ends_with("/>") {
                tag_end
            } else {
                begin + before_xml[begin..].find(&end).unwrap() + end.len()
            };
            assert!(
                after_xml.contains(&before_xml[begin..close]),
                "changed model scene {local}"
            );
        }
    }
    assert!(
        commit
            .package()
            .apply_model3d(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == before
    );
}

#[cfg(feature = "vba-inspection")]
#[test]
fn native_powerpoint_macro_variants_keep_inert_payloads_through_title_edits() {
    for (extension, hash) in [
        (
            "pptm",
            "44e335f3ac81a2722397df0d6dc489d3de6d49ba0b83805fd94e3268d1d47ce2",
        ),
        (
            "potm",
            "12d30c0bc9bd41c473001f5a08420976e3a27cdee0217775abcb5bc8af57e420",
        ),
        (
            "ppsm",
            "2daf73e707470cdc42e154779c3f63372e90c5891c3f65f19e8d3d5312ff372c",
        ),
    ] {
        let package = native(&format!("powerpoint-passive-macro.{extension}"));
        let before = package.to_bytes().unwrap();
        let project = package.presentation().unwrap().vba().unwrap().unwrap();
        assert_eq!(project.project_part_name().as_str(), "/ppt/vbaProject.bin");
        assert_eq!(project.payload_size(package.opc().unwrap()).unwrap(), 8192);
        assert_eq!(
            sha256(project.payload(package.opc().unwrap()).unwrap()),
            hash
        );
        let mut edit = package.edit().unwrap();
        assert!(
            edit.set_shape_text(0, 0, "PPT_MACRO_005 edited title")
                .unwrap()
        );
        let commit = edit.commit().unwrap();
        let after = commit.package().to_bytes().unwrap();
        let reopened = Package::from_bytes(&after).unwrap();
        assert!(
            reopened
                .presentation()
                .unwrap()
                .text()
                .unwrap()
                .contains("PPT_MACRO_005 edited title")
        );
        let a = members(&before);
        let b = members(&after);
        assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
        for (name, bytes) in a {
            if name != "ppt/slides/slide1.xml" {
                assert!(b[&name] == bytes, "{extension}: changed {name}");
            }
        }
        assert!(
            commit
                .package()
                .apply(&commit.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == before
        );
    }
}

#[cfg(feature = "encryption")]
#[test]
fn native_powerpoint_agile_password_ingress_and_rekey_preserve_clear_members() {
    use litchi_crypto::ooxml::{Error as CryptoError, Mode};
    use litchi_pptx::Error;
    use std::io::Cursor;
    let source =
        include_bytes!("../../../test-data/microsoft-office/powerpoint/powerpoint-encrypted.pptx");
    for password in ["", "wrong"] {
        assert!(matches!(
            Package::from_reader_with_password(Cursor::new(source), password),
            Err(Error::Encryption(CryptoError::Password))
        ));
    }
    let package =
        Package::from_reader_with_password(Cursor::new(source), "OfficeFixture42!").unwrap();
    assert_eq!(package.encryption(), Some(Mode::AgileAes256Sha512));
    assert_eq!(package.presentation().unwrap().slide_count().unwrap(), 1);
    let text = package.presentation().unwrap().text().unwrap();
    for marker in [
        "PPT_ENCRYPTED_008",
        "Fixture ID: powerpoint-encrypted-008",
        "VALUE_008: 42",
    ] {
        assert!(text.contains(marker));
    }
    assert!(matches!(
        package.to_bytes(),
        Err(Error::EncryptionPolicy { .. })
    ));
    let before = package.to_plain_bytes().unwrap();
    let encrypted = package.to_reencrypted("ChangedOfficeFixture42!").unwrap();
    assert!(matches!(
        Package::from_reader_with_password(Cursor::new(&encrypted), "OfficeFixture42!"),
        Err(Error::Encryption(CryptoError::Password))
    ));
    let reopened =
        Package::from_reader_with_password(Cursor::new(encrypted), "ChangedOfficeFixture42!")
            .unwrap();
    assert_eq!(reopened.encryption(), Some(Mode::AgileAes256Sha512));
    assert_members_equal(&before, &reopened.to_plain_bytes().unwrap());
}
