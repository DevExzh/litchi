//! Native PowerPoint master, timing and resource semantics.

use litchi_core::Position;
use litchi_ppt::Package;
use std::collections::BTreeMap;
use std::io::Cursor;

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(name),
    )
    .unwrap()
}

fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let names = path.iter().map(String::as_str).collect::<Vec<_>>();
            let data = ole.open_stream(&names).unwrap();
            (path, data)
        })
        .collect()
}

fn unrelated_streams(before: &[u8], after: &[u8]) {
    let before = streams(before);
    let after = streams(after);
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    for (path, data) in before {
        if path != ["PowerPoint Document"] && path != ["Current User"] {
            assert!(after[&path] == data, "changed {path:?}");
        }
    }
}

#[test]
fn native_fly_in_timing_tree_keeps_conditions_behaviors_and_duration() {
    use litchi_ppt::animation::{ExtendedTimeNode, TimeConditionType, TimeNodeKind};
    fn flatten<'a>(node: &'a ExtendedTimeNode, output: &mut Vec<&'a ExtendedTimeNode>) {
        output.push(node);
        for child in &node.children {
            flatten(child, output);
        }
    }
    let bytes = native("powerpoint-master-motion.ppt");
    let package = Package::from_bytes(&bytes).unwrap();
    let presentation = package.presentation().unwrap();
    let slide = presentation.slide_at(1).unwrap().unwrap();
    let extension = slide.animation_extension().unwrap().unwrap();
    let root = extension.time_node.as_ref().unwrap();
    let mut nodes = Vec::new();
    flatten(root, &mut nodes);
    assert_eq!(nodes.len(), 8);
    let sequence = nodes
        .iter()
        .find(|n| n.atom.node_type == Some(TimeNodeKind::Sequential))
        .unwrap();
    assert_eq!(sequence.begin_conditions.len(), 1);
    assert_eq!(
        sequence.begin_conditions[0].condition_type,
        TimeConditionType::Next
    );
    assert_eq!(sequence.end_conditions.len(), 1);
    assert_eq!(
        sequence.end_conditions[0].condition_type,
        TimeConditionType::Previous
    );
    assert_eq!(
        nodes
            .iter()
            .filter(|n| n.atom.duration_ms == Some(1250))
            .count(),
        2
    );
    use litchi_ppt::animation::{
        BuildListEntry, ParagraphBuildType, TimeNodeBehavior, TimeVariantValue, TimeVisualElement,
    };
    let animations = nodes
        .iter()
        .filter_map(|n| match n.behavior.as_ref() {
            Some(TimeNodeBehavior::Animate(value)) => Some(value),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(animations.len(), 2);
    let mut empty_formula_mode = animations[0].clone();
    empty_formula_mode.atom.calculation_mode =
        Some(litchi_ppt::animation::TimeAnimateCalculationMode::Formula);
    assert!(litchi_ppt::animation::write_time_animate_behavior(&empty_formula_mode).is_err());
    for (animation, (attribute, values)) in animations.iter().zip([
        ("ppt_x", ["0-#ppt_w/2", "#ppt_x"]),
        ("ppt_y", ["#ppt_y", "#ppt_y"]),
    ]) {
        assert_eq!(
            animation.behavior.attribute_names.as_deref(),
            Some([attribute.to_string()].as_slice())
        );
        assert!(matches!(
            animation.behavior.target,
            TimeVisualElement::Shape {
                shape_id_ref: 8193,
                ..
            }
        ));
        let entries = &animation.values.as_ref().unwrap().entries;
        assert_eq!(entries.len(), 2);
        for (entry, (time, value)) in entries.iter().zip([0, 1000].into_iter().zip(values)) {
            assert_eq!(entry.time, time);
            assert_eq!(entry.value, Some(TimeVariantValue::String(value.into())));
            assert_eq!(entry.formula.as_deref(), Some(""));
        }
    }
    let builds = &extension.build_list.as_ref().unwrap().builds;
    assert_eq!(builds.len(), 1);
    let BuildListEntry::Paragraph(build) = &builds[0] else {
        panic!("paragraph build")
    };
    assert_eq!(
        (
            build.atom.build_id,
            build.atom.shape_id_ref,
            build.atom.expanded
        ),
        (0, 8193, true)
    );
    assert_eq!(build.paragraph.build_type, ParagraphBuildType::AsAWhole);
    assert!(build.levels.is_empty());
    let encoded = litchi_ppt::animation::write_extended_time_node(root).unwrap();
    let (record, consumed) =
        litchi_ppt::Record::parse(&encoded, 0, &litchi_ppt::SourceLimits::default()).unwrap();
    assert_eq!(consumed, encoded.len());
    assert_eq!(
        litchi_ppt::animation::parse_extended_time_node(&record).unwrap(),
        *root
    );
}

#[test]
fn native_transitions_and_advance_edits_preserve_other_slides_and_resources() {
    use litchi_ppt::slide_order::{SlideAdvance, SlideTransitionVisual};
    use litchi_ppt::{
        AdvanceMode, TransitionDirection as Direction, TransitionSpeed as Speed,
        TransitionType as Kind,
    };
    let bytes = native("powerpoint-master-motion.ppt");
    let source = Package::from_bytes(&bytes).unwrap();
    let presentation = source.presentation().unwrap();
    let slide = presentation.slide_at(1).unwrap().unwrap();
    let original = slide.transition().unwrap().unwrap();
    assert_eq!(
        (
            original.transition_type,
            original.direction,
            original.speed,
            original.advance_mode,
            original.advance_time_ms
        ),
        (
            Kind::Push,
            Direction::FromTop,
            Speed::Medium,
            AdvanceMode::Both,
            Some(3000)
        )
    );
    assert!(original.sound.is_none());
    assert!(!original.loop_sound);
    let morph = presentation
        .slide_at(2)
        .unwrap()
        .unwrap()
        .transition()
        .unwrap()
        .unwrap();
    assert_eq!(morph.transition_type, Kind::AlphaFade); // The saved PPT fallback, not Morph playback.
    assert!(
        presentation
            .slide_at(3)
            .unwrap()
            .unwrap()
            .timing()
            .unwrap()
            .hidden
    );
    let before_records = presentation
        .slides()
        .unwrap()
        .iter()
        .map(|s| s.record().data.to_vec())
        .collect::<Vec<_>>();
    let mut edit = source.edit().unwrap();
    edit.set_slide_transition_visual(
        Position::new(1),
        SlideTransitionVisual::new(Kind::Cover, Direction::FromLeft, Speed::Fast).unwrap(),
    )
    .unwrap();
    edit.set_slide_advance(
        Position::new(1),
        SlideAdvance::new(true, true, 4500).unwrap(),
    )
    .unwrap();
    let committed = edit.commit().unwrap();
    let after = committed.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let output = reopened.presentation().unwrap();
    let actual = output
        .slide_at(1)
        .unwrap()
        .unwrap()
        .transition()
        .unwrap()
        .unwrap();
    assert_eq!(
        (
            actual.transition_type,
            actual.direction,
            actual.speed,
            actual.advance_mode,
            actual.advance_time_ms
        ),
        (
            Kind::Cover,
            Direction::FromLeft,
            Speed::Fast,
            AdvanceMode::Both,
            Some(4500)
        )
    );
    assert_eq!(actual.sound, original.sound);
    assert_eq!(actual.loop_sound, original.loop_sound);
    for (index, slide) in output.slides().unwrap().iter().enumerate() {
        if index != 1 {
            assert!(
                slide.record().data.as_ref() == before_records[index],
                "slide {index}"
            );
        }
    }
    assert_eq!(
        output
            .slide_at(1)
            .unwrap()
            .unwrap()
            .animation_extension()
            .unwrap(),
        slide.animation_extension().unwrap()
    );
    let outside_transition = |record: &litchi_ppt::Record| {
        record
            .children
            .iter()
            .filter(|child| child.record_type != litchi_ppt::RecordType::SSSlideInfoAtom)
            .map(|child| (child.record_type_raw, child.instance, child.data.to_vec()))
            .collect::<Vec<_>>()
    };
    assert_eq!(
        outside_transition(slide.record()),
        outside_transition(output.slide_at(1).unwrap().unwrap().record())
    );
    unrelated_streams(&bytes, &after);
    assert!(
        reopened
            .apply(&committed.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
}

#[test]
fn native_master_inventory_styles_and_theme_keep_exact_owners() {
    use litchi_ppt::{MainMasterMetadata12, RecordType, TextMasterStyle};
    let bytes = native("powerpoint-master-motion.ppt");
    let source = Package::from_bytes(&bytes).unwrap();
    let p = source.presentation().unwrap();
    let masters = p.masters().unwrap();
    assert_eq!(masters.main().len(), 5);
    assert!(masters.title().is_empty());
    assert_eq!(masters.notes().unwrap().persist().id(), 2);
    assert!(masters.handout().is_none());
    for (master, (id, persist, layouts)) in masters.main().iter().zip([
        (2147483660, 3, 13),
        (2147483678, 4, 1),
        (2147483679, 5, 1),
        (2147483680, 6, 1),
        (2147483681, 7, 1),
    ]) {
        assert_eq!(
            (master.identity().master_id(), master.persist().id()),
            (id, persist)
        );
        let styles = master
            .record()
            .children
            .iter()
            .filter(|r| r.record_type == RecordType::TxMasterStyleAtom)
            .map(|r| TextMasterStyle::parse(&r.data, r.instance).unwrap())
            .collect::<Vec<_>>();
        assert_eq!(
            styles.iter().map(|s| s.text_type).collect::<Vec<_>>(),
            [0, 1, 2, 4, 5, 6, 7, 8]
        );
        assert!(styles.iter().all(|s| !s.levels.is_empty()));
        let roundtrip = MainMasterMetadata12::parse(master.record()).unwrap();
        assert_eq!(roundtrip.content_masters.len(), layouts);
        assert!(roundtrip.theme_package.is_some());
        assert_eq!(roundtrip.text_styles.len(), 1);
    }
    let theme = p.theme().unwrap().unwrap();
    assert_eq!(theme.name.as_deref(), Some("Ion"));
    assert_eq!(theme.major.latin.as_deref(), Some("Century Gothic"));
    assert_eq!(theme.minor.latin.as_deref(), Some("Century Gothic"));
    assert_eq!(
        theme.colors,
        [
            [0, 0, 0],
            [255, 255, 255],
            [30, 81, 85],
            [235, 235, 235],
            [176, 21, 19],
            [234, 99, 18],
            [230, 183, 41],
            [106, 172, 144],
            [84, 132, 154],
            [158, 94, 155],
            [88, 193, 186],
            [157, 255, 203],
        ]
        .map(Some)
    );
    let headers = p.header_footers().unwrap();
    let slides = headers.presentation_slides().unwrap();
    assert_eq!(slides.display_footer(), Some("PPT_FOOTER_003"));
    assert_eq!(slides.display_user_date(), Some("2026-01-15"));
    assert_eq!(slides.options.datetime_format.get(), 13);
}

#[test]
fn native_master_design_name_edit_preserves_slide_and_other_master_records() {
    use litchi_ppt::master::Kind;
    let bytes = native("powerpoint-master-motion.ppt");
    let source = Package::from_bytes(&bytes).unwrap();
    let p = source.presentation().unwrap();
    let before = p.masters().unwrap();
    let mut edit = source
        .edit_master_metadata(Kind::Main, Position::new(0))
        .unwrap();
    edit.set_template_name("Native reviewed design").unwrap();
    let committed = edit.commit().unwrap();
    let after = committed.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    assert_eq!(
        reopened
            .edit_master_metadata(Kind::Main, Position::new(0))
            .unwrap()
            .template_name()
            .unwrap()
            .unwrap()
            .as_str(),
        "Native reviewed design"
    );
    let out = reopened.presentation().unwrap();
    let masters = out.masters().unwrap();
    assert_eq!(masters.main().len(), before.main().len());
    for (index, (left, right)) in before.main().iter().zip(masters.main()).enumerate() {
        assert_eq!(left.identity().master_id(), right.identity().master_id());
        if index != 0 {
            assert_eq!(left.record().data, right.record().data);
        } else {
            let outside_name = |record: &litchi_ppt::Record| {
                record
                    .children
                    .iter()
                    .filter(|child| {
                        !(child.record_type == litchi_ppt::RecordType::CString
                            && child.instance == 2)
                    })
                    .map(|child| (child.record_type_raw, child.instance, child.data.to_vec()))
                    .collect::<Vec<_>>()
            };
            assert_eq!(outside_name(left.record()), outside_name(right.record()));
        }
    }
    for (left, right) in p.slides().unwrap().iter().zip(out.slides().unwrap()) {
        assert_eq!(left.record().data, right.record().data);
    }
    unrelated_streams(&bytes, &after);
    assert!(
        reopened
            .apply_master_metadata(&committed.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
}

#[test]
fn native_handout_context_and_legacy_font_conversion_are_explicit() {
    let bytes = native("powerpoint-fonts-math-handout.ppt");
    let source = Package::from_bytes(&bytes).unwrap();
    let p = source.presentation().unwrap();
    let masters = p.masters().unwrap();
    assert_eq!(masters.main().len(), 1);
    assert!(masters.notes().is_none());
    let handout = masters.handout().unwrap();
    assert_eq!(handout.persist().id(), 2);
    assert_eq!(
        handout.record().record_type,
        litchi_ppt::RecordType::Handout
    );
    assert_eq!(
        handout
            .record()
            .children
            .iter()
            .map(|r| r.record_type)
            .collect::<Vec<_>>(),
        [
            litchi_ppt::RecordType::PPDrawing,
            litchi_ppt::RecordType::Unknown,
            litchi_ppt::RecordType::ProgTags,
            litchi_ppt::RecordType::RoundTripTheme12Atom,
            litchi_ppt::RecordType::RoundTripColorMapping12Atom,
        ]
    );
    assert_eq!(handout.record().children[1].record_type_raw, 2032); // ColorSchemeAtom.
    let fonts = p.fonts().unwrap();
    let base = fonts.base.as_ref().unwrap();
    assert_eq!(
        base.fonts
            .iter()
            .map(|f| f.name.as_str())
            .collect::<Vec<_>>(),
        ["Aptos", "Arial", "Aptos Display"]
    );
    let flags = fonts.embedding_flags.unwrap();
    assert_eq!(flags.raw, 1);
    assert!(flags.subset);
    assert!(!flags.subset_option_confirmed);
    assert!(!fonts.has_embedded_fonts());
    assert!(fonts.international.is_none());
    assert!(source.to_bytes().unwrap() == bytes);
}

#[test]
fn native_excel_chart_inventory_keeps_frame_series_and_all_cached_points() {
    use litchi_ograph::chart::cache::Index;
    use litchi_ograph::chart::{Cache, Family, XlValue};
    let bytes = native("powerpoint-visual-objects.ppt");
    let source = Package::from_bytes(&bytes).unwrap();
    let p = source.presentation().unwrap();
    let charts = p.charts().unwrap();
    assert_eq!(
        (charts.seen(), charts.len(), charts.failures().count()),
        (1, 1, 0)
    );
    let chart = charts.get(0).unwrap();
    assert_eq!(chart.info().program(), Some("Excel.Chart.8"));
    assert_eq!(chart.info().object_id(), 1);
    assert_eq!(chart.info().persist_id(), 2);
    assert_eq!(
        (
            chart.info().frame().unwrap().slide(),
            chart.info().frame().unwrap().shape()
        ),
        (4, 5122)
    );
    let model = chart.semantic_chart(0).unwrap().unwrap();
    assert_eq!(
        model
            .series()
            .iter()
            .map(|s| s.ai.get(litchi_ograph::chart::Role::Name).text().unwrap())
            .collect::<Vec<_>>(),
        ["North", "South", "West"]
    );
    assert_eq!(model.groups().len(), 1);
    assert!(matches!(
        model.groups()[0].family,
        Family::Bar { flags: 0, .. }
    ));
    assert_eq!(model.caches().len(), 24);
    let expected = [
        [10.0, 15.0, 8.0],
        [20.0, 18.0, 12.0],
        [30.0, 21.0, 16.0],
        [40.0, 24.0, 20.0],
    ];
    for item in model.caches() {
        let Cache::Excel {
            section,
            row,
            col,
            value,
            ..
        } = item
        else {
            panic!("Excel cache")
        };
        match section {
            Index::Categories => assert_eq!(value, &XlValue::Text(format!("Q{}", row + 1))),
            Index::Values => assert_eq!(
                value,
                &XlValue::Number(expected[usize::from(*row)][usize::from(*col)])
            ),
            _ => panic!("unexpected cache section"),
        }
    }
}

#[test]
fn native_animation_duration_edit_preserves_targets_builds_and_other_slides() {
    use litchi_ppt::animation::ExtendedTimeNode;
    fn change(node: &mut ExtendedTimeNode) -> usize {
        let mut changed = 0;
        if node.atom.duration_ms == Some(1250) {
            node.atom.duration_ms = Some(1500);
            changed += 1;
        }
        changed + node.children.iter_mut().map(change).sum::<usize>()
    }
    let bytes = native("powerpoint-master-motion.ppt");
    let source = Package::from_bytes(&bytes).unwrap();
    let before = source.presentation().unwrap();
    let slide = before.slide_at(1).unwrap().unwrap();
    let persist = slide.persist_id();
    let mut edit = source.edit_animations().unwrap();
    let timeline = edit.find(persist).unwrap().unwrap();
    let mut expected = timeline.extension.clone();
    let root = expected.time_node.as_mut().unwrap();
    assert_eq!(root.children.len(), 1);
    assert_eq!(change(&mut root.children[0]), 2);
    edit.update(persist, 0, root.children[0].clone()).unwrap();
    let committed = edit.commit().unwrap();
    let after = committed.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let output = reopened.presentation().unwrap();
    assert_eq!(
        output
            .slide_at(1)
            .unwrap()
            .unwrap()
            .animation_extension()
            .unwrap(),
        Some(&expected)
    );
    for (index, (left, right)) in before
        .slides()
        .unwrap()
        .iter()
        .zip(output.slides().unwrap())
        .enumerate()
    {
        if index != 1 {
            assert_eq!(left.record().data, right.record().data);
        } else {
            let untouched = |r: &litchi_ppt::Record| {
                r.children
                    .iter()
                    .filter(|c| c.record_type != litchi_ppt::RecordType::ProgTags)
                    .map(|c| (c.record_type_raw, c.instance, c.data.to_vec()))
                    .collect::<Vec<_>>()
            };
            assert_eq!(untouched(left.record()), untouched(right.record()));
        }
    }
    unrelated_streams(&bytes, &after);
    assert!(
        reopened
            .apply_animations(&committed.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    assert!(source.to_bytes().unwrap() == bytes);
}
