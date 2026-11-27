//! Native PowerPoint hyperlink/action metadata; targets remain inert.

use litchi_core::Position;
use litchi_ppt::{
    Interaction, InteractionAction as Action, InteractionJump as Jump,
    InteractionLinkTarget as LinkTarget, InteractionTrigger as Trigger, Package, TextInteraction,
    hyperlink::{Click, Target as Link},
    package::edit::Target,
};
use std::{collections::BTreeMap, error::Error, io::Cursor};

fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint/powerpoint-navigation-review.ppt"),
    )
    .unwrap()
}

fn must<T, E: Error>(result: Result<T, E>) -> T {
    result.unwrap_or_else(|error| {
        let mut message = error.to_string();
        let mut cause = error.source();
        while let Some(error) = cause {
            message.push_str(&format!(": {error}"));
            cause = error.source();
        }
        panic!("{message}")
    })
}

struct TempPpt(std::path::PathBuf);

impl TempPpt {
    fn new() -> Self {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        Self(std::env::temp_dir().join(format!(
            "litchi-native-navigation-{}-{nonce}.ppt",
            std::process::id()
        )))
    }
}
impl Drop for TempPpt {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.0);
    }
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

fn texts(package: &Package) -> Vec<String> {
    package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|s| s.text().unwrap().to_owned())
        .collect()
}

fn text_actions(package: &Package, slide: usize, shape_id: u32) -> Vec<TextInteraction> {
    package
        .presentation()
        .unwrap()
        .slide_at(slide)
        .unwrap()
        .unwrap()
        .shape_text_interactions()
        .unwrap()
        .into_iter()
        .find(|s| s.shape_id == shape_id)
        .unwrap()
        .interactions
}

fn click(package: &Package, slide: usize, shape_id: u32, start: u32) -> Option<Click> {
    let p = package.presentation().unwrap();
    let actions = text_actions(package, slide, shape_id);
    actions
        .iter()
        .filter(|a| a.range.begin() == start)
        .find_map(|a| p.resolve_click(&a.interaction).unwrap())
}

fn action(id: u32, trigger: Trigger, kind: Action, jump: Jump, target: LinkTarget) -> Interaction {
    Interaction {
        trigger,
        sound_id: 0,
        hyperlink_id: id,
        action: kind,
        ole_verb: 0,
        jump,
        animated: false,
        stop_sound: false,
        custom_show_return: false,
        visited: false,
        link_target: target,
        macro_name: None,
        unused: [0; 3],
        macro_name_data: None,
    }
}

fn preserve_other_records(before: &[u8], after: &[u8], changed_slide: usize) {
    let a = streams(before);
    let b = streams(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (path, data) in &a {
        if path != &["PowerPoint Document"] && path != &["Current User"] {
            assert!(data == &b[path], "changed {path:?}");
        }
    }
    assert!(
        b[&vec!["PowerPoint Document".into()]].starts_with(&a[&vec!["PowerPoint Document".into()]])
    );
    let original = Package::from_bytes(before).unwrap();
    let updated = Package::from_bytes(after).unwrap();
    assert_eq!(texts(&updated), texts(&original));
    let p = original.presentation().unwrap();
    let q = updated.presentation().unwrap();
    let left = p.slides().unwrap();
    let right = q.slides().unwrap();
    assert_eq!(left.len(), right.len());
    for (index, (a, b)) in left.iter().zip(&right).enumerate() {
        assert_eq!(a.slide_id(), b.slide_id());
        if index != changed_slide {
            assert!(a.record().data == b.record().data, "changed slide {index}");
        }
    }
}

#[test]
fn native_ppt_hyperlink_definitions_keep_ids_targets_and_tooltips() {
    let package = Package::from_bytes(&native()).unwrap();
    let p = package.presentation().unwrap();
    let links = p.hyperlinks().unwrap();
    assert_eq!(links.id_seed, Some(37));
    assert_eq!(
        links.hyperlinks.iter().map(|h| h.id).collect::<Vec<_>>(),
        [12, 14, 16, 18, 20, 30, 31, 32, 33, 37]
    );
    for (ids, name, target, location, tip) in [
        (
            [12, 30],
            "EXTERNAL_LINK_003",
            Some("https://example.com/litchi-ppt-fixture"),
            None,
            Some("EXTERNAL_TOOLTIP_003"),
        ),
        (
            [14, 31],
            "DESTINATION_003",
            None,
            Some("260,5,DESTINATION_003"),
            None,
        ),
        (
            [16, 32],
            "MAILTO_LINK_003",
            Some("mailto:fixture@example.com?subject=PPT_MAIL_SUBJECT_003"),
            None,
            None,
        ),
        ([18, 33], "NEXT", None, Some("-1,-1,NEXT"), None),
        ([20, 37], "FIRST", None, Some("-1,-1,FIRST"), None),
    ] {
        for id in ids {
            let value = links.hyperlinks.iter().find(|h| h.id == id).unwrap();
            assert_eq!(value.friendly_name.as_deref(), Some(name));
            assert_eq!(value.target.as_deref(), target);
            assert_eq!(value.location.as_deref(), location);
            if name == "FIRST" {
                assert!(value.extension.is_none());
            } else {
                let extension = value.extension.as_ref().unwrap();
                assert_eq!(extension.screen_tip.as_deref(), tip);
                assert!(extension.inserted_with_dialog);
                assert!(
                    !extension.location_is_named_show && !extension.named_show_returns_to_slide
                );
            }
        }
    }
}

#[test]
fn native_ppt_text_clicks_and_mouse_over_keep_utf16_ranges_and_fields() {
    let package = Package::from_bytes(&native()).unwrap();
    let p = package.presentation().unwrap();
    let slides = p.slides().unwrap();
    assert_eq!(
        slides.iter().map(|s| s.slide_id()).collect::<Vec<_>>(),
        [256, 257, 258, 259, 260, 261]
    );
    assert!(
        slides
            .iter()
            .all(|s| s.shape_interactions().unwrap().is_empty())
    );
    let shape = &slides[1].shapes().unwrap()[1];
    assert_eq!(shape.id(), 4098);
    assert_eq!(shape.name(), Some("Content Placeholder 2"));
    let body = shape.text().unwrap().encode_utf16().collect::<Vec<_>>();
    let actions = text_actions(&package, 1, 4098);
    assert_eq!(actions.len(), 4);
    for (entry, (begin, end, marker, id, kind, jump, target, resolved)) in actions.iter().zip([
        (
            0,
            17,
            "EXTERNAL_LINK_003",
            30,
            Action::Hyperlink,
            Jump::None,
            LinkTarget::Url,
            Link::Url("https://example.com/litchi-ppt-fixture".into()),
        ),
        (
            18,
            35,
            "INTERNAL_JUMP_003",
            31,
            Action::Hyperlink,
            Jump::None,
            LinkTarget::SlideNumber,
            Link::Slide(Position::new(4)),
        ),
        (
            36,
            51,
            "MAILTO_LINK_003",
            32,
            Action::Hyperlink,
            Jump::None,
            LinkTarget::Url,
            Link::Url("mailto:fixture@example.com?subject=PPT_MAIL_SUBJECT_003".into()),
        ),
        (
            52,
            67,
            "ACTION_NEXT_003",
            33,
            Action::Jump,
            Jump::NextSlide,
            LinkTarget::NextSlide,
            Link::Jump(Jump::NextSlide),
        ),
    ]) {
        assert_eq!((entry.range.begin(), entry.range.end()), (begin, end));
        assert_eq!(
            String::from_utf16(&body[begin as usize..end as usize]).unwrap(),
            marker
        );
        assert_eq!(
            entry.interaction,
            action(id, Trigger::Click, kind, jump, target)
        );
        let limits = litchi_ppt::InteractionLimits::default();
        let encoded = entry.interaction.to_bytes(&limits).unwrap();
        assert_eq!(
            Interaction::parse_bytes(&encoded, &limits).unwrap(),
            entry.interaction
        );
        let actual = p.resolve_click(&entry.interaction).unwrap().unwrap();
        assert_eq!(actual.target, resolved);
        assert_eq!(
            actual.tooltip.as_deref(),
            if id == 30 {
                Some("EXTERNAL_TOOLTIP_003")
            } else {
                None
            }
        );
    }
    let hover = text_actions(&package, 5, 9218);
    assert_eq!(hover.len(), 1);
    assert_eq!((hover[0].range.begin(), hover[0].range.end()), (0, 16));
    assert_eq!(
        hover[0].interaction,
        action(
            37,
            Trigger::MouseOver,
            Action::Jump,
            Jump::FirstSlide,
            LinkTarget::FirstSlide
        )
    );
    assert!(p.resolve_click(&hover[0].interaction).unwrap().is_none());
    let limits = litchi_ppt::InteractionLimits::default();
    assert_eq!(
        Interaction::parse_bytes(&hover[0].interaction.to_bytes(&limits).unwrap(), &limits)
            .unwrap(),
        hover[0].interaction
    );
    for index in [0, 2, 3, 4] {
        assert!(slides[index].shape_text_interactions().unwrap().is_empty());
    }
}

#[test]
fn native_ppt_text_link_replace_remove_and_inverse_preserve_sibling_actions() {
    let bytes = native();
    let package = Package::from_bytes(&bytes).unwrap();
    let target = Target::new(Position::new(1), Position::new(1));
    let original = click(&package, 1, 4098, 0).unwrap();
    let mut noop = package.edit().unwrap();
    must(noop.set_text_hyperlink(target, 0, 17, Some(&original)));
    assert!(noop.commit().unwrap().patch().is_noop());
    let desired = Click {
        target: Link::Url("https://example.invalid/edited-native?x=1&y=2".into()),
        tooltip: Some("Edited native tip".into()),
    };
    let mut edit = package.edit().unwrap();
    must(edit.set_text_hyperlink(target, 0, 17, Some(&desired)));
    assert!(matches!(
        edit.set_text_hyperlink(target, 1, 5, Some(&desired)),
        Err(litchi_ppt::package::edit::Error::Native(litchi_ppt::slide_order::Error::Text(error)))
            if matches!(*error, litchi_ppt::text_edit::Error::Refused(litchi_ppt::text_edit::Refusal::DependencyClosure))
    ));
    must(edit.set_text_hyperlink(target, 36, 15, None));
    let (updated, patch) = must(edit.commit()).into_parts();
    let after = updated.to_bytes().unwrap();
    preserve_other_records(&bytes, &after, 1);
    let output = TempPpt::new();
    updated.save(&output.0).unwrap();
    let reopened = Package::open(&output.0).unwrap();
    assert_eq!(click(&reopened, 1, 4098, 0), Some(desired));
    assert_eq!(click(&reopened, 1, 4098, 36), None);
    let old = text_actions(&package, 1, 4098);
    let new = text_actions(&reopened, 1, 4098);
    assert_eq!(new.len(), 3);
    for start in [18, 52] {
        assert_eq!(
            new.iter().find(|a| a.range.begin() == start),
            old.iter().find(|a| a.range.begin() == start)
        );
    }
    let before_links = package.presentation().unwrap().hyperlinks().unwrap();
    let after_links = reopened.presentation().unwrap().hyperlinks().unwrap();
    assert_eq!(after_links.id_seed, Some(38));
    assert_eq!(after_links.hyperlinks.len(), 11);
    assert!(
        before_links
            .hyperlinks
            .iter()
            .all(|h| after_links.hyperlinks.contains(h))
    );
    assert_eq!(
        new.iter()
            .find(|a| a.range.begin() == 0)
            .unwrap()
            .interaction
            .hyperlink_id,
        38
    );
    let restored = reopened.apply(&patch.inverse()).unwrap().into_package();
    assert!(restored.to_bytes().unwrap() == bytes);
    assert!(package.to_bytes().unwrap() == bytes);

    // Durable replay carries the changed persisted records. Its inverse
    // restores record semantics through another append-only publication.
    use litchi_core::patch::{BlobLimits, Patch as DurablePatch, PatchLimits, Reversible};
    let limits = PatchLimits::new(
        BlobLimits::new(8, 2 << 20, 8 << 20),
        1 << 20,
        64,
        8,
        4096,
        8 << 20,
    );
    let durable = patch.to_durable(limits).unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let decoded = DurablePatch::<Reversible>::from_deterministic_json(&wire, limits).unwrap();
    assert_eq!(decoded.to_deterministic_json().unwrap(), wire);
    let replayed = must(package.apply_durable(&decoded)).into_package();
    assert_eq!(text_actions(&replayed, 1, 4098), new);
    assert_eq!(
        replayed.presentation().unwrap().hyperlinks().unwrap(),
        after_links
    );
    let reversed = must(replayed.apply_durable(&decoded.inverse())).into_package();
    assert_eq!(text_actions(&reversed, 1, 4098), old);
    assert_eq!(
        reversed.presentation().unwrap().hyperlinks().unwrap(),
        before_links
    );
    preserve_other_records(&bytes, &reversed.to_bytes().unwrap(), 1);
}

#[test]
fn native_ppt_shape_link_add_remove_preserves_native_text_actions() {
    let bytes = native();
    let package = Package::from_bytes(&bytes).unwrap();
    let target = Target::new(Position::new(0), Position::new(0));
    let desired = Click {
        target: Link::Slide(Position::new(4)),
        tooltip: Some("Native destination".into()),
    };
    let mut edit = package.edit().unwrap();
    must(edit.set_shape_hyperlink(target, Some(&desired)));
    let (updated, patch) = must(edit.commit()).into_parts();
    let after = updated.to_bytes().unwrap();
    preserve_other_records(&bytes, &after, 0);
    assert_eq!(
        updated
            .presentation()
            .unwrap()
            .shape_click(0, 3073)
            .unwrap(),
        Some(desired)
    );
    assert_eq!(
        text_actions(&updated, 1, 4098),
        text_actions(&package, 1, 4098)
    );
    assert_eq!(
        text_actions(&updated, 5, 9218),
        text_actions(&package, 5, 9218)
    );
    let mut remove = updated.edit().unwrap();
    must(remove.set_shape_hyperlink(target, None));
    let removed = must(remove.commit()).into_package();
    assert!(
        removed
            .presentation()
            .unwrap()
            .shape_click(0, 3073)
            .unwrap()
            .is_none()
    );
    assert!(
        updated
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap()
            == bytes
    );
}

#[test]
fn native_ppt_click_addition_keeps_an_existing_mouse_over_action() {
    let bytes = native();
    let package = Package::from_bytes(&bytes).unwrap();
    let target = Target::new(Position::new(5), Position::new(1));
    let hover = text_actions(&package, 5, 9218)[0].clone();
    let desired = Click {
        target: Link::Url("https://example.invalid/native-click".into()),
        tooltip: None,
    };
    let mut edit = package.edit().unwrap();
    must(edit.set_text_hyperlink(target, 0, 16, Some(&desired)));
    let (updated, patch) = must(edit.commit()).into_parts();
    let after = updated.to_bytes().unwrap();
    preserve_other_records(&bytes, &after, 5);
    let actions = text_actions(&updated, 5, 9218);
    assert_eq!(actions.len(), 2);
    assert!(actions.contains(&hover));
    assert_eq!(click(&updated, 5, 9218, 0), Some(desired));
    let mut edit = updated.edit().unwrap();
    must(edit.set_text_hyperlink(target, 0, 16, None));
    let removed = must(edit.commit()).into_package();
    assert_eq!(text_actions(&removed, 5, 9218), [hover]);
    assert!(
        updated
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap()
            == bytes
    );
}

#[test]
fn native_ppt_internal_link_uses_stable_slide_id_after_reorder() {
    let bytes = native();
    let package = Package::from_bytes(&bytes).unwrap();
    let before_links = package.presentation().unwrap().hyperlinks().unwrap();
    let before = package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|s| (s.slide_id(), s.record().data.to_vec()))
        .collect::<BTreeMap<_, _>>();
    let mut edit = package.edit().unwrap();
    must(edit.move_slide(Position::new(4), Position::new(0)));
    let (updated, patch) = must(edit.commit()).into_parts();
    let p = updated.presentation().unwrap();
    let slides = p.slides().unwrap();
    assert_eq!(
        slides.iter().map(|s| s.slide_id()).collect::<Vec<_>>(),
        [260, 256, 257, 258, 259, 261]
    );
    assert_eq!(p.hyperlinks().unwrap(), before_links);
    for slide in &slides {
        assert!(slide.record().data.as_ref() == before[&slide.slide_id()]);
    }
    assert_eq!(
        click(&updated, 2, 4098, 18).unwrap().target,
        Link::Slide(Position::new(0))
    );
    let reopened = Package::from_bytes(&updated.to_bytes().unwrap()).unwrap();
    assert_eq!(
        click(&reopened, 2, 4098, 18).unwrap().target,
        Link::Slide(Position::new(0))
    );
    assert!(
        reopened
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap()
            == bytes
    );
}
