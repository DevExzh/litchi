use litchi_ppt::animation::AnimationInfo;
use litchi_ppt::writer::{Hyperlink, Writer};
use litchi_ppt::{
    Interaction, InteractionAction, InteractionJump, InteractionLimits, InteractionLinkTarget,
    InteractionTrigger, Package,
};
use std::io::Cursor;

fn write(writer: &mut Writer) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn typed_click_hover_and_file_link_round_trip() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();

    writer
        .add_textbox(slide, 10, 10, 240, 40, "Macro and program actions")
        .unwrap();
    let mut click = Interaction::new(
        InteractionTrigger::Click,
        InteractionAction::Macro,
        InteractionLinkTarget::Nil,
    )
    .with_macro_name("RefreshDeck")
    .unwrap();
    click.ole_verb = 9;
    click.jump = InteractionJump::LastSlideViewed;
    click.animated = true;
    click.stop_sound = true;
    click.custom_show_return = true;
    click.visited = true;
    click.unused = [0xAA, 0xBB, 0xCC];
    writer
        .set_last_shape_interaction(slide, click.clone())
        .unwrap();

    let hover = Interaction::new(
        InteractionTrigger::MouseOver,
        InteractionAction::RunProgram,
        InteractionLinkTarget::OtherFile,
    )
    .with_macro_name("viewer.exe")
    .unwrap();
    writer
        .set_last_shape_interaction(slide, hover.clone())
        .unwrap();

    writer
        .add_textbox(slide, 10, 60, 240, 40, "File hyperlink")
        .unwrap();
    let file_id = writer.add_hyperlink(Hyperlink::file("report.xlsx"));
    writer.set_last_shape_hyperlink(slide, file_id).unwrap();

    writer
        .add_textbox(slide, 10, 110, 240, 40, "Animated next-slide action")
        .unwrap();
    writer
        .set_shape_animation(slide, 2, AnimationInfo::new())
        .unwrap();
    let mut next = Interaction::new(
        InteractionTrigger::Click,
        InteractionAction::Jump,
        InteractionLinkTarget::NextSlide,
    );
    next.jump = InteractionJump::NextSlide;
    next.animated = true;
    writer
        .set_last_shape_interaction(slide, next.clone())
        .unwrap();

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    let hyperlinks = presentation.hyperlinks().unwrap();
    let entries = presentation.slides().unwrap()[0]
        .shape_interactions()
        .unwrap();
    assert_eq!(entries.len(), 3);

    let paired = entries
        .iter()
        .find(|entry| entry.interactions.len() == 2)
        .unwrap();
    assert_eq!(paired.interactions, vec![click, hover]);

    let file = entries
        .iter()
        .flat_map(|entry| &entry.interactions)
        .find(|interaction| interaction.hyperlink_id == file_id)
        .unwrap();
    assert_eq!(file.action, InteractionAction::Hyperlink);
    assert_eq!(file.link_target, InteractionLinkTarget::OtherFile);
    assert_eq!(
        file.hyperlink(&hyperlinks).unwrap().target.as_deref(),
        Some("report.xlsx")
    );

    assert!(
        entries
            .iter()
            .any(|entry| entry.interactions.as_slice() == std::slice::from_ref(&next))
    );
}

#[test]
fn failed_replacement_is_atomic_and_trigger_removal_is_precise() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 10, 10, 240, 40, "Atomic interaction")
        .unwrap();
    let existing = Interaction::new(
        InteractionTrigger::Click,
        InteractionAction::Macro,
        InteractionLinkTarget::Nil,
    )
    .with_macro_name("Existing")
    .unwrap();
    writer
        .set_last_shape_interaction(slide, existing.clone())
        .unwrap();

    let too_small = InteractionLimits {
        max_record_bytes: 31,
        ..InteractionLimits::default()
    };
    assert!(
        writer
            .set_last_shape_interaction_with_limits(slide, existing.clone(), too_small)
            .is_err()
    );

    let mut dangling = Interaction::new(
        InteractionTrigger::Click,
        InteractionAction::Hyperlink,
        InteractionLinkTarget::Url,
    );
    dangling.hyperlink_id = 999;
    assert!(writer.set_last_shape_interaction(slide, dangling).is_err());

    let mut inconsistent_name = existing.clone();
    inconsistent_name.macro_name = Some("Different".to_string());
    assert!(
        writer
            .set_last_shape_interaction(slide, inconsistent_name)
            .is_err()
    );

    let hover = Interaction::new(
        InteractionTrigger::MouseOver,
        InteractionAction::Ole,
        InteractionLinkTarget::Nil,
    );
    writer.set_last_shape_interaction(slide, hover).unwrap();
    assert!(
        writer
            .clear_last_shape_interaction(slide, InteractionTrigger::MouseOver)
            .unwrap()
    );
    assert!(
        !writer
            .clear_last_shape_interaction(slide, InteractionTrigger::MouseOver)
            .unwrap()
    );

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    let entries = presentation.slides().unwrap()[0]
        .shape_interactions()
        .unwrap();
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].interactions, [existing]);
}
