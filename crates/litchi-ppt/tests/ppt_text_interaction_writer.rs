use litchi_ppt::writer::text_format::{Paragraph, TextRun};
use litchi_ppt::writer::{Hyperlink, PptWriter};
use litchi_ppt::{
    Interaction, InteractionAction, InteractionJump, InteractionLinkTarget, InteractionTrigger,
    Package, PowerPointTextInteraction, PowerPointTextInteractionLimits, PowerPointTextRange,
};
use std::io::Cursor;

fn write(writer: &mut PptWriter) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn plain_and_rich_utf16_ranges_round_trip_with_shape_actions() {
    let mut writer = PptWriter::new();
    let slide = writer.add_slide().unwrap();
    writer.add_textbox(slide, 10, 10, 240, 40, "A😀BC").unwrap();

    let hyperlink_id = writer.add_hyperlink(Hyperlink::url("https://example.invalid/text"));
    writer
        .set_last_shape_text_hyperlink(slide, PowerPointTextRange::new(1, 3).unwrap(), hyperlink_id)
        .unwrap();
    let hover = PowerPointTextInteraction::new(
        PowerPointTextRange::new(3, 5).unwrap(),
        Interaction::new(
            InteractionTrigger::MouseOver,
            InteractionAction::RunProgram,
            InteractionLinkTarget::OtherFile,
        )
        .with_macro_name("viewer.exe")
        .unwrap(),
    )
    .unwrap();
    writer
        .set_last_shape_text_interaction(slide, hover.clone())
        .unwrap();

    let mut shape_action = Interaction::new(
        InteractionTrigger::Click,
        InteractionAction::Jump,
        InteractionLinkTarget::NextSlide,
    );
    shape_action.jump = InteractionJump::NextSlide;
    writer
        .set_last_shape_interaction(slide, shape_action.clone())
        .unwrap();

    writer
        .add_rich_textbox(
            slide,
            10,
            60,
            240,
            60,
            vec![
                Paragraph::with_runs(vec![TextRun::new("Hi😀").bold()]),
                Paragraph::new("there"),
            ],
        )
        .unwrap();
    let rich = PowerPointTextInteraction::new(
        PowerPointTextRange::new(5, 10).unwrap(),
        Interaction::new(
            InteractionTrigger::Click,
            InteractionAction::Macro,
            InteractionLinkTarget::Nil,
        )
        .with_macro_name("FormatSecondParagraph")
        .unwrap(),
    )
    .unwrap();
    writer
        .set_last_shape_text_interaction(slide, rich.clone())
        .unwrap();

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    let hyperlinks = presentation.hyperlinks().unwrap();
    let slide = &presentation.slides().unwrap()[0];
    let entries = slide.shape_text_interactions().unwrap();
    assert_eq!(entries.len(), 2);

    let paired = entries
        .iter()
        .find(|entry| entry.interactions.len() == 2)
        .unwrap();
    assert_eq!(
        paired.interactions[0].range,
        PowerPointTextRange::new(1, 3).unwrap()
    );
    assert_eq!(
        paired.interactions[0]
            .interaction
            .hyperlink(&hyperlinks)
            .unwrap()
            .target
            .as_deref(),
        Some("https://example.invalid/text")
    );
    assert_eq!(paired.interactions[1], hover);
    assert!(
        entries
            .iter()
            .any(|entry| entry.interactions.as_slice() == std::slice::from_ref(&rich))
    );

    let shape_entries = slide.shape_interactions().unwrap();
    assert_eq!(shape_entries.len(), 1);
    assert_eq!(shape_entries[0].interactions, [shape_action]);
}

#[test]
fn invalid_text_ranges_references_and_limits_are_atomic() {
    let mut writer = PptWriter::new();
    let slide = writer.add_slide().unwrap();
    writer.add_textbox(slide, 10, 10, 240, 40, "ABCDE").unwrap();
    let existing = PowerPointTextInteraction::new(
        PowerPointTextRange::new(1, 3).unwrap(),
        Interaction::new(
            InteractionTrigger::Click,
            InteractionAction::Macro,
            InteractionLinkTarget::Nil,
        )
        .with_macro_name("Existing")
        .unwrap(),
    )
    .unwrap();
    writer
        .set_last_shape_text_interaction(slide, existing.clone())
        .unwrap();

    let outside = PowerPointTextInteraction::new(
        PowerPointTextRange::new(0, 7).unwrap(),
        Interaction::new(
            InteractionTrigger::Click,
            InteractionAction::NoAction,
            InteractionLinkTarget::Nil,
        ),
    )
    .unwrap();
    assert!(
        writer
            .set_last_shape_text_interaction(slide, outside)
            .is_err()
    );

    let mut dangling_action = Interaction::new(
        InteractionTrigger::MouseOver,
        InteractionAction::Hyperlink,
        InteractionLinkTarget::Url,
    );
    dangling_action.hyperlink_id = 99;
    let dangling =
        PowerPointTextInteraction::new(PowerPointTextRange::new(3, 5).unwrap(), dangling_action)
            .unwrap();
    assert!(
        writer
            .set_last_shape_text_interaction(slide, dangling)
            .is_err()
    );

    let second = PowerPointTextInteraction::new(
        PowerPointTextRange::new(3, 5).unwrap(),
        Interaction::new(
            InteractionTrigger::MouseOver,
            InteractionAction::Ole,
            InteractionLinkTarget::Nil,
        ),
    )
    .unwrap();
    assert!(
        writer
            .set_last_shape_text_interaction_with_limits(
                slide,
                second,
                PowerPointTextInteractionLimits {
                    max_interactions: 1,
                    ..Default::default()
                }
            )
            .is_err()
    );
    assert!(
        !writer
            .clear_last_shape_text_interaction(
                slide,
                PowerPointTextRange::new(3, 5).unwrap(),
                InteractionTrigger::MouseOver,
            )
            .unwrap()
    );

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    let entries = presentation.slides().unwrap()[0]
        .shape_text_interactions()
        .unwrap();
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].interactions, [existing]);
}
