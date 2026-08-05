use litchi_ppt::animation::{AnimationInfo, AnimationSound, BuiltinSound};
use litchi_ppt::writer::PptWriter;
use litchi_ppt::{
    BuiltinId, Interaction, InteractionAction, InteractionLinkTarget, InteractionTrigger, Package,
    PowerPointTextInteraction, PowerPointTextRange,
};
use std::{io::Cursor, num::NonZeroU32};

fn write(writer: &mut PptWriter) -> Result<Vec<u8>, litchi_ppt::writer::PptWriteError> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn minimal_wave() -> Vec<u8> {
    b"RIFF\x04\0\0\0WAVE".to_vec()
}

#[test]
fn action_only_builtin_sounds_are_collected_remapped_and_repeatable() {
    let mut writer = PptWriter::new();
    let slide = writer.add_slide().unwrap();
    writer.add_textbox(slide, 10, 10, 240, 40, "A😀BC").unwrap();

    let click = Interaction::new(
        InteractionTrigger::Click,
        InteractionAction::NoAction,
        InteractionLinkTarget::Nil,
    )
    .with_builtin_sound(BuiltinSound::Click);
    writer.set_last_shape_interaction(slide, click).unwrap();
    let hover = PowerPointTextInteraction::new(
        PowerPointTextRange::new(1, 3).unwrap(),
        Interaction::new(
            InteractionTrigger::MouseOver,
            InteractionAction::NoAction,
            InteractionLinkTarget::Nil,
        )
        .with_builtin_sound(BuiltinSound::Whoosh),
    )
    .unwrap();
    writer
        .set_last_shape_text_interaction(slide, hover)
        .unwrap();

    let first = write(&mut writer).unwrap();
    let second = write(&mut writer).unwrap();
    assert_eq!(first, second, "serialization must not remap writer state");

    let mut package = Package::from_reader(Cursor::new(first)).unwrap();
    let presentation = package.presentation().unwrap();
    presentation
        .validate_interaction_sound_references()
        .unwrap();
    let sounds = presentation.embedded_sounds().unwrap().unwrap();
    assert_eq!(sounds.sound_id_seed, 2);
    assert_eq!(sounds.sounds.len(), 2);

    let slide = &presentation.slides().unwrap()[0];
    let shape = &slide.shape_interactions().unwrap()[0].interactions[0];
    let shape_sound = shape.sound(&sounds).unwrap();
    assert_eq!(shape_sound.name, "Click");
    assert_eq!(shape_sound.builtin_id, Some(BuiltinId::Click));
    shape.validate_sound_collection(&sounds).unwrap();

    let text = &slide.shape_text_interactions().unwrap()[0].interactions[0].interaction;
    let text_sound = text.sound(&sounds).unwrap();
    assert_eq!(text_sound.name, "Whoosh");
    assert_eq!(text_sound.builtin_id, Some(BuiltinId::Whoosh));
    text.validate_sound_collection(&sounds).unwrap();
}

#[test]
fn embedded_animation_sound_can_be_shared_by_shape_action() {
    let mut writer = PptWriter::new();
    let sound_id = writer
        .add_embedded_sound("Custom tone", minimal_wave())
        .unwrap();
    assert_eq!(writer.embedded_sound_count(), 1);
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 10, 10, 240, 40, "Shared custom sound")
        .unwrap();

    let mut animation = AnimationInfo::new();
    animation.sound = Some(AnimationSound::embedded(
        "Custom tone",
        minimal_wave(),
        sound_id.get(),
    ));
    writer.set_shape_animation(slide, 0, animation).unwrap();

    let interaction = Interaction::new(
        InteractionTrigger::Click,
        InteractionAction::NoAction,
        InteractionLinkTarget::Nil,
    )
    .with_sound_reference(sound_id);
    writer
        .set_last_shape_interaction(slide, interaction)
        .unwrap();

    let mut package = Package::from_reader(Cursor::new(write(&mut writer).unwrap())).unwrap();
    let presentation = package.presentation().unwrap();
    presentation
        .validate_interaction_sound_references()
        .unwrap();
    let sounds = presentation.embedded_sounds().unwrap().unwrap();
    assert_eq!(sounds.sounds.len(), 1);
    assert_eq!(sounds.sounds[0].name, "Custom tone");
    assert_eq!(sounds.sounds[0].data, minimal_wave());

    let slide = &presentation.slides().unwrap()[0];
    let interaction = &slide.shape_interactions().unwrap()[0].interactions[0];
    assert_eq!(interaction.sound(&sounds).unwrap().name, "Custom tone");
    let animation = &slide.animations().unwrap()[0].animation;
    let atom = animation.legacy_atom.as_ref().unwrap();
    assert!(atom.has_sound);
    assert_eq!(atom.sound_id_ref, sounds.sounds[0].id);
}

#[test]
fn action_only_embedded_sound_uses_the_explicit_registry() {
    let mut writer = PptWriter::new();
    let sound_id = writer
        .add_embedded_sound("Action tone", minimal_wave())
        .unwrap();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 10, 10, 240, 40, "Custom action sound")
        .unwrap();
    writer
        .set_last_shape_interaction(
            slide,
            Interaction::new(
                InteractionTrigger::Click,
                InteractionAction::NoAction,
                InteractionLinkTarget::Nil,
            )
            .with_sound_reference(sound_id),
        )
        .unwrap();

    let mut package = Package::from_reader(Cursor::new(write(&mut writer).unwrap())).unwrap();
    let presentation = package.presentation().unwrap();
    presentation
        .validate_interaction_sound_references()
        .unwrap();
    let sounds = presentation.embedded_sounds().unwrap().unwrap();
    assert_eq!(sounds.sounds.len(), 1);
    assert_eq!(sounds.sounds[0].name, "Action tone");
    assert_eq!(sounds.sounds[0].data, minimal_wave());

    assert!(writer.remove_embedded_sound(sound_id));
    assert_eq!(writer.embedded_sound_count(), 0);
    assert!(
        write(&mut writer).is_err(),
        "removing a referenced resource must not emit a dangling ID"
    );
}

#[test]
fn missing_invalid_and_linked_sound_resources_fail_before_output() {
    let mut writer = PptWriter::new();
    assert!(writer.add_embedded_sound("Not audio", vec![0; 12]).is_err());
    assert_eq!(
        writer.embedded_sound_count(),
        0,
        "failed registration must not mutate the registry"
    );
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 10, 10, 240, 40, "Bad sound")
        .unwrap();
    let interaction = Interaction::new(
        InteractionTrigger::Click,
        InteractionAction::NoAction,
        InteractionLinkTarget::Nil,
    )
    .with_sound_reference(NonZeroU32::new(99).unwrap());
    writer
        .set_last_shape_interaction(slide, interaction)
        .unwrap();
    assert!(write(&mut writer).is_err());

    let valid = Interaction::new(
        InteractionTrigger::Click,
        InteractionAction::NoAction,
        InteractionLinkTarget::Nil,
    )
    .with_builtin_sound(BuiltinSound::Chime);
    writer.set_last_shape_interaction(slide, valid).unwrap();
    let mut invalid = AnimationInfo::new();
    invalid.sound = Some(AnimationSound::embedded("Not audio", vec![0; 12], 42));
    writer.set_shape_animation(slide, 0, invalid).unwrap();
    assert!(write(&mut writer).is_err());

    let mut linked = AnimationInfo::new();
    linked.sound = Some(AnimationSound::linked(
        "External",
        "https://example.invalid/sound.wav",
        42,
    ));
    writer.set_shape_animation(slide, 0, linked).unwrap();
    assert!(write(&mut writer).is_err());

    writer
        .set_shape_animation(slide, 0, AnimationInfo::new())
        .unwrap();
    assert!(write(&mut writer).is_ok());
}
