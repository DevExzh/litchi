//! Add, update, retime, move, or remove an all-at-once Keynote object build.

use std::env;

use litchi_iwa::keynote::{
    BuildAcceleration, BuildStart, KeynoteBuildSettings, KeynoteEditor, KeynoteFlipDirection,
    KeynoteHorizontalBuildDirection, KeynoteJiggleIntensity, KeynoteKeyboardDirection,
    KeynoteMotionPath, KeynoteMotionPathPoint, KeynoteRotationDirection, KeynoteSwooshDirection,
};
use litchi_keynote::{Seconds, build};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(usage())?;
    let output = arguments.next().ok_or(usage())?;
    let slide_index = arguments.next().ok_or(usage())?.parse::<usize>()?;
    let operation = arguments.next().ok_or(usage())?;
    let object_id = arguments.next().ok_or(usage())?.parse::<u64>()?;

    let mut editor = KeynoteEditor::open(input)?;
    match operation.as_str() {
        "add" => {
            let mut settings = match arguments.next().as_deref() {
                None | Some("in") => KeynoteBuildSettings::appear_in(),
                Some("out") => KeynoteBuildSettings::appear_out(),
                Some(_) => return Err(usage().into()),
            };
            if let Some(effect) = arguments.next() {
                settings.set_effect(build::Effect::from_identifier(&effect)?)?;
            }
            if let Some(duration) = arguments.next() {
                settings.set_duration(Seconds::new(duration.parse()?)?)?;
            }
            let build = editor.add_slide_build(slide_index, object_id, settings)?;
            println!(
                "added build {} with chunk {}",
                build.object_id, build.chunks[0].object_id
            );
        },
        "add-rotate" => {
            let total_degrees = arguments.next().ok_or(usage())?.parse()?;
            let direction = match arguments.next().ok_or(usage())?.as_str() {
                "clockwise" => KeynoteRotationDirection::Clockwise,
                "counterclockwise" => KeynoteRotationDirection::Counterclockwise,
                _ => return Err(usage().into()),
            };
            let mut settings = KeynoteBuildSettings::rotate_action(total_degrees, direction);
            if let Some(acceleration) = arguments.next() {
                settings.set_action_acceleration(parse_acceleration(&acceleration)?)?;
            }
            if let Some(duration) = arguments.next() {
                settings.set_duration(Seconds::new(duration.parse()?)?)?;
            }
            let build = editor.add_slide_build(slide_index, object_id, settings)?;
            println!(
                "added Rotate action {} with chunk {}",
                build.object_id, build.chunks[0].object_id
            );
        },
        "add-scale" => {
            let scale_factor = arguments.next().ok_or(usage())?.parse()?;
            let mut settings = KeynoteBuildSettings::scale_action(scale_factor);
            if let Some(acceleration) = arguments.next() {
                settings.set_action_acceleration(parse_acceleration(&acceleration)?)?;
            }
            if let Some(duration) = arguments.next() {
                settings.set_duration(Seconds::new(duration.parse()?)?)?;
            }
            let build = editor.add_slide_build(slide_index, object_id, settings)?;
            println!(
                "added Scale action {} with chunk {}",
                build.object_id, build.chunks[0].object_id
            );
        },
        "add-opacity" => {
            let opacity_percent = arguments.next().ok_or(usage())?.parse()?;
            let mut settings = KeynoteBuildSettings::opacity_action(opacity_percent);
            if let Some(acceleration) = arguments.next() {
                settings.set_action_acceleration(parse_acceleration(&acceleration)?)?;
            }
            if let Some(duration) = arguments.next() {
                settings.set_duration(Seconds::new(duration.parse()?)?)?;
            }
            let build = editor.add_slide_build(slide_index, object_id, settings)?;
            println!(
                "added Opacity action {} with chunk {}",
                build.object_id, build.chunks[0].object_id
            );
        },
        "add-move" => {
            let delta_x = arguments.next().ok_or(usage())?.parse()?;
            let delta_y = arguments.next().ok_or(usage())?.parse()?;
            let align_to_path = match arguments.next().ok_or(usage())?.as_str() {
                "align" => true,
                "no-align" => false,
                _ => return Err(usage().into()),
            };
            let mut settings = KeynoteBuildSettings::move_action(delta_x, delta_y);
            settings.set_move_alignment(align_to_path)?;
            if let Some(acceleration) = arguments.next() {
                settings.set_action_acceleration(parse_acceleration(&acceleration)?)?;
            }
            if let Some(duration) = arguments.next() {
                settings.set_duration(Seconds::new(duration.parse()?)?)?;
            }
            let build = editor.add_slide_build(slide_index, object_id, settings)?;
            println!(
                "added Move action {} with chunk {}",
                build.object_id, build.chunks[0].object_id
            );
        },
        "add-move-bezier" => {
            let delta_x = arguments.next().ok_or(usage())?.parse()?;
            let delta_y = arguments.next().ok_or(usage())?.parse()?;
            let control_1_x = arguments.next().ok_or(usage())?.parse()?;
            let control_1_y = arguments.next().ok_or(usage())?.parse()?;
            let control_2_x = arguments.next().ok_or(usage())?.parse()?;
            let control_2_y = arguments.next().ok_or(usage())?.parse()?;
            let mut path = KeynoteMotionPath::straight(delta_x, delta_y);
            path.subpaths[0].nodes[0].out_control_point =
                KeynoteMotionPathPoint::new(control_1_x, control_1_y);
            path.subpaths[0].nodes[1].in_control_point =
                KeynoteMotionPathPoint::new(control_2_x, control_2_y);
            path.recalculate_natural_size();
            let settings = KeynoteBuildSettings::move_along_path(path);
            let build = editor.add_slide_build(slide_index, object_id, settings)?;
            println!(
                "added Bézier Move action {} with chunk {}",
                build.object_id, build.chunks[0].object_id
            );
        },
        "add-blink" => {
            let repeat_count = arguments.next().ok_or(usage())?.parse()?;
            let fade = parse_toggle(arguments.next().ok_or(usage())?.as_str(), "fade", "no-fade")?;
            add_typed_build(
                &mut editor,
                slide_index,
                object_id,
                KeynoteBuildSettings::blink_action(repeat_count, fade),
                arguments.next(),
                "Blink",
            )?;
        },
        "add-bounce" => {
            let repeat_count = arguments.next().ok_or(usage())?.parse()?;
            let decay = parse_toggle(
                arguments.next().ok_or(usage())?.as_str(),
                "decay",
                "no-decay",
            )?;
            add_typed_build(
                &mut editor,
                slide_index,
                object_id,
                KeynoteBuildSettings::bounce_action(repeat_count, decay),
                arguments.next(),
                "Bounce",
            )?;
        },
        "add-flip" => {
            let repeat_count = arguments.next().ok_or(usage())?.parse()?;
            let direction = match arguments.next().ok_or(usage())?.as_str() {
                "left-to-right" => KeynoteFlipDirection::LeftToRight,
                "right-to-left" => KeynoteFlipDirection::RightToLeft,
                _ => return Err(usage().into()),
            };
            add_typed_build(
                &mut editor,
                slide_index,
                object_id,
                KeynoteBuildSettings::flip_action(repeat_count, direction),
                arguments.next(),
                "Flip",
            )?;
        },
        "add-jiggle" => {
            let intensity = match arguments.next().ok_or(usage())?.as_str() {
                "small" => KeynoteJiggleIntensity::Small,
                "medium" => KeynoteJiggleIntensity::Medium,
                "large" => KeynoteJiggleIntensity::Large,
                _ => return Err(usage().into()),
            };
            add_typed_build(
                &mut editor,
                slide_index,
                object_id,
                KeynoteBuildSettings::jiggle_action(intensity),
                arguments.next(),
                "Jiggle",
            )?;
        },
        "add-pop" => {
            let scale_percent = arguments.next().ok_or(usage())?.parse()?;
            add_typed_build(
                &mut editor,
                slide_index,
                object_id,
                KeynoteBuildSettings::pop_action(scale_percent),
                arguments.next(),
                "Pop",
            )?;
        },
        "add-pulse" => {
            let repeat_count = arguments.next().ok_or(usage())?.parse()?;
            let scale_percent = arguments.next().ok_or(usage())?.parse()?;
            add_typed_build(
                &mut editor,
                slide_index,
                object_id,
                KeynoteBuildSettings::pulse_action(repeat_count, scale_percent),
                arguments.next(),
                "Pulse",
            )?;
        },
        "add-keyboard-in" | "add-keyboard-out" => {
            let direction = match arguments.next().ok_or(usage())?.as_str() {
                "forward" => KeynoteKeyboardDirection::Forward,
                "backward" => KeynoteKeyboardDirection::Backward,
                _ => return Err(usage().into()),
            };
            let show_cursor = parse_toggle(
                arguments.next().ok_or(usage())?.as_str(),
                "cursor",
                "no-cursor",
            )?;
            let settings = if operation == "add-keyboard-in" {
                KeynoteBuildSettings::keyboard_in(direction, show_cursor)
            } else {
                KeynoteBuildSettings::keyboard_out(direction, show_cursor)
            };
            add_typed_build(
                &mut editor,
                slide_index,
                object_id,
                settings,
                arguments.next(),
                "Keyboard",
            )?;
        },
        "add-shimmer-in" | "add-shimmer-out" => {
            let settings = if operation == "add-shimmer-in" {
                KeynoteBuildSettings::shimmer_in()
            } else {
                KeynoteBuildSettings::shimmer_out()
            };
            add_typed_build(
                &mut editor,
                slide_index,
                object_id,
                settings,
                arguments.next(),
                "Shimmer",
            )?;
        },
        "add-dissolve-in" | "add-dissolve-out" => {
            let settings = if operation == "add-dissolve-in" {
                KeynoteBuildSettings::dissolve_in()
            } else {
                KeynoteBuildSettings::dissolve_out()
            };
            add_typed_build(
                &mut editor,
                slide_index,
                object_id,
                settings,
                arguments.next(),
                "Dissolve",
            )?;
        },
        "add-skid-in" | "add-skid-out" | "add-trace-in" | "add-trace-out" => {
            let direction = match arguments.next().ok_or(usage())?.as_str() {
                "left-to-right" => KeynoteHorizontalBuildDirection::LeftToRight,
                "right-to-left" => KeynoteHorizontalBuildDirection::RightToLeft,
                _ => return Err(usage().into()),
            };
            let settings = match operation.as_str() {
                "add-skid-in" => KeynoteBuildSettings::skid_in(direction),
                "add-skid-out" => KeynoteBuildSettings::skid_out(direction),
                "add-trace-in" => KeynoteBuildSettings::trace_in(direction),
                "add-trace-out" => KeynoteBuildSettings::trace_out(direction),
                _ => unreachable!(),
            };
            add_typed_build(
                &mut editor,
                slide_index,
                object_id,
                settings,
                arguments.next(),
                if operation.contains("skid") {
                    "Skid"
                } else {
                    "Trace"
                },
            )?;
        },
        "add-swoosh-in" | "add-swoosh-out" => {
            let direction = match arguments.next().ok_or(usage())?.as_str() {
                "center" => KeynoteSwooshDirection::Center,
                "from-left" => KeynoteSwooshDirection::FromLeft,
                "from-right" => KeynoteSwooshDirection::FromRight,
                _ => return Err(usage().into()),
            };
            let settings = if operation == "add-swoosh-in" {
                KeynoteBuildSettings::swoosh_in(direction)
            } else {
                KeynoteBuildSettings::swoosh_out(direction)
            };
            add_typed_build(
                &mut editor,
                slide_index,
                object_id,
                settings,
                arguments.next(),
                "Swoosh",
            )?;
        },
        "update" => {
            let effect = arguments.next().ok_or(usage())?;
            let duration = Seconds::new(arguments.next().ok_or(usage())?.parse()?)?;
            let build = editor
                .slide_builds(slide_index)?
                .into_iter()
                .find(|build| build.object_id == object_id)
                .ok_or("build is not owned by the requested slide")?;
            let mut settings = build.settings;
            settings.set_effect(build::Effect::from_identifier(&effect)?)?;
            settings.set_duration(duration)?;
            editor.set_slide_build(slide_index, object_id, settings)?;
            println!("updated build {object_id}");
        },
        "move" => {
            let target_index = arguments.next().ok_or(usage())?.parse()?;
            editor.move_slide_build(slide_index, object_id, target_index)?;
            println!("moved build {object_id} to index {target_index}");
        },
        "timing" => {
            let start = match arguments.next().ok_or(usage())?.as_str() {
                "on-click" => BuildStart::OnClick,
                "after-transition" => BuildStart::AfterTransition,
                "with-previous" => BuildStart::WithPrevious,
                "after-previous" => BuildStart::AfterPrevious,
                _ => return Err(usage().into()),
            };
            let delay = Seconds::new(arguments.next().ok_or(usage())?.parse()?)?;
            let build = editor
                .slide_builds(slide_index)?
                .into_iter()
                .find(|build| build.object_id == object_id)
                .ok_or("build is not owned by the requested slide")?;
            let mut settings = build.settings;
            if matches!(start, BuildStart::OnClick | BuildStart::WithPrevious) {
                settings.set_delay(Seconds::ZERO)?;
            }
            settings.set_start(start)?;
            settings.set_delay(delay)?;
            editor.set_slide_build(slide_index, object_id, settings)?;
            println!(
                "retimed build {object_id}: start={start:?}, delay={}",
                delay.as_f64()
            );
        },
        "remove" => {
            editor.remove_slide_build(slide_index, object_id)?;
            println!("removed build {object_id}");
        },
        _ => return Err(usage().into()),
    }
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    editor.save(output)?;
    Ok(())
}

fn usage() -> &'static str {
    "usage: edit_keynote_build <input.key> <output.key> <slide-index> \
     <add|add-rotate|add-scale|add-opacity|add-move|add-move-bezier|add-blink|add-bounce|\
     add-flip|add-jiggle|add-pop|add-pulse|add-keyboard-in|add-keyboard-out|\
     add-dissolve-in|add-dissolve-out|add-shimmer-in|add-shimmer-out|\
     add-skid-in|add-skid-out|\
     add-swoosh-in|add-swoosh-out|add-trace-in|add-trace-out|\
     update|timing|move|remove> \
     <drawable-or-build-id> \
     [operation-specific arguments]\n\
     add-rotate arguments: <total-degrees> <clockwise|counterclockwise> \
     [none|ease-in|ease-out|ease-in-out] [duration]\n\
     add-scale arguments: <scale-factor> \
     [none|ease-in|ease-out|ease-in-out] [duration]\n\
     add-opacity arguments: <opacity-percent> \
     [none|ease-in|ease-out|ease-in-out] [duration]\n\
     add-move arguments: <delta-x> <delta-y> <align|no-align> \
     [none|ease-in|ease-out|ease-in-out] [duration]\n\
     add-move-bezier arguments: <delta-x> <delta-y> \
     <control-1-x> <control-1-y> <control-2-x> <control-2-y>\n\
     add-blink arguments: <repeat-count> <fade|no-fade> [duration]\n\
     add-bounce arguments: <repeat-count> <decay|no-decay> [duration]\n\
     add-flip arguments: <repeat-count> <left-to-right|right-to-left> [duration]\n\
     add-jiggle arguments: <small|medium|large> [duration]\n\
     add-pop arguments: <scale-percent> [duration]\n\
     add-pulse arguments: <repeat-count> <scale-percent> [duration]\n\
     add-keyboard-in/out arguments: <forward|backward> <cursor|no-cursor> [duration]\n\
     add-dissolve-in/out arguments: [duration]\n\
     add-shimmer-in/out arguments: [duration]\n\
     add-skid-in/out and add-trace-in/out arguments: \
     <left-to-right|right-to-left> [duration]\n\
     add-swoosh-in/out arguments: <center|from-left|from-right> [duration]"
}

fn add_typed_build(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    object_id: u64,
    mut settings: KeynoteBuildSettings,
    duration: Option<String>,
    name: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    if let Some(duration) = duration {
        settings.set_duration(Seconds::new(duration.parse()?)?)?;
    }
    let build = editor.add_slide_build(slide_index, object_id, settings)?;
    println!(
        "added {name} build {} with chunk {}",
        build.object_id, build.chunks[0].object_id
    );
    Ok(())
}

fn parse_toggle(value: &str, enabled: &str, disabled: &str) -> Result<bool, &'static str> {
    match value {
        value if value == enabled => Ok(true),
        value if value == disabled => Ok(false),
        _ => Err(usage()),
    }
}

fn parse_acceleration(value: &str) -> Result<BuildAcceleration, &'static str> {
    match value {
        "none" => Ok(BuildAcceleration::None),
        "ease-in" => Ok(BuildAcceleration::EaseIn),
        "ease-out" => Ok(BuildAcceleration::EaseOut),
        "ease-in-out" => Ok(BuildAcceleration::EaseInOut),
        _ => Err(usage()),
    }
}
