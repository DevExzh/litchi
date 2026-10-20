//! Add, update, retime, or remove an all-at-once Keynote object build.

use std::env;

use litchi_iwa::keynote::{self, BuildAcceleration, BuildStart};
use litchi_keynote::{Seconds, build};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(usage())?;
    let output = arguments.next().ok_or(usage())?;
    let slide_index = arguments.next().ok_or(usage())?.parse::<usize>()?;
    let operation = arguments.next().ok_or(usage())?;
    let object_id = arguments.next().ok_or(usage())?.parse::<u64>()?;

    let mut editor = keynote::Editor::open(input)?;
    match operation.as_str() {
        "add" => {
            let mut settings = match arguments.next().as_deref() {
                None | Some("in") => keynote::BuildSettings::appear_in(),
                Some("out") => keynote::BuildSettings::appear_out(),
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
                "clockwise" => keynote::RotationDirection::Clockwise,
                "counterclockwise" => keynote::RotationDirection::Counterclockwise,
                _ => return Err(usage().into()),
            };
            let mut settings = keynote::BuildSettings::rotate_action(total_degrees, direction);
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
            let mut settings = keynote::BuildSettings::scale_action(scale_factor);
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
            let mut settings = keynote::BuildSettings::opacity_action(opacity_percent);
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
            let mut settings = keynote::BuildSettings::move_action(delta_x, delta_y);
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
            let mut path = keynote::MotionPath::straight(delta_x, delta_y);
            path.subpaths[0].nodes[0].out_control_point =
                keynote::MotionPathPoint::new(control_1_x, control_1_y);
            path.subpaths[0].nodes[1].in_control_point =
                keynote::MotionPathPoint::new(control_2_x, control_2_y);
            path.recalculate_natural_size();
            let settings = keynote::BuildSettings::move_along_path(path);
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
                keynote::BuildSettings::blink_action(repeat_count, fade),
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
                keynote::BuildSettings::bounce_action(repeat_count, decay),
                arguments.next(),
                "Bounce",
            )?;
        },
        "add-flip" => {
            let repeat_count = arguments.next().ok_or(usage())?.parse()?;
            let direction = match arguments.next().ok_or(usage())?.as_str() {
                "left-to-right" => keynote::FlipDirection::LeftToRight,
                "right-to-left" => keynote::FlipDirection::RightToLeft,
                _ => return Err(usage().into()),
            };
            add_typed_build(
                &mut editor,
                slide_index,
                object_id,
                keynote::BuildSettings::flip_action(repeat_count, direction),
                arguments.next(),
                "Flip",
            )?;
        },
        "add-jiggle" => {
            let intensity = match arguments.next().ok_or(usage())?.as_str() {
                "small" => keynote::JiggleIntensity::Small,
                "medium" => keynote::JiggleIntensity::Medium,
                "large" => keynote::JiggleIntensity::Large,
                _ => return Err(usage().into()),
            };
            add_typed_build(
                &mut editor,
                slide_index,
                object_id,
                keynote::BuildSettings::jiggle_action(intensity),
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
                keynote::BuildSettings::pop_action(scale_percent),
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
                keynote::BuildSettings::pulse_action(repeat_count, scale_percent),
                arguments.next(),
                "Pulse",
            )?;
        },
        "add-keyboard-in" | "add-keyboard-out" => {
            let direction = match arguments.next().ok_or(usage())?.as_str() {
                "forward" => keynote::KeyboardDirection::Forward,
                "backward" => keynote::KeyboardDirection::Backward,
                _ => return Err(usage().into()),
            };
            let show_cursor = parse_toggle(
                arguments.next().ok_or(usage())?.as_str(),
                "cursor",
                "no-cursor",
            )?;
            let settings = if operation == "add-keyboard-in" {
                keynote::BuildSettings::keyboard_in(direction, show_cursor)
            } else {
                keynote::BuildSettings::keyboard_out(direction, show_cursor)
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
                keynote::BuildSettings::shimmer_in()
            } else {
                keynote::BuildSettings::shimmer_out()
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
                keynote::BuildSettings::dissolve_in()
            } else {
                keynote::BuildSettings::dissolve_out()
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
                "left-to-right" => keynote::HorizontalBuildDirection::LeftToRight,
                "right-to-left" => keynote::HorizontalBuildDirection::RightToLeft,
                _ => return Err(usage().into()),
            };
            let settings = match operation.as_str() {
                "add-skid-in" => keynote::BuildSettings::skid_in(direction),
                "add-skid-out" => keynote::BuildSettings::skid_out(direction),
                "add-trace-in" => keynote::BuildSettings::trace_in(direction),
                "add-trace-out" => keynote::BuildSettings::trace_out(direction),
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
                "center" => keynote::SwooshDirection::Center,
                "from-left" => keynote::SwooshDirection::FromLeft,
                "from-right" => keynote::SwooshDirection::FromRight,
                _ => return Err(usage().into()),
            };
            let settings = if operation == "add-swoosh-in" {
                keynote::BuildSettings::swoosh_in(direction)
            } else {
                keynote::BuildSettings::swoosh_out(direction)
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
     update|timing|remove> \
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
    editor: &mut keynote::Editor,
    slide_index: usize,
    object_id: u64,
    mut settings: keynote::BuildSettings,
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
