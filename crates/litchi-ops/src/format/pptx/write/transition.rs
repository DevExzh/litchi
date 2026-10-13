//! OfficeCLI's slide `transition` shorthand, `TYPE[-DIR...][-SPEED|-MS]`,
//! parsed into the native typed transition (record 0269).
//!
//! The vocabulary follows OfficeCLI's `transition.json` and its handler
//! (`PowerPointHandler.Animations.cs`, `ApplyTransition`): the type token
//! picks the effect, direction tokens refine it (each effect has its own
//! direction domain and default), `slow`/`medium`/`fast` set the speed and
//! an integer sets the duration in milliseconds (for `wheel`, an integer up
//! to 32 is the spoke count). `none` (or `false`) removes the slide's direct
//! transition. Effects the native model does not represent (`cube`) and
//! values outside a domain are refused, never approximated.
use litchi_pptx::transition::{
    Axis, Corner, FlyThrough, Glitter, GlitterPattern, InOut, Kind, LeftRight, Morph, Ms, Origin,
    Preset, Prism, Reveal, Ripple, Shape, Shred, ShredPattern, Side, Speed, Spokes, Transition,
};

use super::error;
use crate::{Code, Error, Result};

/// The direction words OfficeCLI accepts after the type token.
const DIRECTIONS: &[&str] = &[
    "l",
    "left",
    "r",
    "right",
    "u",
    "up",
    "d",
    "down",
    "lu",
    "leftup",
    "upleft",
    "ru",
    "rightup",
    "upright",
    "ld",
    "leftdown",
    "downleft",
    "rd",
    "rightdown",
    "downright",
    "in",
    "out",
    "h",
    "horiz",
    "horz",
    "horizontal",
    "v",
    "vert",
    "vertical",
    "byobject",
    "object",
    "bypage",
    "byword",
    "word",
    "bychar",
    "char",
    "character",
    "byletter",
];

/// PowerPoint 2013 "Exciting" presets (`p15:prstTrans`), OfficeCLI token to
/// native preset name.
const PRESETS: &[&str] = &[
    "box",
    "fallOver",
    "drape",
    "curtains",
    "wind",
    "prestige",
    "fracture",
    "crush",
    "peelOff",
    "pageCurlDouble",
    "pageCurlSingle",
    "airplane",
    "origami",
];

fn invalid(message: &'static str) -> Error {
    error(Code::InvalidValue, message)
        .with_property("transition")
        .with_suggestion("use TYPE[-DIRECTION][-SPEED|-MS], for example fade, push-right or split-vertical-in, or none")
}

/// Parse `text`; `Ok(None)` removes the transition.
pub(super) fn parse(text: &str) -> Result<Option<Transition>> {
    let text = text.trim();
    if text.eq_ignore_ascii_case("none") || text.eq_ignore_ascii_case("false") {
        return Ok(None);
    }
    let mut parts = text.split('-');
    let kind = parts.next().unwrap_or_default();
    if kind.is_empty() {
        return Err(invalid("a transition needs a type"));
    }
    let mut through_black = false;
    let mut speed = None;
    let mut number: Option<u32> = None;
    let mut directions: Vec<&str> = Vec::new();
    let mut rest = parts.peekable();
    while let Some(part) = rest.next() {
        let lower = part.to_ascii_lowercase();
        if kind.eq_ignore_ascii_case("fade") {
            if matches!(lower.as_str(), "thru" | "through")
                && rest
                    .peek()
                    .is_some_and(|next| next.eq_ignore_ascii_case("black"))
            {
                rest.next();
                through_black = true;
                continue;
            }
            if matches!(lower.as_str(), "thrublack" | "throughblack" | "thrublk") {
                through_black = true;
                continue;
            }
        }
        if !part.is_empty() && part.bytes().all(|byte| byte.is_ascii_digit()) {
            number = Some(
                part.parse()
                    .map_err(|_| invalid("the transition duration is out of range"))?,
            );
        } else if lower == "slow" {
            speed = Some(Speed::Slow);
        } else if lower == "fast" {
            speed = Some(Speed::Fast);
        } else if lower == "medium" || lower == "med" {
            speed = Some(Speed::Medium);
        } else if DIRECTIONS.contains(&lower.as_str()) {
            directions
                .try_reserve(1)
                .map_err(|_| error(Code::InternalError, "allocation failed"))?;
            directions.push(part);
        } else {
            return Err(invalid("unknown transition modifier"));
        }
    }
    let direction = match directions.as_slice() {
        [] => None,
        [one] => Some(*one),
        // Only `split` takes two direction words (orientation and in/out).
        _ if kind.eq_ignore_ascii_case("split") => None,
        _ => return Err(invalid("this transition takes one direction")),
    };
    let lower = kind.to_ascii_lowercase();
    if direction.is_some() && matches!(lower.as_str(), "circle" | "diamond" | "plus" | "wedge") {
        return Err(invalid("this transition takes no direction"));
    }
    let mut duration = number;
    let kind = match lower.as_str() {
        "fade" => Kind::Fade {
            black: through_black.then_some(true),
        },
        "cut" => Kind::Cut { black: None },
        "dissolve" => Kind::Dissolve,
        "circle" => Kind::Shape(Shape::Circle),
        "diamond" => Kind::Shape(Shape::Diamond),
        "plus" => Kind::Shape(Shape::Plus),
        "newsflash" => Kind::Newsflash,
        "random" => Kind::Random,
        "wedge" => Kind::Wedge,
        "wipe" => Kind::Wipe(side(direction.unwrap_or("left"))?),
        "push" => Kind::Push(side(direction.unwrap_or("left"))?),
        "cover" => Kind::Cover(origin(direction.unwrap_or("left"))?),
        "pull" | "uncover" => Kind::Uncover(origin(direction.unwrap_or("right"))?),
        "wheel" | "clock" => {
            let mut count = if lower == "clock" { 1 } else { 4 };
            if lower == "wheel" && duration.is_some_and(|value| (1..=32).contains(&value)) {
                count = duration.take().unwrap_or(4);
            }
            no_direction(direction)?;
            Kind::Wheel(spokes(count)?)
        },
        "zoom" => Kind::Zoom(in_out(direction.unwrap_or("in"))?),
        "split" => split(&directions)?,
        "blinds" | "venetian" => Kind::Blinds(axis(direction.unwrap_or("horizontal"))?),
        "checker" | "checkerboard" => Kind::Checker(axis(direction.unwrap_or("horizontal"))?),
        "comb" => Kind::Comb(axis(direction.unwrap_or("horizontal"))?),
        "bars" | "randombar" => Kind::RandomBars(axis(direction.unwrap_or("horizontal"))?),
        "strips" | "diagonal" => Kind::Strips(corner(direction.unwrap_or("rd"))?),
        "flash" => Kind::Flash,
        "honeycomb" => Kind::Honeycomb,
        "vortex" => Kind::Vortex(side(direction.unwrap_or("left"))?),
        "switch" => Kind::Switch(left_right(direction.unwrap_or("left"))?),
        "flip" => Kind::Flip(left_right(direction.unwrap_or("left"))?),
        "ripple" => {
            no_direction(direction)?;
            Kind::Ripple(Ripple::Center)
        },
        "glitter" => Kind::Glitter(Glitter::new(
            side(direction.unwrap_or("left"))?,
            GlitterPattern::Diamond,
        )),
        "prism" | "rotate" | "orbit" => {
            no_direction(direction)?;
            Kind::Prism(Prism::new(Side::Left, lower != "prism", lower == "orbit"))
        },
        "doors" => Kind::Doors(axis(direction.unwrap_or("horizontal"))?),
        "window" => Kind::Window(axis(direction.unwrap_or("horizontal"))?),
        "shred" => Kind::Shred(Shred::new(
            ShredPattern::Strip,
            in_out(direction.unwrap_or("in"))?,
        )),
        "ferris" => Kind::Ferris(left_right(direction.unwrap_or("left"))?),
        "flythrough" => {
            Kind::FlyThrough(FlyThrough::new(in_out(direction.unwrap_or("in"))?, false))
        },
        "warp" => Kind::Warp(in_out(direction.unwrap_or("in"))?),
        "gallery" => Kind::Gallery(left_right(direction.unwrap_or("left"))?),
        "conveyor" => Kind::Conveyor(left_right(direction.unwrap_or("left"))?),
        "pan" => Kind::Pan(side(direction.unwrap_or("left"))?),
        "reveal" => Kind::Reveal(Reveal::new(left_right(direction.unwrap_or("left"))?, false)),
        "morph" => Kind::Morph(match direction.map(str::to_ascii_lowercase).as_deref() {
            None | Some("byobject" | "object") => Morph::ByObject,
            Some("byword" | "word") => Morph::ByWord,
            Some("bychar" | "char" | "character") => Morph::ByChar,
            Some(_) => return Err(invalid("morph takes byObject, byWord or byChar")),
        }),
        "cube" => {
            return Err(error(
                Code::UnsupportedCapability,
                "the cube transition is not modeled natively",
            )
            .with_property("transition")
            .with_suggestion("use prism, which PowerPoint shows as Cube"));
        },
        _ => {
            let Some(name) = PRESETS
                .iter()
                .find(|preset| preset.eq_ignore_ascii_case(kind))
            else {
                return Err(invalid("unknown transition type"));
            };
            let invert = match direction.map(str::to_ascii_lowercase).as_deref() {
                None | Some("in") => false,
                Some("out") => true,
                Some(_) => return Err(invalid("this transition takes only in or out")),
            };
            let preset = Preset::with_options(name, invert, false)
                .map_err(|_| invalid("unknown transition preset"))?;
            Kind::Preset(preset)
        },
    };
    let mut transition = Transition::new(kind);
    if let Some(speed) = speed {
        transition = transition.with_speed(speed);
    }
    if let Some(duration) = duration {
        let duration =
            Ms::new(duration).map_err(|_| invalid("the transition duration is out of range"))?;
        transition = transition.with_duration(duration);
    }
    Ok(Some(transition))
}

fn no_direction(direction: Option<&str>) -> Result<()> {
    match direction {
        None => Ok(()),
        Some(_) => Err(invalid("this transition takes no direction")),
    }
}

fn side(direction: &str) -> Result<Side> {
    Ok(match direction.to_ascii_lowercase().as_str() {
        "l" | "left" => Side::Left,
        "r" | "right" => Side::Right,
        "u" | "up" => Side::Up,
        "d" | "down" => Side::Down,
        _ => return Err(invalid("expected left, right, up or down")),
    })
}

fn origin(direction: &str) -> Result<Origin> {
    Ok(match direction.to_ascii_lowercase().as_str() {
        "l" | "left" => Origin::Left,
        "r" | "right" => Origin::Right,
        "u" | "up" => Origin::Up,
        "d" | "down" => Origin::Down,
        "lu" | "leftup" | "upleft" => Origin::LeftUp,
        "ru" | "rightup" | "upright" => Origin::RightUp,
        "ld" | "leftdown" | "downleft" => Origin::LeftDown,
        "rd" | "rightdown" | "downright" => Origin::RightDown,
        _ => return Err(invalid("expected a side or a corner")),
    })
}

fn corner(direction: &str) -> Result<Corner> {
    Ok(match direction.to_ascii_lowercase().as_str() {
        "lu" | "leftup" | "upleft" => Corner::LeftUp,
        "ru" | "rightup" | "upright" => Corner::RightUp,
        "ld" | "leftdown" | "downleft" => Corner::LeftDown,
        "rd" | "rightdown" | "downright" => Corner::RightDown,
        _ => return Err(invalid("expected leftup, rightup, leftdown or rightdown")),
    })
}

fn in_out(direction: &str) -> Result<InOut> {
    Ok(match direction.to_ascii_lowercase().as_str() {
        "in" => InOut::In,
        "out" => InOut::Out,
        _ => return Err(invalid("expected in or out")),
    })
}

fn axis(direction: &str) -> Result<Axis> {
    Ok(match direction.to_ascii_lowercase().as_str() {
        "h" | "horiz" | "horz" | "horizontal" => Axis::Horizontal,
        "v" | "vert" | "vertical" => Axis::Vertical,
        _ => return Err(invalid("expected horizontal or vertical")),
    })
}

fn left_right(direction: &str) -> Result<LeftRight> {
    Ok(match direction.to_ascii_lowercase().as_str() {
        "l" | "left" => LeftRight::Left,
        "r" | "right" => LeftRight::Right,
        _ => return Err(invalid("expected left or right")),
    })
}

fn spokes(count: u32) -> Result<Spokes> {
    Ok(match count {
        1 => Spokes::One,
        2 => Spokes::Two,
        3 => Spokes::Three,
        4 => Spokes::Four,
        8 => Spokes::Eight,
        _ => return Err(invalid("a wheel has 1, 2, 3, 4 or 8 spokes")),
    })
}

/// `split[-ORIENTATION][-in|-out]`: a bare split keeps no in/out; an
/// orientation alone means `in`, as OfficeCLI writes it.
fn split(directions: &[&str]) -> Result<Kind> {
    let mut orientation = None;
    let mut toward = None;
    for direction in directions {
        match direction.to_ascii_lowercase().as_str() {
            "in" => toward = Some(InOut::In),
            "out" => toward = Some(InOut::Out),
            other => orientation = Some(axis(other)?),
        }
    }
    if orientation.is_some() && toward.is_none() {
        toward = Some(InOut::In);
    }
    Ok(Kind::Split {
        axis: orientation.unwrap_or(Axis::Horizontal),
        toward,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn kind(text: &str) -> Kind {
        parse(text)
            .expect("parses")
            .expect("a transition")
            .kind()
            .clone()
    }

    #[test]
    fn shorthand_matches_officecli() {
        assert_eq!(kind("fade"), Kind::Fade { black: None });
        assert_eq!(kind("fade-thru-black"), Kind::Fade { black: Some(true) });
        assert_eq!(kind("push"), Kind::Push(Side::Left));
        assert_eq!(kind("push-right"), Kind::Push(Side::Right));
        assert_eq!(kind("uncover"), Kind::Uncover(Origin::Right));
        assert_eq!(kind("cover-rd"), Kind::Cover(Origin::RightDown));
        assert_eq!(
            kind("split"),
            Kind::Split {
                axis: Axis::Horizontal,
                toward: None
            }
        );
        assert_eq!(
            kind("split-vertical-out"),
            Kind::Split {
                axis: Axis::Vertical,
                toward: Some(InOut::Out)
            }
        );
        assert_eq!(
            kind("split-vertical"),
            Kind::Split {
                axis: Axis::Vertical,
                toward: Some(InOut::In)
            }
        );
        assert_eq!(kind("wheel-8"), Kind::Wheel(Spokes::Eight));
        assert_eq!(kind("clock"), Kind::Wheel(Spokes::One));
        assert_eq!(kind("strips-leftup"), Kind::Strips(Corner::LeftUp));
        assert_eq!(kind("morph"), Kind::Morph(Morph::ByObject));
        assert_eq!(kind("morph-byword"), Kind::Morph(Morph::ByWord));
        assert!(matches!(kind("peelOff-out"), Kind::Preset(preset) if preset.invert_x()));
        assert!(matches!(kind("orbit"), Kind::Prism(prism) if prism.content() && prism.inverted()));
        let timed = parse("push-right-1500").unwrap().unwrap();
        assert_eq!(timed.duration().map(Ms::get), Some(1500));
        let slow = parse("fade-slow").unwrap().unwrap();
        assert_eq!(slow.speed(), Speed::Slow);
        assert!(parse("none").unwrap().is_none());
        assert!(parse("FALSE").unwrap().is_none());
    }

    #[test]
    fn every_documented_type_parses() {
        // `transition.json` values, except `cube` (not modeled natively).
        for token in [
            "none",
            "morph",
            "fade",
            "cut",
            "dissolve",
            "circle",
            "diamond",
            "newsflash",
            "plus",
            "random",
            "wedge",
            "wipe",
            "push",
            "cover",
            "pull",
            "uncover",
            "wheel",
            "zoom",
            "box",
            "split",
            "blinds",
            "venetian",
            "checker",
            "checkerboard",
            "comb",
            "bars",
            "randombar",
            "strips",
            "diagonal",
            "flash",
            "honeycomb",
            "vortex",
            "switch",
            "flip",
            "ripple",
            "glitter",
            "prism",
            "rotate",
            "orbit",
            "clock",
            "doors",
            "window",
            "shred",
            "ferris",
            "flythrough",
            "warp",
            "gallery",
            "conveyor",
            "pan",
            "reveal",
            "fallOver",
            "drape",
            "curtains",
            "wind",
            "prestige",
            "fracture",
            "crush",
            "peelOff",
            "pageCurlDouble",
            "pageCurlSingle",
            "airplane",
            "origami",
        ] {
            assert!(parse(token).is_ok(), "{token}");
        }
    }

    #[test]
    fn invalid_shorthand_is_refused() {
        for text in [
            "",
            "sparkle",
            "fade-xyz",
            "circle-left",
            "push-in",
            "wheel-5",
            "morph-left",
        ] {
            let error = parse(text).expect_err(text);
            assert_eq!(error.code(), Code::InvalidValue, "{text}");
            assert_eq!(error.property(), Some("transition"));
        }
        assert_eq!(
            parse("cube").expect_err("cube").code(),
            Code::UnsupportedCapability
        );
    }
}
