//! Structural and resource validation for Word 2010 run effects.

use crate::error::{Error, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

use super::{Effect, MAX_EFFECTS, MAX_OPAQUE_BYTES, OpaqueExtension, RunEffects};

const MAX_OPAQUE_DEPTH: usize = 64;
const MAX_OPAQUE_NODES: usize = 8_192;

/// Validate one complete run-effects collection.
pub(crate) fn validate(value: &RunEffects) -> Result<()> {
    if value.values.len() > MAX_EFFECTS {
        return Err(Error::Invalid(format!(
            "Word run effects exceed {MAX_EFFECTS} children"
        )));
    }
    let mut known = [false; 7];
    for effect in &value.values {
        if let Some(index) = known_index(effect) {
            if known[index] {
                return Err(Error::Invalid(format!(
                    "duplicate Word run effect '{}'",
                    effect.kind().as_str()
                )));
            }
            known[index] = true;
        }
        effect.validate()?;
    }
    Ok(())
}

/// Validate a bounded opaque element without interpreting its vocabulary.
pub(crate) fn validate_opaque(value: &OpaqueExtension) -> Result<()> {
    let xml = value.as_bytes();
    if xml.is_empty() || xml.len() > MAX_OPAQUE_BYTES {
        return Err(Error::Invalid("opaque run effect exceeds its bound".into()));
    }
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut nodes = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(_)) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("opaque XML node counter overflow".into()))?;
                if nodes > MAX_OPAQUE_NODES {
                    return Err(Error::Invalid(format!(
                        "opaque run effect exceeds {MAX_OPAQUE_NODES} XML nodes"
                    )));
                }
                if depth == 0 {
                    roots += 1;
                    if roots > 1 {
                        return Err(Error::Invalid(
                            "opaque run effect contains multiple roots".into(),
                        ));
                    }
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("opaque XML depth overflow".into()))?;
                if depth > MAX_OPAQUE_DEPTH {
                    return Err(Error::Invalid(format!(
                        "opaque run effect exceeds depth {MAX_OPAQUE_DEPTH}"
                    )));
                }
            },
            Ok(Event::Empty(_)) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("opaque XML node counter overflow".into()))?;
                if nodes > MAX_OPAQUE_NODES {
                    return Err(Error::Invalid(format!(
                        "opaque run effect exceeds {MAX_OPAQUE_NODES} XML nodes"
                    )));
                }
                if depth == 0 {
                    roots += 1;
                    if roots > 1 {
                        return Err(Error::Invalid(
                            "opaque run effect contains multiple roots".into(),
                        ));
                    }
                }
            },
            Ok(Event::End(_)) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Invalid("opaque XML has an unexpected end".into()))?;
            },
            Ok(Event::Text(text)) => {
                if depth == 0 && !text.into_inner().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::Invalid(
                        "opaque run effect has text outside its root".into(),
                    ));
                }
            },
            Ok(Event::CData(text)) => {
                if depth == 0 && !text.into_inner().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::Invalid(
                        "opaque run effect has text outside its root".into(),
                    ));
                }
            },
            Ok(Event::Decl(_)) | Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(Error::Invalid(
                    "opaque run effect cannot contain declarations or processing instructions"
                        .into(),
                ));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(Error::Xml(error.to_string())),
        }
    }
    if roots != 1 || depth != 0 {
        return Err(Error::Invalid(
            "opaque run effect must be one complete XML element".into(),
        ));
    }
    Ok(())
}

fn known_index(effect: &Effect) -> Option<usize> {
    Some(match effect {
        Effect::Glow(_) => 0,
        Effect::Shadow(_) => 1,
        Effect::Reflection(_) => 2,
        Effect::TextOutline(_) => 3,
        Effect::TextFill(_) => 4,
        Effect::Scene3d(_) => 5,
        Effect::Props3d(_) => 6,
        Effect::Unknown(_) => return None,
    })
}
