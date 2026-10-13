//! ODF helpers of the ODP adapter (record 0278).
//!
//! The `/metadata` set path for OpenDocument packages, and the conversion
//! of unified lengths to ODF length attributes. Nothing here is ODP
//! specific; the ODT and ODS adapters landed concurrently with their own
//! copies, so this file is a consolidation candidate (record 0278 §8).

use litchi_core::{FileFormat, Metadata};

use crate::error::{Code, Error, Result};
use crate::props::{PropValue, Props};
use crate::target::TargetArg;
use crate::value::{Length, LengthStyle};

/// Whether `target` is the `/metadata` path.
pub(crate) fn is_metadata_target(target: &TargetArg<'_>) -> bool {
    target.path().is_some_and(|path| {
        let path = path.path();
        path.is_absolute()
            && matches!(path.steps(), [step]
                if step.predicates().is_empty() && step.name().eq_ignore_ascii_case("metadata"))
    })
}

/// The metadata keys the retained-source `meta.xml` patcher owns, sorted.
const KEYS: [&str; 5] = ["author", "description", "keywords", "subject", "title"];

/// Apply `/metadata` assignments to `metadata`: `title`, `author`,
/// `subject`, `description` and `keywords` take text, `null` removes.
///
/// # Errors
///
/// `unsupported_property` for any other key (the ODF patcher owns only these
/// five), `invalid_value` for a value that is not text or null.
pub(crate) fn apply_metadata(
    metadata: &mut Metadata,
    props: &Props<'_>,
    format: FileFormat,
) -> Result<()> {
    for (key, value) in props.iter() {
        let key = key.as_str();
        let value = match value {
            PropValue::Null => None,
            other => Some(
                other
                    .as_text()
                    .ok_or_else(|| {
                        Error::new(Code::InvalidValue, "a metadata value must be text or null")
                            .with_format(format)
                            .with_property(key.to_owned())
                    })?
                    .to_owned(),
            ),
        };
        let field = match key.to_ascii_lowercase().as_str() {
            "title" => &mut metadata.title,
            "author" | "creator" => &mut metadata.author,
            "subject" => &mut metadata.subject,
            "description" => &mut metadata.description,
            "keywords" => &mut metadata.keywords,
            _ => {
                return Err(Error::new(
                    Code::UnsupportedProperty,
                    "the OpenDocument metadata patcher owns title, author, subject, description and keywords only",
                )
                .with_format(format)
                .with_property(key.to_owned())
                .with_suggestion(format!("use one of: {}", KEYS.join(", "))));
            },
        };
        *field = value;
    }
    Ok(())
}

/// An ODF length attribute value for `length`: the unified canonical
/// spelling (`2.54cm`, `72pt`, `1in`), or centimetres with six decimals when
/// no two-decimal unit is exact.
pub(crate) fn odf_length(length: Length) -> String {
    let display = length.display(LengthStyle::Auto).to_string();
    if !display.ends_with("emu") {
        return display;
    }
    let micro = i128::from(length.emu()) * 1_000_000 / 360_000;
    let negative = micro < 0;
    let micro = micro.unsigned_abs();
    let mut text = format!(
        "{}{}.{:06}",
        if negative { "-" } else { "" },
        micro / 1_000_000,
        micro % 1_000_000
    );
    while text.ends_with('0') {
        text.pop();
    }
    if text.ends_with('.') {
        text.pop();
    }
    text.push_str("cm");
    text
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn odf_lengths_round_trip() {
        for emu in [
            0, 914_400, 12_700, 360_000, 1_143_360, 4_525_920, -360_001, 1,
        ] {
            let length = Length::from_emu(emu);
            let text = odf_length(length);
            let back = Length::parse(&text).unwrap();
            assert!(
                (back.emu() - emu).abs() <= 1,
                "{emu} -> {text} -> {}",
                back.emu()
            );
        }
        assert_eq!(odf_length(Length::from_emu(914_400)), "2.54cm");
    }
}
