//! Set or reset the focused Image inspector controls in a Keynote slide image.
//!
//! Slide selection is semantic (`index:N` or `name:NAME`); images use their
//! checked source-order index. Native object identifiers and archive objects
//! never enter this command-line interface or its diagnostics.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::OsString;
use std::io;
use std::path::{Path, PathBuf};

use litchi_keynote::{
    ImageSelector, Package, SlideImageAdjustmentsError, SlideSelector,
    slide::image::{ImageAdjustment, ImageAdjustments, ImageEnhancement},
};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_image_adjustments <input.key> <output.key> \
                     [index:N|name:SLIDE] [index:N] \
                     [--exposure VALUE|reset] [--saturation VALUE|reset] \
                     [--enhancement on|off|reset]";

enum SelectedSlide {
    Index(usize),
    Name(String),
}

impl SelectedSlide {
    fn selector(&self) -> SlideSelector<'_> {
        match self {
            Self::Index(index) => SlideSelector::index(*index),
            Self::Name(name) => SlideSelector::name(name),
        }
    }
}

#[derive(Clone, Copy)]
enum FieldChange<T> {
    Set(T),
    Reset,
}

#[derive(Clone, Copy, Default)]
struct RequestedChanges {
    exposure: Option<FieldChange<ImageAdjustment>>,
    saturation: Option<FieldChange<ImageAdjustment>>,
    enhancement: Option<FieldChange<ImageEnhancement>>,
}

impl RequestedChanges {
    fn is_empty(self) -> bool {
        self.exposure.is_none() && self.saturation.is_none() && self.enhancement.is_none()
    }

    fn field_count(self) -> usize {
        usize::from(self.exposure.is_some())
            + usize::from(self.saturation.is_some())
            + usize::from(self.enhancement.is_some())
    }

    fn apply(self, before: ImageAdjustments) -> ImageAdjustments {
        let after = match self.exposure {
            Some(FieldChange::Set(value)) => before.with_exposure(Some(value)),
            Some(FieldChange::Reset) => before.with_exposure(None),
            None => before,
        };
        let after = match self.saturation {
            Some(FieldChange::Set(value)) => after.with_saturation(Some(value)),
            Some(FieldChange::Reset) => after.with_saturation(None),
            None => after,
        };
        match self.enhancement {
            Some(FieldChange::Set(value)) => after.with_enhancement(Some(value)),
            Some(FieldChange::Reset) => after.with_enhancement(None),
            None => after,
        }
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    if input == output {
        return Err(invalid_input("input and output paths must differ"));
    }

    let remaining = arguments.collect::<Vec<_>>();
    let mut cursor = 0;
    let slide = if remaining
        .get(cursor)
        .is_some_and(|argument| !is_flag(argument))
    {
        let value = remaining[cursor].clone();
        cursor += 1;
        parse_slide(value)?
    } else {
        SelectedSlide::Index(0)
    };
    let image_index = if remaining
        .get(cursor)
        .is_some_and(|argument| !is_flag(argument))
    {
        let value = remaining[cursor].clone();
        cursor += 1;
        parse_image_index(value)?
    } else {
        0
    };
    let changes = parse_changes(remaining.into_iter().skip(cursor))?;
    if changes.is_empty() {
        return Err(invalid_input(
            "at least one image adjustment option is required",
        ));
    }

    let package = Package::open(&input)
        .map_err(|error| io::Error::other(format!("opening input package failed: {error}")))?;
    package
        .validate()
        .map_err(|error| io::Error::other(format!("validating input package failed: {error}")))?;
    let image = ImageSelector::index(image_index);
    let before = package
        .slide_image_adjustments(slide.selector(), image)
        .map_err(|error| io::Error::other(format!("reading image adjustments failed: {error}")))?;
    let source_bytes = exact_bytes(&package)
        .map_err(|error| io::Error::other(format!("serializing input package failed: {error}")))?;

    let noop = package
        .edit_slide_image_adjustments(slide.selector(), image)
        .map_err(|error| {
            io::Error::other(format!("image-adjustment no-op preflight failed: {error}"))
        })?
        .set(before)
        .map_err(|error| {
            io::Error::other(format!("staging image-adjustment no-op failed: {error}"))
        })?
        .commit()
        .map_err(|error| {
            io::Error::other(format!("image-adjustment no-op commit failed: {error}"))
        })?;
    let noop_value = noop
        .package()
        .slide_image_adjustments(slide.selector(), image)
        .map_err(|error| {
            io::Error::other(format!(
                "reading image-adjustment no-op result failed: {error}"
            ))
        })?;
    let noop_bytes = exact_bytes(noop.package()).map_err(|error| {
        io::Error::other(format!(
            "serializing image-adjustment no-op result failed: {error}"
        ))
    })?;
    if noop.diagnostics().changed()
        || !noop.patch().is_noop()
        || noop_value != before
        || noop_bytes != source_bytes
    {
        return Err(invalid_input(
            "restaging the existing image adjustments was not an exact no-op",
        ));
    }

    let expected = changes.apply(before);
    let commit = package
        .edit_slide_image_adjustments(slide.selector(), image)
        .map_err(|error| io::Error::other(format!("image-adjustment preflight failed: {error}")))?
        .set(expected)
        .map_err(|error| io::Error::other(format!("staging image adjustments failed: {error}")))?
        .commit()
        .map_err(|error| io::Error::other(format!("image-adjustment commit failed: {error}")))?;
    let committed = commit
        .package()
        .slide_image_adjustments(
            SlideSelector::position(commit.patch().slide_position()),
            ImageSelector::position(commit.patch().image_position()),
        )
        .map_err(|error| {
            io::Error::other(format!(
                "reading committed image adjustments failed: {error}"
            ))
        })?;
    if committed != expected {
        return Err(invalid_input(
            "committed package did not expose the requested image adjustments",
        ));
    }
    if commit.diagnostics().changed() != (before != committed) {
        return Err(invalid_input(
            "image-adjustment diagnostics disagreed with the semantic change",
        ));
    }

    let source_conflict_checked = !commit.patch().is_noop();
    if source_conflict_checked
        && !matches!(
            commit
                .package()
                .apply_slide_image_adjustments(commit.patch()),
            Err(SlideImageAdjustmentsError::PatchConflict)
        )
    {
        return Err(invalid_input(
            "forward image-adjustment patch was not rejected on its committed target",
        ));
    }

    let restored = commit
        .package()
        .apply_slide_image_adjustments(&commit.patch().inverse())
        .map_err(|error| io::Error::other(format!("image-adjustment inverse failed: {error}")))?;
    let restored_value = restored
        .package()
        .slide_image_adjustments(slide.selector(), image)
        .map_err(|error| {
            io::Error::other(format!("reading inverse image adjustments failed: {error}"))
        })?;
    let restored_bytes = exact_bytes(restored.package()).map_err(|error| {
        io::Error::other(format!(
            "serializing inverse image-adjustment package failed: {error}"
        ))
    })?;
    if restored_value != before || restored_bytes != source_bytes {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package and image adjustments",
        ));
    }

    save_new(&output, commit.package()).map_err(|error| {
        io::Error::other(format!("writing image-adjustment output failed: {error}"))
    })?;
    let reopened = Package::open(&output).map_err(|error| {
        io::Error::other(format!("reopening image-adjustment output failed: {error}"))
    })?;
    reopened.validate().map_err(|error| {
        io::Error::other(format!(
            "validating reopened image-adjustment output failed: {error}"
        ))
    })?;
    let after = reopened
        .slide_image_adjustments(
            SlideSelector::position(commit.patch().slide_position()),
            ImageSelector::position(commit.patch().image_position()),
        )
        .map_err(|error| {
            io::Error::other(format!(
                "reading reopened image adjustments failed: {error}"
            ))
        })?;
    if after != expected {
        return Err(invalid_input(
            "reopened output did not preserve the requested image adjustments",
        ));
    }

    println!(
        "slide image adjustments: image_index={}, changed_fields={}, before={before:?}, after={after:?}, changed={}, touched_components={}, full_reparse={}, source_conflict_checked={}, source_fingerprint={:016x}, target_fingerprint={:016x}",
        image_index,
        changes.field_count(),
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().full_reparse_performed(),
        source_conflict_checked,
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint(),
    );
    Ok(())
}

fn parse_slide(argument: OsString) -> Result<SelectedSlide, Box<dyn Error>> {
    let value = argument
        .into_string()
        .map_err(|_| invalid_input("slide selector must be valid UTF-8"))?;
    if let Some(index) = value.strip_prefix("index:") {
        return parse_index(index, "slide").map(SelectedSlide::Index);
    }
    if let Some(name) = value.strip_prefix("name:") {
        if name.is_empty() {
            return Err(invalid_input("slide name cannot be empty"));
        }
        return Ok(SelectedSlide::Name(name.to_owned()));
    }
    Err(invalid_input(
        "slide selector must start with index: or name:",
    ))
}

fn parse_image_index(argument: OsString) -> Result<usize, Box<dyn Error>> {
    let value = argument
        .into_string()
        .map_err(|_| invalid_input("image selector must be valid UTF-8"))?;
    let index = value
        .strip_prefix("index:")
        .ok_or_else(|| invalid_input("image selector must use index:N"))?;
    parse_index(index, "image")
}

fn parse_index(value: &str, kind: &str) -> Result<usize, Box<dyn Error>> {
    if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(invalid_input(format!(
            "{kind} index must be a non-negative decimal integer"
        )));
    }
    value
        .parse::<usize>()
        .map_err(|_| invalid_input(format!("{kind} index is too large for this platform")))
}

fn parse_changes(
    arguments: impl IntoIterator<Item = OsString>,
) -> Result<RequestedChanges, Box<dyn Error>> {
    let mut arguments = arguments.into_iter();
    let mut changes = RequestedChanges::default();
    while let Some(flag) = arguments.next() {
        let flag = flag
            .into_string()
            .map_err(|_| invalid_input("adjustment options must be valid UTF-8"))?;
        match flag.as_str() {
            "--exposure" => {
                if changes.exposure.is_some() {
                    return Err(invalid_input("--exposure may be specified only once"));
                }
                changes.exposure = Some(parse_adjustment_change(
                    next_option(&mut arguments, "missing --exposure value")?,
                    "exposure",
                )?);
            },
            "--saturation" => {
                if changes.saturation.is_some() {
                    return Err(invalid_input("--saturation may be specified only once"));
                }
                changes.saturation = Some(parse_adjustment_change(
                    next_option(&mut arguments, "missing --saturation value")?,
                    "saturation",
                )?);
            },
            "--enhancement" => {
                if changes.enhancement.is_some() {
                    return Err(invalid_input("--enhancement may be specified only once"));
                }
                changes.enhancement = Some(parse_enhancement_change(next_option(
                    &mut arguments,
                    "missing --enhancement value",
                )?)?);
            },
            _ => {
                return Err(invalid_input(
                    "unexpected argument; expected --exposure, --saturation, or --enhancement",
                ));
            },
        }
    }
    Ok(changes)
}

fn parse_adjustment_change(
    value: OsString,
    kind: &str,
) -> Result<FieldChange<ImageAdjustment>, Box<dyn Error>> {
    let value = value
        .into_string()
        .map_err(|_| invalid_input(format!("{kind} value must be valid UTF-8")))?;
    if value == "reset" {
        return Ok(FieldChange::Reset);
    }
    let value = value.parse::<f32>().map_err(|_| {
        invalid_input(format!(
            "{kind} must be a finite value in -1.0..=1.0 or reset"
        ))
    })?;
    let value = ImageAdjustment::new(value)
        .map_err(|error| invalid_input(format!("invalid {kind} value: {error}")))?;
    Ok(FieldChange::Set(value))
}

fn parse_enhancement_change(
    value: OsString,
) -> Result<FieldChange<ImageEnhancement>, Box<dyn Error>> {
    let value = value
        .into_string()
        .map_err(|_| invalid_input("enhancement value must be valid UTF-8"))?;
    match value.as_str() {
        "reset" => Ok(FieldChange::Reset),
        "on" | "enabled" => Ok(FieldChange::Set(ImageEnhancement::Enabled)),
        "off" | "disabled" => Ok(FieldChange::Set(ImageEnhancement::Disabled)),
        _ => Err(invalid_input("enhancement must be on, off, or reset")),
    }
}

fn next_option(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<OsString, Box<dyn Error>> {
    arguments.next().ok_or_else(|| invalid_input(message))
}

fn is_flag(argument: &OsString) -> bool {
    argument
        .to_str()
        .is_some_and(|value| value.starts_with("--"))
}

fn required_argument(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<OsString, Box<dyn Error>> {
    arguments.next().ok_or_else(|| invalid_input(message))
}

fn save_new(path: &Path, package: &Package) -> Result<(), Box<dyn Error>> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut temporary = NamedTempFile::new_in(parent)?;
    package.write_to(temporary.as_file_mut())?;
    temporary.as_file().sync_all()?;
    temporary
        .persist_noclobber(path)
        .map_err(|error| -> Box<dyn Error> { Box::new(error.error) })?;
    Ok(())
}

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
