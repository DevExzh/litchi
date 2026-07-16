//! Typed paragraph-property inheritance with cycle and depth guards.

use std::collections::HashSet;

use crate::protobuf::tswp;
use crate::text::paragraph_tabs::ParagraphTabStops;
use crate::text::style::{
    ParagraphIndentPoints, ParagraphIndents, ParagraphLineSpacing, ParagraphSpacing,
    ParagraphSpacingPoints, TextAlignment,
};
use crate::{Error, IWorkPackage, Result};

use super::{line_spacing_from_archive, locate_style, tabs};

const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum InheritanceControl {
    Continue,
    Complete,
}

pub(super) fn alignment(package: &IWorkPackage, first_style_id: u64) -> Result<TextAlignment> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(alignment) = style
            .para_properties
            .as_ref()
            .and_then(|properties| properties.alignment)
        else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(TextAlignment::from_native_value(alignment)?);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or(TextAlignment::Natural))
}

pub(super) fn line_spacing(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphLineSpacing> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(properties) = style.para_properties.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        if properties.line_spacing_null == Some(true) {
            *value = Some(ParagraphLineSpacing::default());
            return Ok(InheritanceControl::Complete);
        }
        let Some(spacing) = properties.line_spacing.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(line_spacing_from_archive(spacing)?);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

pub(super) fn spacing(package: &IWorkPackage, first_style_id: u64) -> Result<ParagraphSpacing> {
    let (before, after) = walk(
        package,
        first_style_id,
        (None, None),
        |(before, after), style| {
            if let Some(properties) = style.para_properties.as_ref() {
                if before.is_none() {
                    *before = properties
                        .space_before
                        .map(ParagraphSpacingPoints::from_points)
                        .transpose()?;
                }
                if after.is_none() {
                    *after = properties
                        .space_after
                        .map(ParagraphSpacingPoints::from_points)
                        .transpose()?;
                }
            }
            Ok(if before.is_some() && after.is_some() {
                InheritanceControl::Complete
            } else {
                InheritanceControl::Continue
            })
        },
    )?;
    Ok(ParagraphSpacing::new(
        before.unwrap_or(ParagraphSpacingPoints::ZERO),
        after.unwrap_or(ParagraphSpacingPoints::ZERO),
    ))
}

pub(super) fn indents(package: &IWorkPackage, first_style_id: u64) -> Result<ParagraphIndents> {
    let (first_line, left, right) = walk(
        package,
        first_style_id,
        (None, None, None),
        |(first_line, left, right), style| {
            if let Some(properties) = style.para_properties.as_ref() {
                if first_line.is_none() {
                    *first_line = properties
                        .first_line_indent
                        .map(ParagraphIndentPoints::from_points)
                        .transpose()?;
                }
                if left.is_none() {
                    *left = properties
                        .left_indent
                        .map(ParagraphIndentPoints::from_points)
                        .transpose()?;
                }
                if right.is_none() {
                    *right = properties
                        .right_indent
                        .map(ParagraphIndentPoints::from_points)
                        .transpose()?;
                }
            }
            Ok(
                if first_line.is_some() && left.is_some() && right.is_some() {
                    InheritanceControl::Complete
                } else {
                    InheritanceControl::Continue
                },
            )
        },
    )?;
    Ok(ParagraphIndents::new(
        first_line.unwrap_or(ParagraphIndentPoints::ZERO),
        left.unwrap_or(ParagraphIndentPoints::ZERO),
        right.unwrap_or(ParagraphIndentPoints::ZERO),
    ))
}

pub(super) fn tab_stops(package: &IWorkPackage, first_style_id: u64) -> Result<ParagraphTabStops> {
    let value = walk(package, first_style_id, None, |value, style| {
        let Some(properties) = style.para_properties.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        if properties.tabs_null == Some(true) {
            if properties.tabs.is_some() {
                return Err(Error::InvalidFormat(
                    "native iWork paragraph tabs are both null and populated".to_owned(),
                ));
            }
            *value = Some(ParagraphTabStops::default());
            return Ok(InheritanceControl::Complete);
        }
        let Some(archive) = properties.tabs.as_ref() else {
            return Ok(InheritanceControl::Continue);
        };
        *value = Some(tabs::from_archive(archive)?);
        Ok(InheritanceControl::Complete)
    })?;
    Ok(value.unwrap_or_default())
}

fn walk<T, F>(package: &IWorkPackage, first_style_id: u64, mut state: T, mut visit: F) -> Result<T>
where
    F: FnMut(&mut T, &tswp::ParagraphStyleArchive) -> Result<InheritanceControl>,
{
    let mut visited = HashSet::new();
    let mut style_id = Some(first_style_id);
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = style_id else {
            return Ok(state);
        };
        if !visited.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style inheritance cycles at {identifier}"
            )));
        }
        let location = locate_style(package, identifier)?;
        if visit(&mut state, &location.style)? == InheritanceControl::Complete {
            return Ok(state);
        }
        style_id = location.style.super_.parent.map(|parent| parent.identifier);
    }
    Err(Error::InvalidFormat(format!(
        "iWork paragraph style inheritance exceeds {MAX_STYLE_INHERITANCE_DEPTH} levels"
    )))
}
