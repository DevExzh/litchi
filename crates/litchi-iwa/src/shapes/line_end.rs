//! Typed native line-end geometry and copy-on-write shape-style updates.

mod native;
mod registry;

use std::collections::HashSet;

use prost::Message;

use crate::archive::RawMessage;
use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    next_object_identifier, set_package_last_object_identifier,
};
use crate::protobuf::{tsd, tsp, tswp};
use crate::{Error, IWorkPackage, Result};

use super::shape_line_segment;
#[cfg(test)]
use native::endpoint_archive;
use registry::{
    LineEndVariationLocation, collapse_line_end_variation, endpoint_from_archive,
    insert_style_variation, is_collapsible_line_end_variation, line_style_is_exclusive,
    line_style_object, patch_shape_style_reference, replace_style_variation,
};

const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;

/// A native endpoint decoration supported by Pages, Numbers, and Keynote.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum LineEndpoint {
    /// No endpoint decoration.
    #[default]
    None,
    /// A compact filled triangular arrowhead.
    SimpleArrow,
    /// A filled circular endpoint.
    FilledCircle,
    /// A filled diamond endpoint.
    FilledDiamond,
    /// An outlined arrowhead with a short center stem.
    OpenArrow,
    /// A broad filled arrowhead with an inset base.
    FilledArrow,
    /// A filled square endpoint.
    FilledSquare,
    /// An outlined square endpoint.
    OpenSquare,
    /// An outlined circular endpoint.
    OpenCircle,
    /// A filled arrowhead pointing toward the line segment.
    InvertedArrow,
    /// A perpendicular bar endpoint.
    Line,
}

/// Decorations at the directed start and end of a straight line.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct LineEndpoints {
    /// Decoration at the directed start point.
    pub start: LineEndpoint,
    /// Decoration at the directed end point.
    pub end: LineEndpoint,
}

impl LineEndpoints {
    /// Construct independently typed start and end decorations.
    pub const fn new(start: LineEndpoint, end: LineEndpoint) -> Self {
        Self { start, end }
    }
}

#[allow(deprecated)]
pub(crate) fn shape_line_endpoints(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
) -> Result<LineEndpoints> {
    let shape = shape_payload(package, archive_name, drawable_id)?;
    if shape_line_segment(&shape)?.is_none() {
        return Err(Error::ParseError(format!(
            "iWork drawable {drawable_id} is not a native straight line"
        )));
    }
    let mut head = shape.super_.head_line_end.clone();
    let mut tail = shape.super_.tail_line_end.clone();
    if head.is_none() || tail.is_none() {
        let style_id = shape
            .super_
            .style
            .as_ref()
            .map(|reference| reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(format!("iWork line {drawable_id} has no shape style"))
            })?;
        let inherited = inherited_line_ends(package, style_id)?;
        head = head.or(inherited.0);
        tail = tail.or(inherited.1);
    }
    let mut endpoints = LineEndpoints {
        start: endpoint_from_archive(tail.as_ref())?,
        end: endpoint_from_archive(head.as_ref())?,
    };
    if shape
        .super_
        .pathsource
        .as_ref()
        .is_some_and(|path| path.horizontal_flip == Some(true))
    {
        std::mem::swap(&mut endpoints.start, &mut endpoints.end);
    }
    Ok(endpoints)
}

#[allow(deprecated)]
pub(crate) fn set_shape_line_endpoints(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    endpoints: LineEndpoints,
) -> Result<()> {
    let shape = shape_payload(package, archive_name, drawable_id)?;
    if shape_line_segment(&shape)?.is_none() {
        return Err(Error::ParseError(format!(
            "iWork drawable {drawable_id} is not a native straight line"
        )));
    }
    if shape_line_endpoints(package, archive_name, drawable_id)? == endpoints {
        return Ok(());
    }
    let old_style_id = shape
        .super_
        .style
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork line {drawable_id} has no style")))?;
    let style_archive_name = object_archive_name(package, old_style_id)?;
    let old_style_message = shape_style_message(package, &style_archive_name, old_style_id)?;
    let old_style = tswp::ShapeStyleArchive::decode(old_style_message.data.as_slice())?;
    let stylesheet_id = old_style
        .super_
        .super_
        .stylesheet
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork shape style {old_style_id} has no stylesheet"
            ))
        })?;
    let stylesheet_archive_name = object_archive_name(package, stylesheet_id)?;
    if stylesheet_archive_name != style_archive_name {
        return Err(Error::InvalidFormat(format!(
            "iWork shape style {old_style_id} is not stored with stylesheet {stylesheet_id}"
        )));
    }
    let mut stored = endpoints;
    if shape
        .super_
        .pathsource
        .as_ref()
        .is_some_and(|path| path.horizontal_flip == Some(true))
    {
        std::mem::swap(&mut stored.start, &mut stored.end);
    }
    let disposable = is_collapsible_line_end_variation(&old_style, &old_style_message.data)?
        && line_style_is_exclusive(package, old_style_id)?;
    let parent_style_id = old_style
        .super_
        .super_
        .parent
        .as_ref()
        .map(|reference| reference.identifier);

    if disposable && endpoints == LineEndpoints::default() {
        let parent_style_id = parent_style_id.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork endpoint variation {old_style_id} has no parent style"
            ))
        })?;
        let inherited = inherited_line_ends(package, parent_style_id)?;
        let inherited = LineEndpoints {
            start: endpoint_from_archive(inherited.1.as_ref())?,
            end: endpoint_from_archive(inherited.0.as_ref())?,
        };
        if inherited == LineEndpoints::default() {
            return collapse_line_end_variation(
                package,
                LineEndVariationLocation {
                    drawable_archive_name: archive_name,
                    style_archive_name: &style_archive_name,
                    drawable_id,
                    stylesheet_id,
                    style_id: old_style_id,
                    parent_style_id,
                },
            );
        }
    }

    if disposable {
        let parent_style_id = parent_style_id.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork endpoint variation {old_style_id} has no parent style"
            ))
        })?;
        let replacement = line_style_object(old_style_id, parent_style_id, stylesheet_id, stored)?;
        let mut staged = package.clone();
        replace_style_variation(&mut staged, &style_archive_name, old_style_id, replacement)?;
        if shape_line_endpoints(&staged, archive_name, drawable_id)? != endpoints {
            return Err(Error::InvalidFormat(
                "iWork line endpoint-style update failed validation".to_owned(),
            ));
        }
        *package = staged;
        return Ok(());
    }

    let new_style_id = next_object_identifier(package)?;
    let new_style = line_style_object(new_style_id, old_style_id, stylesheet_id, stored)?;

    let mut staged = package.clone();
    patch_shape_style_reference(
        &mut staged,
        archive_name,
        drawable_id,
        old_style_id,
        new_style_id,
    )?;
    insert_style_variation(
        &mut staged,
        &style_archive_name,
        stylesheet_id,
        old_style_id,
        new_style_id,
        new_style,
    )?;
    if let Some(style_component) = component_identifier_for_entry(&staged, &style_archive_name)? {
        add_component_object_uuids(&mut staged, style_component, &[new_style_id])?;
        if let Some(drawable_component) = component_identifier_for_entry(&staged, archive_name)?
            && drawable_component != style_component
        {
            add_component_external_reference(
                &mut staged,
                drawable_component,
                style_component,
                new_style_id,
            )?;
        }
    }
    set_package_last_object_identifier(&mut staged, new_style_id)?;
    if shape_line_endpoints(&staged, archive_name, drawable_id)? != endpoints {
        return Err(Error::InvalidFormat(
            "iWork line endpoint-style update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

fn shape_payload(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
) -> Result<tswp::ShapeInfoArchive> {
    let archive = package.archive(archive_name)?;
    let object = archive.object(drawable_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork drawable object {drawable_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork drawable {drawable_id} must have exactly one ShapeInfo payload"
        )));
    };
    Ok(tswp::ShapeInfoArchive::decode(message.data.as_slice())?)
}

fn shape_style(
    package: &IWorkPackage,
    archive_name: &str,
    style_id: u64,
) -> Result<tswp::ShapeStyleArchive> {
    Ok(tswp::ShapeStyleArchive::decode(
        shape_style_message(package, archive_name, style_id)?
            .data
            .as_slice(),
    )?)
}

fn shape_style_message(
    package: &IWorkPackage,
    archive_name: &str,
    style_id: u64,
) -> Result<RawMessage> {
    let archive = package.archive(archive_name)?;
    let object = archive
        .object(style_id)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork shape style {style_id} is missing")))?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == SHAPE_STYLE_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork style {style_id} must have exactly one ShapeStyle payload"
        )));
    };
    Ok((*message).clone())
}

fn inherited_line_ends(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<(Option<tsd::LineEndArchive>, Option<tsd::LineEndArchive>)> {
    let mut visited = HashSet::new();
    let mut style_id = Some(first_style_id);
    let mut head = None;
    let mut tail = None;
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = style_id else {
            return Ok((head, tail));
        };
        if !visited.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork shape style inheritance cycles at {identifier}"
            )));
        }
        let archive_name = object_archive_name(package, identifier)?;
        let style = shape_style(package, &archive_name, identifier)?;
        if let Some(properties) = style.super_.shape_properties {
            head = head.or(properties.head_line_end);
            tail = tail.or(properties.tail_line_end);
        }
        if head.is_some() && tail.is_some() {
            return Ok((head, tail));
        }
        style_id = style
            .super_
            .super_
            .parent
            .map(|reference| reference.identifier);
    }
    Err(Error::InvalidFormat(format!(
        "iWork shape style inheritance exceeds {MAX_STYLE_INHERITANCE_DEPTH} levels"
    )))
}

fn object_archive_name(package: &IWorkPackage, identifier: u64) -> Result<String> {
    let mut found = None;
    for name in package.iwa_entry_names() {
        if package.archive(name)?.object(identifier).is_some()
            && found.replace(name.to_owned()).is_some()
        {
            return Err(Error::InvalidFormat(format!(
                "iWork object {identifier} occurs in multiple archives"
            )));
        }
    }
    found.ok_or_else(|| Error::InvalidFormat(format!("iWork object {identifier} is missing")))
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_typed_endpoint_round_trips_through_native_identifier() {
        for endpoint in [
            LineEndpoint::None,
            LineEndpoint::SimpleArrow,
            LineEndpoint::FilledCircle,
            LineEndpoint::FilledDiamond,
            LineEndpoint::OpenArrow,
            LineEndpoint::FilledArrow,
            LineEndpoint::FilledSquare,
            LineEndpoint::OpenSquare,
            LineEndpoint::OpenCircle,
            LineEndpoint::InvertedArrow,
            LineEndpoint::Line,
        ] {
            let archive = endpoint_archive(endpoint);
            assert_eq!(endpoint_from_archive(Some(&archive)).unwrap(), endpoint);
            assert!(archive.path.is_some());
        }
    }

    #[test]
    fn line_endpoints_default_to_undecorated() {
        assert_eq!(
            LineEndpoints::default(),
            LineEndpoints::new(LineEndpoint::None, LineEndpoint::None)
        );
    }
}
