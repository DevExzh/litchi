//! Structural, semantic, and resource-limit validation for diagram data.

use crate::{Error, Result};
use std::collections::{HashMap, HashSet};
use std::fmt;

use super::{
    Conformance, Connection, ConnectionType, DiagramDataModel, Id, MAX_CONNECTIONS,
    MAX_DATA_MODEL_XML, MAX_POINTS, MAX_TEXT_BYTES, Point, PointType,
};

impl DiagramDataModel {
    /// Validates identifiers, references, Office's single-parent rule, and
    /// configured resource limits without changing the model.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn validate(&self) -> Result<()> {
        self.serialized_xml_len(Conformance::Transitional)
            .map(|_| ())
    }

    fn validated_xml_capacity(&self) -> Result<usize> {
        if self.points.len() > MAX_POINTS {
            return Err(limit("diagram point count"));
        }
        if self.connections.len() > MAX_CONNECTIONS {
            return Err(limit("diagram connection count"));
        }
        let mut xml_capacity = 256usize;
        let mut model_ids = HashSet::with_capacity(self.points.len() + self.connections.len());
        let mut points = HashMap::with_capacity(self.points.len());
        let mut document_seen = false;
        for point in &self.points {
            if !model_ids.insert(point.id) {
                return Err(invalid(format!(
                    "duplicate diagram model identifier {}",
                    point.id
                )));
            }
            points.insert(point.id, point);
            if point.kind == PointType::Document {
                if document_seen {
                    return Err(invalid("diagram data model has multiple document points"));
                }
                document_seen = true;
            }
            if point.text.len() > MAX_TEXT_BYTES {
                return Err(limit("diagram text bytes"));
            }
            let has_modeled_children = !point.text.is_empty()
                || point.layout_type_id.is_some()
                || point.quick_style_type_id.is_some()
                || point.color_style_type_id.is_some();
            xml_capacity = xml_capacity
                .checked_add(if has_modeled_children { 320 } else { 128 })
                .ok_or_else(|| limit("serialized data-model bytes"))?;
            xml_capacity = add_xml_value(xml_capacity, &point.text, "diagram point text")?;
            for (value, description) in [
                (
                    point.layout_type_id.as_deref(),
                    "diagram layout type identifier",
                ),
                (
                    point.quick_style_type_id.as_deref(),
                    "diagram quick-style type identifier",
                ),
                (
                    point.color_style_type_id.as_deref(),
                    "diagram color-style type identifier",
                ),
            ] {
                if let Some(value) = value {
                    xml_capacity = add_xml_value(xml_capacity, value, description)?;
                }
            }
        }

        let mut connections = HashMap::with_capacity(self.connections.len());
        let mut parent_destinations = HashSet::new();
        let mut presentation_id: Option<&str> = None;
        for connection in &self.connections {
            xml_capacity = xml_capacity
                .checked_add(512)
                .ok_or_else(|| limit("serialized data-model bytes"))?;
            if !model_ids.insert(connection.id) {
                return Err(invalid(format!(
                    "duplicate diagram model identifier {}",
                    connection.id
                )));
            }
            connections.insert(connection.id, connection);
            if !points.contains_key(&connection.source) {
                return Err(invalid(format!(
                    "diagram connection {} has no source point {}",
                    connection.id, connection.source
                )));
            }
            if !points.contains_key(&connection.destination) {
                return Err(invalid(format!(
                    "diagram connection {} has no destination point {}",
                    connection.id, connection.destination
                )));
            }
            if connection.kind.is_parent() && !parent_destinations.insert(connection.destination) {
                return Err(invalid(format!(
                    "diagram point {} has multiple parents",
                    connection.destination
                )));
            }
            if let ConnectionType::Presentation(presentation) = &connection.kind {
                if presentation_id.is_some_and(|expected| expected != presentation) {
                    return Err(invalid(
                        "diagram presentation connections use different presentation identifiers",
                    ));
                }
                presentation_id.get_or_insert(presentation);
                xml_capacity = add_xml_value(
                    xml_capacity,
                    presentation,
                    "diagram presentation identifier",
                )?;
            }
            validate_connection_transitions(connection, |id| points.get(&id).copied())?;
        }
        for point in &self.points {
            if let Some(connection_id) = point.kind.connection() {
                let connection = connections.get(&connection_id).copied().ok_or_else(|| {
                    invalid(format!(
                        "diagram transition point {} refers to missing connection {}",
                        point.id, connection_id
                    ))
                })?;
                validate_transition_point(point, connection)?;
            }
        }
        Ok(xml_capacity.min(MAX_DATA_MODEL_XML))
    }
    pub(super) fn serialized_xml_len(&self, conformance: Conformance) -> Result<usize> {
        self.validated_xml_capacity()?;
        let mut count = XmlByteCount::default();
        self.write_validated_xml(&mut count, conformance)?;
        if count.0 > MAX_DATA_MODEL_XML {
            return Err(limit("serialized data-model bytes"));
        }
        Ok(count.0)
    }
}
pub(super) fn validate_connection_transitions<'a>(
    connection: &Connection,
    mut point: impl FnMut(Id) -> Option<&'a Point>,
) -> Result<()> {
    let ConnectionType::Parent {
        parent_transition,
        sibling_transition,
    } = &connection.kind
    else {
        return Ok(());
    };
    let parent = point(*parent_transition).ok_or_else(|| {
        invalid(format!(
            "diagram parent connection {} refers to missing parent transition point {}",
            connection.id, parent_transition
        ))
    })?;
    if parent.kind != PointType::ParentTransition(connection.id) {
        return Err(invalid(format!(
            "diagram point {} is not the parent transition for connection {}",
            parent.id, connection.id
        )));
    }
    let sibling = point(*sibling_transition).ok_or_else(|| {
        invalid(format!(
            "diagram parent connection {} refers to missing sibling transition point {}",
            connection.id, sibling_transition
        ))
    })?;
    if sibling.kind != PointType::SiblingTransition(connection.id) {
        return Err(invalid(format!(
            "diagram point {} is not the sibling transition for connection {}",
            sibling.id, connection.id
        )));
    }
    Ok(())
}

pub(super) fn validate_transition_point(point: &Point, connection: &Connection) -> Result<()> {
    let matches = match point.kind {
        PointType::ParentTransition(owner) if owner == connection.id => matches!(
            &connection.kind,
            ConnectionType::Parent {
                parent_transition,
                ..
            } if *parent_transition == point.id
        ),
        PointType::SiblingTransition(owner) if owner == connection.id => matches!(
            &connection.kind,
            ConnectionType::Parent {
                sibling_transition,
                ..
            } if *sibling_transition == point.id
        ),
        PointType::Node
        | PointType::Document
        | PointType::Assistant
        | PointType::ParentTransition(_)
        | PointType::SiblingTransition(_)
        | PointType::Presentation => false,
    };
    if matches {
        Ok(())
    } else {
        Err(invalid(format!(
            "diagram transition point {} and connection {} do not refer to each other",
            point.id, connection.id
        )))
    }
}

#[derive(Default)]
pub(super) struct XmlByteCount(pub(super) usize);

impl fmt::Write for XmlByteCount {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        self.0 = self.0.saturating_add(value.len());
        Ok(())
    }
}
fn add_xml_value(total: usize, value: &str, description: &str) -> Result<usize> {
    let mut escaped = 0usize;
    for character in value.chars() {
        if !is_xml_character(character) {
            return Err(invalid(format!(
                "{description} contains a character forbidden by XML 1.0"
            )));
        }
        let bytes = match character {
            '&' => 5,
            '<' | '>' => 4,
            '\'' | '"' => 6,
            character => character.len_utf8(),
        };
        escaped = escaped
            .checked_add(bytes)
            .ok_or_else(|| limit("serialized data-model bytes"))?;
    }
    total
        .checked_add(escaped)
        .ok_or_else(|| limit("serialized data-model bytes"))
}

#[inline]
const fn is_xml_character(character: char) -> bool {
    matches!(
        character as u32,
        0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
    )
}
pub(super) fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn limit(label: &str) -> Error {
    invalid(format!("diagram {label} limit exceeded"))
}
