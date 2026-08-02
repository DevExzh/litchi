//! Protobuf Message Type Registry for iWork Applications
//!
//! iWork applications use integer type IDs to identify different protobuf message types.
//! This registry provides mappings from type IDs to message names for different applications.

use once_cell::sync::Lazy;
use std::{collections::HashMap, str::FromStr};

/// Application type for iWork documents
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Application {
    /// Apple Pages
    Pages,
    /// Apple Keynote
    Keynote,
    /// Apple Numbers
    Numbers,
    /// Common/shared types
    Common,
}

impl FromStr for Application {
    type Err = &'static str;
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s.to_lowercase().as_str() {
            "pages" => Ok(Self::Pages),
            "keynote" => Ok(Self::Keynote),
            "numbers" => Ok(Self::Numbers),
            "common" => Ok(Self::Common),
            _ => Err("Invalid input"),
        }
    }
}

/// Message type information
#[derive(Debug, Clone)]
pub struct MessageType {
    /// Human-readable name of the message type
    pub name: &'static str,
    /// Application this type belongs to
    pub application: Application,
}

/// Global registry of message types
pub struct MessageRegistry {
    types: HashMap<u32, MessageType>,
}

impl Default for MessageRegistry {
    fn default() -> Self {
        Self::new()
    }
}

impl MessageRegistry {
    /// Create a new empty registry
    pub fn new() -> Self {
        Self {
            types: HashMap::new(),
        }
    }

    /// Register a message type
    pub fn register(&mut self, id: u32, name: &'static str, application: Application) {
        self.types.insert(id, MessageType { name, application });
    }

    /// Look up a message type by ID
    pub fn lookup(&self, id: u32) -> Option<&MessageType> {
        self.types.get(&id)
    }

    /// Get all message types for a specific application
    pub fn types_for_application(&self, app: Application) -> Vec<(u32, &MessageType)> {
        self.types
            .iter()
            .filter(|(_, mt)| mt.application == app)
            .map(|(id, mt)| (*id, mt))
            .collect()
    }
}

/// Global message type registry instance
pub static MESSAGE_REGISTRY: Lazy<MessageRegistry> = Lazy::new(|| {
    let mut registry = MessageRegistry::new();

    // Common/Shared Types (TSP - Telesphoreo?)
    register_common_types(&mut registry);

    // Keynote Types (KN)
    register_keynote_types(&mut registry);

    // Numbers Types (TN)
    register_numbers_types(&mut registry);

    // Pages Types (TP)
    register_pages_types(&mut registry);

    // Additional shared types
    register_shared_types(&mut registry);

    registry
});

/// Register common message types used across applications
fn register_common_types(registry: &mut MessageRegistry) {
    // TSP (Telesphoreo?) common types
    registry.register(1, "TSP.ArchiveInfo", Application::Common);
    registry.register(2, "TSP.MessageInfo", Application::Common);
    registry.register(10, "TSP.DatabaseData", Application::Common);
    registry.register(100, "TSP.DocumentMetadata", Application::Common);
    registry.register(110, "TSP.ObjectReference", Application::Common);
    registry.register(200, "TSP.DataReference", Application::Common);
}

/// Register Keynote-specific message types
fn register_keynote_types(registry: &mut MessageRegistry) {
    // KN Archives (Keynote Archives)
    registry.register(8, "KN.BuildArchive", Application::Keynote);
    registry.register(153, "KN.BuildChunkArchive", Application::Keynote);
    registry.register(100, "KN.ArchiveInfo", Application::Keynote);
    registry.register(101, "KN.ShowArchive", Application::Keynote);
    registry.register(102, "KN.SlideArchive", Application::Keynote);
    registry.register(103, "KN.SlideNodeArchive", Application::Keynote);
    registry.register(104, "KN.PlaceholderArchive", Application::Keynote);
    registry.register(105, "KN.MasterSlideArchive", Application::Keynote);
    registry.register(106, "KN.ThemeArchive", Application::Keynote);
    registry.register(107, "KN.SlideStyleArchive", Application::Keynote);

    // KN Command Archives
    registry.register(
        148,
        "KN.CommandSlideReapplyMasterArchive",
        Application::Keynote,
    );
    registry.register(
        147,
        "KN.SlideCollectionCommandSelectionBehaviorArchive",
        Application::Keynote,
    );
    registry.register(
        146,
        "KN.CommandSlideReapplyMasterArchive",
        Application::Keynote,
    );
    registry.register(
        145,
        "KN.CommandMasterSetBodyStylesArchive",
        Application::Keynote,
    );

    // Additional Keynote types
    registry.register(200, "KN.PresentationArchive", Application::Keynote);
    registry.register(201, "KN.SlideTreeArchive", Application::Keynote);
    registry.register(202, "KN.BuildArchive", Application::Keynote);
    registry.register(203, "KN.TransitionArchive", Application::Keynote);
}

/// Register Numbers-specific message types
fn register_numbers_types(registry: &mut MessageRegistry) {
    // TN Archives (Numbers Archives)
    registry.register(1, "TN.SheetArchive", Application::Numbers);
    registry.register(2, "TN.TableArchive", Application::Numbers);
    registry.register(3, "TN.CellArchive", Application::Numbers);
    registry.register(4, "TN.FormulaArchive", Application::Numbers);
    registry.register(5, "TN.ChartArchive", Application::Numbers);
    registry.register(6, "TN.DocumentArchive", Application::Numbers);
    registry.register(7, "TN.WorkbookArchive", Application::Numbers);

    // TN Command Archives
    registry.register(100, "TN.CommandSetTableDataArchive", Application::Numbers);
    registry.register(101, "TN.CommandSetCellValueArchive", Application::Numbers);
    registry.register(102, "TN.CommandAddTableArchive", Application::Numbers);
    registry.register(103, "TN.CommandRemoveTableArchive", Application::Numbers);
}

/// Register Pages-specific message types
fn register_pages_types(registry: &mut MessageRegistry) {
    // TP Archives (Pages Archives)
    registry.register(1, "TP.DocumentArchive", Application::Pages);
    registry.register(2, "TP.SectionArchive", Application::Pages);
    registry.register(3, "TP.PageArchive", Application::Pages);
    registry.register(4, "TP.TextArchive", Application::Pages);
    registry.register(5, "TP.ParagraphArchive", Application::Pages);
    registry.register(6, "TP.CharacterArchive", Application::Pages);
    registry.register(7, "TP.ImageArchive", Application::Pages);

    // TP Command Archives
    registry.register(100, "TP.CommandSetTextArchive", Application::Pages);
    registry.register(101, "TP.CommandInsertTextArchive", Application::Pages);
    registry.register(102, "TP.CommandDeleteTextArchive", Application::Pages);
    registry.register(103, "TP.CommandSetStyleArchive", Application::Pages);
}

/// Register additional shared message types
fn register_shared_types(registry: &mut MessageRegistry) {
    // TSA (Text Style Archives?)
    registry.register(1, "TSA.StyleArchive", Application::Common);
    registry.register(2, "TSA.ParagraphStyleArchive", Application::Common);
    registry.register(3, "TSA.CharacterStyleArchive", Application::Common);
    registry.register(4, "TSA.ListStyleArchive", Application::Common);

    // TSD (Drawing?)
    registry.register(1, "TSD.DrawingArchive", Application::Common);
    registry.register(2, "TSD.ShapeArchive", Application::Common);
    registry.register(3, "TSD.ImageArchive", Application::Common);
    registry.register(4, "TSD.GroupArchive", Application::Common);

    // TSCH (Charts)
    registry.register(1, "TSCH.ChartArchive", Application::Common);
    registry.register(2, "TSCH.ChartSeriesArchive", Application::Common);
    registry.register(3, "TSCH.ChartAxisArchive", Application::Common);
    registry.register(4, "TSCH.ChartLegendArchive", Application::Common);

    // TSK (Task?)
    registry.register(1, "TSK.DocumentArchive", Application::Common);
    registry.register(2, "TSK.TaskArchive", Application::Common);

    // TSS (Style Sheet?)
    registry.register(1, "TSS.StyleSheetArchive", Application::Common);
    registry.register(2, "TSS.StylesArchive", Application::Common);

    // TST (Table?)
    registry.register(1, "TST.TableArchive", Application::Common);
    registry.register(2, "TST.TableCellArchive", Application::Common);
    registry.register(3, "TST.TableRowArchive", Application::Common);
    registry.register(4, "TST.TableColumnArchive", Application::Common);

    // TSWP (Word Processing?)
    registry.register(1, "TSWP.DocumentArchive", Application::Pages);
    registry.register(2, "TSWP.SectionArchive", Application::Pages);
    registry.register(3, "TSWP.ParagraphArchive", Application::Pages);
    registry.register(4, "TSWP.CharacterArchive", Application::Pages);
    registry.register(5, "TSWP.TextArchive", Application::Pages);
}

/// Get message type information by ID
pub fn get_message_type(id: u32) -> Option<&'static MessageType> {
    MESSAGE_REGISTRY.lookup(id)
}

/// Get all message types for a specific application
pub fn get_message_types_for_app(app: Application) -> Vec<(u32, &'static MessageType)> {
    MESSAGE_REGISTRY.types_for_application(app)
}

/// Attempt to determine application type from a collection of message types
pub fn detect_application(message_type_ids: &[u32]) -> Option<Application> {
    let mut app_counts = std::collections::HashMap::new();

    for &id in message_type_ids {
        if let Some(msg_type) = get_message_type(id) {
            *app_counts.entry(msg_type.application).or_insert(0) += 1;
        }
    }

    // Return the application with the most message types
    app_counts
        .into_iter()
        .max_by_key(|&(_, count)| count)
        .map(|(app, _)| app)
}

/// Detect the owning iWork application from the root `DocumentArchive` payload.
///
/// Message type identifiers overlap between Pages, Numbers, and Keynote, so they
/// cannot reliably identify an application. The root protobuf schemas have
/// stable, application-specific required message shapes: Pages uses its shared
/// document at field 15, Numbers uses references at fields 4/5/6 plus its shared
/// document at field 8, and Keynote uses a reference at field 2 plus its shared
/// document at field 3. Malformed or multiply matching payloads fail closed.
pub fn detect_application_from_document(payload: &[u8]) -> Option<Application> {
    let fields = wire_fields(payload)?;
    let pages = unique_field(&fields, 15, 2).is_some_and(valid_shared_document);
    let numbers = [4, 5, 6]
        .into_iter()
        .all(|field| unique_field(&fields, field, 2).is_some_and(valid_reference))
        && unique_field(&fields, 8, 2).is_some_and(valid_shared_document);
    let keynote = unique_field(&fields, 2, 2).is_some_and(valid_reference)
        && unique_field(&fields, 3, 2).is_some_and(valid_shared_document);

    match (pages, numbers, keynote) {
        (true, false, false) => Some(Application::Pages),
        (false, true, false) => Some(Application::Numbers),
        (false, false, true) => Some(Application::Keynote),
        _ => None,
    }
}

#[derive(Debug, Clone, Copy)]
struct WireField<'a> {
    number: u32,
    wire_type: u8,
    value: &'a [u8],
}

fn wire_fields(payload: &[u8]) -> Option<Vec<WireField<'_>>> {
    let mut fields = Vec::new();
    let mut position = 0;

    while position < payload.len() {
        let tag = read_varint(payload, &mut position)?;
        let number = u32::try_from(tag >> 3).ok()?;
        let wire_type = (tag & 0x07) as u8;
        if number == 0 {
            return None;
        }

        let value = match wire_type {
            0 => {
                let start = position;
                read_varint(payload, &mut position)?;
                payload.get(start..position)?
            },
            1 => take(payload, &mut position, 8)?,
            2 => {
                let length = usize::try_from(read_varint(payload, &mut position)?).ok()?;
                take(payload, &mut position, length)?
            },
            5 => take(payload, &mut position, 4)?,
            _ => return None,
        };

        fields.push(WireField {
            number,
            wire_type,
            value,
        });
    }

    Some(fields)
}

fn take<'a>(payload: &'a [u8], position: &mut usize, length: usize) -> Option<&'a [u8]> {
    let end = position.checked_add(length)?;
    let value = payload.get(*position..end)?;
    *position = end;
    Some(value)
}

fn unique_field<'a>(fields: &[WireField<'a>], number: u32, wire_type: u8) -> Option<&'a [u8]> {
    let mut matches = fields.iter().filter(|field| field.number == number);
    let field = matches.next()?;
    if matches.next().is_some() || field.wire_type != wire_type {
        return None;
    }
    Some(field.value)
}

fn valid_reference(payload: &[u8]) -> bool {
    wire_fields(payload)
        .and_then(|fields| unique_field(&fields, 1, 0))
        .is_some()
}

fn valid_shared_document(payload: &[u8]) -> bool {
    wire_fields(payload)
        .and_then(|fields| unique_field(&fields, 1, 2))
        .and_then(wire_fields)
        .is_some()
}

fn read_varint(payload: &[u8], position: &mut usize) -> Option<u64> {
    let mut value = 0u64;
    for shift in (0..=63).step_by(7) {
        let byte = *payload.get(*position)?;
        *position += 1;
        if shift == 63 && byte > 1 {
            return None;
        }
        value |= u64::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            return Some(value);
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::{kn, tn, tp, tsa, tsk, tsp};
    use prost::Message;

    fn shared_document() -> tsa::DocumentArchive {
        tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        }
    }

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            ..Default::default()
        }
    }

    fn document_payload(application: Application) -> Vec<u8> {
        match application {
            Application::Pages => tp::DocumentArchive {
                super_: shared_document(),
                ..Default::default()
            }
            .encode_to_vec(),
            Application::Numbers => tn::DocumentArchive {
                super_: shared_document(),
                stylesheet: reference(1),
                sidebar_order: reference(2),
                theme: reference(3),
                ..Default::default()
            }
            .encode_to_vec(),
            Application::Keynote => kn::DocumentArchive {
                super_: shared_document(),
                show: reference(1),
                ..Default::default()
            }
            .encode_to_vec(),
            Application::Common => Vec::new(),
        }
    }

    #[test]
    fn test_message_type_lookup() {
        // Test that we can look up known message types
        let archive_info = get_message_type(1);
        assert!(archive_info.is_some());

        // Test Keynote types
        let kn_show = get_message_type(101);
        assert!(kn_show.is_some());

        // Test that we can look up message types (basic functionality)
        assert!(get_message_type(1).is_some());
        assert!(get_message_type(999).is_none()); // Non-existent type
    }

    #[test]
    fn test_application_detection() {
        // Test Keynote detection
        let keynote_ids = vec![101, 102, 103]; // KN.ShowArchive, KN.SlideArchive, etc.
        let keynote_result = detect_application(&keynote_ids);
        assert!(keynote_result.is_some()); // Should detect some application

        // Test with common types
        let common_ids = vec![1, 2, 3]; // Common types
        let common_result = detect_application(&common_ids);
        assert!(common_result.is_some()); // Should detect some application

        // Test empty input
        assert_eq!(detect_application(&[]), None);
    }

    #[test]
    fn test_document_payload_detection() {
        assert_eq!(
            detect_application_from_document(&document_payload(Application::Pages)),
            Some(Application::Pages)
        );
        assert_eq!(
            detect_application_from_document(&document_payload(Application::Numbers)),
            Some(Application::Numbers)
        );
        assert_eq!(
            detect_application_from_document(&document_payload(Application::Keynote)),
            Some(Application::Keynote)
        );

        let pages_with_references = tp::DocumentArchive {
            super_: shared_document(),
            stylesheet: Some(reference(1)),
            floating_drawables: Some(reference(2)),
            ..Default::default()
        }
        .encode_to_vec();
        assert_eq!(
            detect_application_from_document(&pages_with_references),
            Some(Application::Pages)
        );

        let mut conflicting = document_payload(Application::Pages);
        conflicting.extend(document_payload(Application::Numbers));
        assert_eq!(detect_application_from_document(&conflicting), None);

        let mut conflicting = document_payload(Application::Pages);
        conflicting.extend(document_payload(Application::Keynote));
        assert_eq!(detect_application_from_document(&conflicting), None);

        assert_eq!(detect_application_from_document(&[0x78, 0x00]), None);
        assert_eq!(detect_application_from_document(&[0x7a, 0x00]), None);
        assert_eq!(detect_application_from_document(&[0x80]), None);
    }

    #[test]
    fn test_application_from_string() {
        assert_eq!(Application::from_str("pages"), Ok(Application::Pages));
        assert_eq!(Application::from_str("Pages"), Ok(Application::Pages));
        assert_eq!(Application::from_str("keynote"), Ok(Application::Keynote));
        assert_eq!(Application::from_str("numbers"), Ok(Application::Numbers));
        assert_eq!(Application::from_str("unknown"), Err("Invalid input"));
    }
}
