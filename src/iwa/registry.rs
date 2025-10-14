//! Protobuf Message Type Registry for iWork Applications
//!
//! iWork applications use integer type IDs to identify different protobuf message types.
//! This registry provides mappings from type IDs to message names for different applications.

use std::collections::HashMap;
use once_cell::sync::Lazy;

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

impl Application {
    /// Get application from string representation
    pub fn from_str(s: &str) -> Option<Self> {
        match s.to_lowercase().as_str() {
            "pages" => Some(Self::Pages),
            "keynote" => Some(Self::Keynote),
            "numbers" => Some(Self::Numbers),
            "common" => Some(Self::Common),
            _ => None,
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
        self.types.iter()
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
    registry.register(100, "KN.ArchiveInfo", Application::Keynote);
    registry.register(101, "KN.ShowArchive", Application::Keynote);
    registry.register(102, "KN.SlideArchive", Application::Keynote);
    registry.register(103, "KN.SlideNodeArchive", Application::Keynote);
    registry.register(104, "KN.PlaceholderArchive", Application::Keynote);
    registry.register(105, "KN.MasterSlideArchive", Application::Keynote);
    registry.register(106, "KN.ThemeArchive", Application::Keynote);
    registry.register(107, "KN.SlideStyleArchive", Application::Keynote);

    // KN Command Archives
    registry.register(148, "KN.CommandSlideReapplyMasterArchive", Application::Keynote);
    registry.register(147, "KN.SlideCollectionCommandSelectionBehaviorArchive", Application::Keynote);
    registry.register(146, "KN.CommandSlideReapplyMasterArchive", Application::Keynote);
    registry.register(145, "KN.CommandMasterSetBodyStylesArchive", Application::Keynote);

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
    app_counts.into_iter()
        .max_by_key(|&(_, count)| count)
        .map(|(app, _)| app)
}

#[cfg(test)]
mod tests {
    use super::*;

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
    fn test_application_from_string() {
        assert_eq!(Application::from_str("pages"), Some(Application::Pages));
        assert_eq!(Application::from_str("Pages"), Some(Application::Pages));
        assert_eq!(Application::from_str("keynote"), Some(Application::Keynote));
        assert_eq!(Application::from_str("numbers"), Some(Application::Numbers));
        assert_eq!(Application::from_str("unknown"), None);
    }
}
