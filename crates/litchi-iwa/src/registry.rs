//! Protobuf Message Type Registry for iWork Applications
//!
//! iWork applications use integer type IDs to identify different protobuf message types.
//! This registry provides mappings from type IDs to message names for different applications.

use crate::application::Application;
use once_cell::sync::Lazy;
use std::{collections::HashMap, fmt};

/// Error returned when a numeric message ID has more than one definition in
/// the requested registry scope.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MessageLookupError {
    /// Multiple definitions match the ID, so selecting one would be unsafe.
    Ambiguous {
        /// The numeric protobuf message ID.
        id: u32,
        /// `None` for the complete registry, or the application scope used.
        application: Option<Application>,
    },
}

impl fmt::Display for MessageLookupError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Ambiguous {
                id,
                application: Some(application),
            } => write!(
                formatter,
                "message type ID {id} is ambiguous for {application}"
            ),
            Self::Ambiguous {
                id,
                application: None,
            } => write!(formatter, "message type ID {id} is ambiguous"),
        }
    }
}

impl std::error::Error for MessageLookupError {}

impl MessageLookupError {
    fn ambiguous(id: u32, application: Option<Application>) -> Self {
        Self::Ambiguous { id, application }
    }
}

/// Message type information
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MessageType {
    /// Human-readable name of the message type
    pub name: &'static str,
    /// Application this type belongs to
    pub application: Application,
}

/// Global registry of message types
pub struct MessageRegistry {
    types: HashMap<u32, Vec<MessageType>>,
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
        self.types
            .entry(id)
            .or_default()
            .push(MessageType { name, application });
    }

    /// Look up a message type when the numeric ID is unambiguous.
    ///
    /// iWork reuses numeric message IDs across application namespaces. An
    /// ambiguous ID therefore has no safe single result and returns `None`.
    pub fn lookup(&self, id: u32) -> Option<&MessageType> {
        self.types
            .get(&id)
            .and_then(|types| (types.len() == 1).then_some(&types[0]))
    }

    /// Look up a message type and distinguish absence from ambiguity.
    pub fn lookup_unique(&self, id: u32) -> Result<Option<&MessageType>, MessageLookupError> {
        let mut definitions = self.lookup_all(id).iter();
        let Some(definition) = definitions.next() else {
            return Ok(None);
        };
        if definitions.next().is_some() {
            return Err(MessageLookupError::ambiguous(id, None));
        }
        Ok(Some(definition))
    }

    /// Return every registered definition for a numeric message ID.
    pub fn lookup_all(&self, id: u32) -> &[MessageType] {
        self.types.get(&id).map_or(&[], Vec::as_slice)
    }

    /// Look up the unique definition for an application-scoped message ID.
    ///
    /// A numeric ID may be shared by applications, while an application can
    /// still have only one definition for that ID. Missing definitions return
    /// `Ok(None)` and multiple definitions return a typed ambiguity error.
    pub fn lookup_for_application(
        &self,
        id: u32,
        application: Application,
    ) -> Result<Option<&MessageType>, MessageLookupError> {
        let mut definitions = self
            .lookup_all(id)
            .iter()
            .filter(|definition| definition.application == application);
        let Some(definition) = definitions.next() else {
            return Ok(None);
        };
        if definitions.next().is_some() {
            return Err(MessageLookupError::ambiguous(id, Some(application)));
        }
        Ok(Some(definition))
    }

    /// Get all message types for a specific application
    pub fn types_for_application(&self, app: Application) -> Vec<(u32, &MessageType)> {
        let mut types = self
            .types
            .iter()
            .flat_map(|(id, definitions)| {
                definitions
                    .iter()
                    .filter(move |definition| definition.application == app)
                    .map(|definition| (*id, definition))
            })
            .collect::<Vec<_>>();
        types.sort_unstable_by(|(left_id, left), (right_id, right)| {
            left_id
                .cmp(right_id)
                .then_with(|| left.name.cmp(right.name))
        });
        types
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

/// Strictly look up one message definition by numeric ID.
pub fn get_message_type_unique(
    id: u32,
) -> Result<Option<&'static MessageType>, MessageLookupError> {
    MESSAGE_REGISTRY.lookup_unique(id)
}

/// Get every registered message definition for a numeric ID.
pub fn get_message_types(id: u32) -> &'static [MessageType] {
    MESSAGE_REGISTRY.lookup_all(id)
}

/// Strictly look up one message definition within an application namespace.
pub fn get_message_type_for_app(
    id: u32,
    application: Application,
) -> Result<Option<&'static MessageType>, MessageLookupError> {
    MESSAGE_REGISTRY.lookup_for_application(id, application)
}

/// Get all message types for a specific application
pub fn get_message_types_for_app(app: Application) -> Vec<(u32, &'static MessageType)> {
    MESSAGE_REGISTRY.types_for_application(app)
}

/// Attempt to determine an application from numeric message IDs.
///
/// This is intentionally conservative. Shared definitions are ignored, each
/// numeric ID contributes at most once per application, and a tie is treated as
/// ambiguity. Callers that need document ownership must use validated root
/// `DocumentArchive` evidence instead.
pub fn detect_application(message_type_ids: &[u32]) -> Option<Application> {
    let applications = [
        Application::Pages,
        Application::Keynote,
        Application::Numbers,
    ];
    let mut app_counts = [0usize; 3];

    for &id in message_type_ids {
        let definitions = get_message_types(id);
        for (index, application) in applications.into_iter().enumerate() {
            if definitions
                .iter()
                .any(|definition| definition.application == application)
            {
                app_counts[index] += 1;
            }
        }
    }

    let maximum = app_counts.into_iter().max().unwrap_or(0);
    if maximum == 0 || app_counts.iter().filter(|count| **count == maximum).count() != 1 {
        return None;
    }
    app_counts
        .iter()
        .position(|count| *count == maximum)
        .map(|index| applications[index])
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::application::ParseError;
    use std::str::FromStr;

    #[test]
    fn test_message_type_lookup() {
        // Numeric IDs are intentionally shared by multiple application
        // namespaces, so a single-definition lookup must fail closed.
        assert!(get_message_type(1).is_none());
        assert!(get_message_type(101).is_none());
        assert!(get_message_types(1).len() > 1);
        assert!(get_message_types(101).len() > 1);
        assert!(matches!(
            get_message_type_unique(1),
            Err(MessageLookupError::Ambiguous {
                id: 1,
                application: None
            })
        ));
        assert_eq!(
            get_message_type(8).map(|message| message.application),
            Some(Application::Keynote)
        );
        assert_eq!(
            get_message_type_unique(8)
                .unwrap()
                .map(|message| message.name),
            Some("KN.BuildArchive")
        );
        assert_eq!(
            get_message_type_for_app(7, Application::Pages)
                .unwrap()
                .map(|message| message.name),
            Some("TP.ImageArchive")
        );
        assert!(
            get_message_type_for_app(999, Application::Pages)
                .unwrap()
                .is_none()
        );
        assert!(get_message_type(999).is_none()); // Non-existent type

        let pages = get_message_types_for_app(Application::Pages);
        assert!(
            pages
                .iter()
                .any(|(_, message)| message.name == "TP.DocumentArchive")
        );
        let keynote = get_message_types_for_app(Application::Keynote);
        assert!(
            keynote
                .iter()
                .any(|(_, message)| message.name == "KN.ShowArchive")
        );
    }

    #[test]
    fn registry_retains_colliding_definitions_without_overwriting() {
        let mut registry = MessageRegistry::new();
        registry.register(7, "TP.DocumentArchive", Application::Pages);
        registry.register(7, "TN.DocumentArchive", Application::Numbers);

        assert!(registry.lookup(7).is_none());
        assert!(matches!(
            registry.lookup_unique(7),
            Err(MessageLookupError::Ambiguous {
                id: 7,
                application: None
            })
        ));
        assert_eq!(
            registry
                .lookup_for_application(7, Application::Pages)
                .unwrap()
                .map(|message| message.name),
            Some("TP.DocumentArchive")
        );
        assert_eq!(
            registry.lookup_all(7),
            &[
                MessageType {
                    name: "TP.DocumentArchive",
                    application: Application::Pages,
                },
                MessageType {
                    name: "TN.DocumentArchive",
                    application: Application::Numbers,
                },
            ]
        );
    }

    #[test]
    fn test_application_detection() {
        // These IDs are shared by every concrete application namespace and
        // must not be used as an ownership guess.
        assert_eq!(detect_application(&[101, 102, 103]), None);
        assert_eq!(detect_application(&[1, 2, 3]), None);

        // A set containing only Keynote-specific IDs remains inferable.
        assert_eq!(
            detect_application(&[145, 146, 147, 148]),
            Some(Application::Keynote)
        );

        // Test empty input
        assert_eq!(detect_application(&[]), None);
    }

    #[test]
    fn test_application_from_string() {
        assert_eq!(Application::from_str("pages"), Ok(Application::Pages));
        assert_eq!(Application::from_str("Pages"), Ok(Application::Pages));
        assert_eq!(Application::from_str("keynote"), Ok(Application::Keynote));
        assert_eq!(Application::from_str("numbers"), Ok(Application::Numbers));
        assert_eq!(Application::from_str("unknown"), Err(ParseError));
        assert_eq!(Application::Pages.as_str(), "pages");
        assert_eq!(Application::Pages.to_string(), "pages");
        assert!(Application::Pages.is_concrete());
        assert!(!Application::Common.is_concrete());
    }
}
