//! Public semantic models for ODT-specific parsed structures.

/// Complete inert tracked-change declarations and their container policy.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TrackedChanges {
    /// Whether the producer requested change tracking to remain enabled.
    pub track_changes: Option<bool>,
    /// Stored base64 protection-key material. It is never used to unlock changes.
    pub protection_key: Option<String>,
    /// Digest algorithm URI associated with the protection key.
    pub protection_key_digest_algorithm: Option<String>,
    /// Change declarations in document order.
    pub changes: Vec<TrackChange>,
}

/// Represents a tracked change in the document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TrackChange {
    /// Change ID
    pub id: String,
    /// Optional XML identifier retained separately from `text:id`.
    pub xml_id: Option<String>,
    /// Author who made the change
    pub author: Option<String>,
    /// Date/time of the change
    pub date: Option<String>,
    /// Optional review comment stored in `office:change-info`.
    pub comment: Option<String>,
    /// Type of change (insertion, deletion, format-change)
    pub change_type: ChangeType,
    /// Style referenced by a format change. The style is not resolved automatically.
    pub style_name: Option<String>,
    /// Deletion paragraph-merge behavior when explicitly stored.
    pub merge_last_paragraph: Option<bool>,
    /// Changed text content
    pub content: String,
}

/// Type of tracked change
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChangeType {
    /// Text insertion
    Insertion,
    /// Text deletion
    Deletion,
    /// Formatting change
    FormatChange,
}

/// Represents a comment/annotation in the document
#[derive(Debug, Clone)]
pub struct Comment {
    /// Comment ID
    pub id: String,
    /// Author of the comment
    pub author: Option<String>,
    /// Date/time of the comment
    pub date: Option<String>,
    /// Comment text content
    pub content: String,
    /// Referenced text in the document
    pub reference: Option<String>,
}

/// Represents a section in the document
#[derive(Debug, Clone)]
pub struct Section {
    /// Section name
    pub name: String,
    /// Section style
    pub style: Option<String>,
    /// Whether the section is protected
    pub protected: bool,
    /// Optional XML identifier.
    pub xml_id: Option<String>,
    /// Stored protection-key material; never used to unlock content automatically.
    pub protection_key: Option<String>,
    /// Digest algorithm URI for the protection key.
    pub protection_key_digest_algorithm: Option<String>,
    /// Visibility behavior.
    pub display: SectionDisplay,
    /// Inert condition expression for conditionally displayed sections.
    pub condition: Option<String>,
    /// Optional linked-section source; never fetched.
    pub source: Option<SectionSource>,
    /// Optional DDE source; never activated.
    pub dde_source: Option<SectionDdeSource>,
    /// Text content within the section
    pub content: String,
}

/// Section visibility behavior.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionDisplay {
    /// Display normally.
    Visible,
    /// Do not display.
    Hidden,
    /// Display according to the stored inert condition.
    Condition,
}

/// An inert linked-section source declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionSource {
    /// External or package-local URI. It is never fetched automatically.
    pub href: Option<String>,
    /// Named section within the source document.
    pub section_name: Option<String>,
    /// Producer-specific import filter name.
    pub filter_name: Option<String>,
}

/// An inert Dynamic Data Exchange source declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionDdeSource {
    /// DDE source name.
    pub name: Option<String>,
    /// Stored conversion mode.
    pub conversion_mode: Option<String>,
    /// Whether the producer requested automatic updates; no update is performed.
    pub automatic_update: Option<bool>,
}
