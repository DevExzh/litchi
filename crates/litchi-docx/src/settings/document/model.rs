//! Typed document-settings model and ergonomic facade.

use super::super::model::Settings;
use super::super::notes::NoteNumberingProperties;
use crate::mail_merge::Settings as MailMergeSettings;
use crate::numbering::Format;
use crate::settings::{
    ColorSchemeMapping, CompatibilityOption, CompatibilitySetting, ProofingState, ProtectionType,
    SmartTagType, ThemeFontLanguages, View,
};

/// An inert reference to the external template associated with a document.
///
/// The target is never opened, fetched, normalized, or executed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AttachedTemplate {
    pub(super) relationship_id: String,
    pub(super) target_uri: String,
}

impl AttachedTemplate {
    /// Relationship ID used by `w:attachedTemplate` in `settings.xml`.
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// External relationship target exactly as stored in the package.
    pub fn target_uri(&self) -> &str {
        &self.target_uri
    }
}

/// Document settings including protection status.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// if let Some(settings) = doc.settings()? {
///     if settings.is_protected() {
///         println!("Document is protected");
///         println!("Protection type: {:?}", settings.protection_type());
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct DocumentSettings {
    /// Format-owned scalar settings parsed from `settings.xml`.
    pub(super) values: OwnedSettings,
    /// Smart-tag type declarations.
    pub(super) smart_tag_types: Vec<SmartTagType>,
    /// Whether applications should omit embedded smart-tag data when saving.
    pub(super) do_not_embed_smart_tags: bool,
    /// Inert mail-merge connection and display metadata.
    pub(super) mail_merge: Option<MailMergeSettings>,
    /// Inert external attached-template reference.
    pub(super) attached_template: Option<AttachedTemplate>,
}

type OwnedSettings = Settings<Format>;
impl DocumentSettings {
    /// Create a new Settings value with default values.
    pub fn new() -> Self {
        Self {
            values: OwnedSettings::new(),
            smart_tag_types: Vec::new(),
            do_not_embed_smart_tags: false,
            mail_merge: None,
            attached_template: None,
        }
    }

    /// Check if the document is protected.
    #[inline]
    pub fn is_protected(&self) -> bool {
        self.values.is_protected()
    }

    /// Get the type of protection applied.
    #[inline]
    pub fn protection_type(&self) -> Option<ProtectionType> {
        self.values.protection_type()
    }

    /// Check if track revisions is enabled.
    #[inline]
    pub fn track_revisions(&self) -> bool {
        self.values.track_revisions()
    }

    /// Get the zoom percentage.
    #[inline]
    pub fn zoom_percent(&self) -> Option<u32> {
        self.values.zoom_percent()
    }

    /// Return the declared smart-tag vocabularies in document order.
    #[inline]
    pub fn smart_tag_types(&self) -> &[SmartTagType] {
        &self.smart_tag_types
    }

    /// Whether embedded smart-tag data should be omitted when saving.
    #[inline]
    pub fn do_not_embed_smart_tags(&self) -> bool {
        self.do_not_embed_smart_tags
    }

    /// Return the document's inert mail-merge metadata, if present.
    #[inline]
    pub fn mail_merge(&self) -> Option<&MailMergeSettings> {
        self.mail_merge.as_ref()
    }

    /// Return the inert attached-template reference, if present.
    #[inline]
    pub fn attached_template(&self) -> Option<&AttachedTemplate> {
        self.attached_template.as_ref()
    }

    /// Return the on/off compatibility option flags in document order.
    #[inline]
    pub fn compatibility_options(&self) -> &[CompatibilityOption] {
        self.values.compatibility_options()
    }

    /// Return the `w:compatSetting` triples in document order.
    #[inline]
    pub fn compatibility_settings(&self) -> &[CompatibilitySetting] {
        self.values.compatibility_settings()
    }

    /// Look up a `w:compatSetting` triple by name and URI.
    pub fn compatibility_setting(&self, name: &str, uri: &str) -> Option<&CompatibilitySetting> {
        self.values.compatibility_setting(name, uri)
    }

    /// Return the Word compatibility mode (`compatibilityMode` value), when
    /// declared — for example `15` targets Word 2013 behavior.
    pub fn compatibility_mode(&self) -> Option<u32> {
        self.values.compatibility_mode()
    }

    /// Return the document-level footnote properties, if present.
    #[inline]
    pub fn footnote_properties(&self) -> Option<&NoteNumberingProperties<Format>> {
        self.values.footnote_properties()
    }

    /// Return the document-level endnote properties, if present.
    #[inline]
    pub fn endnote_properties(&self) -> Option<&NoteNumberingProperties<Format>> {
        self.values.endnote_properties()
    }

    /// Whether applications should recommend write protection for the
    /// document (`w:writeProtection`).
    #[inline]
    pub fn is_write_protected(&self) -> bool {
        self.values.is_write_protected()
    }

    /// Return the document view mode (`w:view`), when specified.
    #[inline]
    pub fn view(&self) -> Option<View> {
        self.values.view()
    }

    /// Return the proofing completion markers (`w:proofState`), if present.
    #[inline]
    pub fn proofing_state(&self) -> Option<&ProofingState> {
        self.values.proofing_state()
    }

    /// Return the default tab stop interval in twips (`w:defaultTabStop`),
    /// when specified.
    #[inline]
    pub fn default_tab_stop_twips(&self) -> Option<u32> {
        self.values.default_tab_stop_twips()
    }

    /// Return the theme font language defaults (`w:themeFontLang`), if
    /// present.
    #[inline]
    pub fn theme_font_languages(&self) -> Option<&ThemeFontLanguages> {
        self.values.theme_font_languages()
    }

    /// Return the theme color slot remapping (`w:clrSchemeMapping`), if
    /// present.
    #[inline]
    pub fn color_scheme_mapping(&self) -> Option<&ColorSchemeMapping> {
        self.values.color_scheme_mapping()
    }

    /// Serialize the editing view, proofing, and theme default elements
    /// (`w:writeProtection`, `w:view`, `w:proofState`, `w:defaultTabStop`,
    /// `w:themeFontLang`, `w:clrSchemeMapping`) in ECMA-376 schema order.
    pub fn to_editing_settings_xml(&self, prefix: &str) -> String {
        self.values.to_editing_settings_xml(prefix)
    }
}

impl Default for DocumentSettings {
    fn default() -> Self {
        Self::new()
    }
}
