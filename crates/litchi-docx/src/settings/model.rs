#![expect(
    clippy::struct_field_names,
    reason = "the public model retains established field names"
)]
use super::colors::ColorSchemeMapping;
use super::compatibility::{
    COMPATIBILITY_MODE_SETTING_NAME, COMPATIBILITY_SETTING_URI, CompatibilityOption,
    CompatibilitySetting,
};
use super::editing::{ProofingState, ProtectionType, ThemeFontLanguages, View};
use super::extensions::Extensions;
use super::notes::{NoteNumberFormat, NoteNumberingProperties};

/// Format-owned scalar settings extracted from a Word settings part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Settings<F = NoteNumberFormat> {
    pub(crate) protected: bool,
    pub(crate) protection_type: Option<ProtectionType>,
    pub(crate) track_revisions: bool,
    pub(crate) zoom_percent: Option<u32>,
    pub(crate) compatibility_options: Vec<CompatibilityOption>,
    pub(crate) compatibility_settings: Vec<CompatibilitySetting>,
    pub(crate) footnote_properties: Option<NoteNumberingProperties<F>>,
    pub(crate) endnote_properties: Option<NoteNumberingProperties<F>>,
    pub(crate) write_protection: bool,
    pub(crate) view: Option<View>,
    pub(crate) proofing_state: Option<ProofingState>,
    pub(crate) default_tab_stop_twips: Option<u32>,
    pub(crate) theme_font_languages: Option<ThemeFontLanguages>,
    pub(crate) color_scheme_mapping: Option<ColorSchemeMapping>,
    pub(crate) extensions: Extensions,
}

impl<F> Default for Settings<F> {
    fn default() -> Self {
        Self {
            protected: false,
            protection_type: None,
            track_revisions: false,
            zoom_percent: None,
            compatibility_options: Vec::new(),
            compatibility_settings: Vec::new(),
            footnote_properties: None,
            endnote_properties: None,
            write_protection: false,
            view: None,
            proofing_state: None,
            default_tab_stop_twips: None,
            theme_font_languages: None,
            color_scheme_mapping: None,
            extensions: Extensions::new(),
        }
    }
}

impl<F> Settings<F> {
    /// Create empty settings values.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the document-protection marker.
    pub fn set_protected(&mut self, value: bool) -> &mut Self {
        self.protected = value;
        self
    }

    /// Set the document-protection editing mode.
    pub fn set_protection_type(&mut self, value: Option<ProtectionType>) -> &mut Self {
        self.protection_type = value;
        self
    }

    /// Set the tracked-revisions marker.
    pub fn set_track_revisions(&mut self, value: bool) -> &mut Self {
        self.track_revisions = value;
        self
    }

    /// Set the document zoom percentage.
    pub fn set_zoom_percent(&mut self, value: Option<u32>) -> &mut Self {
        self.zoom_percent = value;
        self
    }

    /// Replace the compatibility option sequence.
    pub fn set_compatibility_options(&mut self, value: Vec<CompatibilityOption>) -> &mut Self {
        self.compatibility_options = value;
        self
    }

    /// Replace the compatibility-setting sequence.
    pub fn set_compatibility_settings(&mut self, value: Vec<CompatibilitySetting>) -> &mut Self {
        self.compatibility_settings = value;
        self
    }

    /// Replace the footnote properties.
    pub fn set_footnote_properties(
        &mut self,
        value: Option<NoteNumberingProperties<F>>,
    ) -> &mut Self {
        self.footnote_properties = value;
        self
    }

    /// Replace the endnote properties.
    pub fn set_endnote_properties(
        &mut self,
        value: Option<NoteNumberingProperties<F>>,
    ) -> &mut Self {
        self.endnote_properties = value;
        self
    }

    /// Set the write-protection marker.
    pub fn set_write_protected(&mut self, value: bool) -> &mut Self {
        self.write_protection = value;
        self
    }

    /// Set the document view mode.
    pub fn set_view(&mut self, value: Option<View>) -> &mut Self {
        self.view = value;
        self
    }

    /// Replace the proofing state.
    pub fn set_proofing_state(&mut self, value: Option<ProofingState>) -> &mut Self {
        self.proofing_state = value;
        self
    }

    /// Set the default tab stop interval.
    pub fn set_default_tab_stop_twips(&mut self, value: Option<u32>) -> &mut Self {
        self.default_tab_stop_twips = value;
        self
    }

    /// Replace the theme-font language defaults.
    pub fn set_theme_font_languages(&mut self, value: Option<ThemeFontLanguages>) -> &mut Self {
        self.theme_font_languages = value;
        self
    }

    /// Replace the theme color mapping.
    pub fn set_color_scheme_mapping(&mut self, value: Option<ColorSchemeMapping>) -> &mut Self {
        self.color_scheme_mapping = value;
        self
    }

    /// Replace the ordered Word 2010/2012 settings extensions.
    pub fn set_extensions(&mut self, value: Extensions) -> &mut Self {
        self.extensions = value;
        self
    }
}

impl<F: Copy> Settings<F> {
    /// Check if the document is protected.
    #[inline]
    pub const fn is_protected(&self) -> bool {
        self.protected
    }

    /// Get the type of protection applied.
    #[inline]
    pub const fn protection_type(&self) -> Option<ProtectionType> {
        self.protection_type
    }

    /// Check if track revisions is enabled.
    #[inline]
    pub const fn track_revisions(&self) -> bool {
        self.track_revisions
    }

    /// Get the zoom percentage.
    #[inline]
    pub const fn zoom_percent(&self) -> Option<u32> {
        self.zoom_percent
    }

    /// Return the on/off compatibility options in document order.
    #[inline]
    pub fn compatibility_options(&self) -> &[CompatibilityOption] {
        &self.compatibility_options
    }

    /// Return the `w:compatSetting` triples in document order.
    #[inline]
    pub fn compatibility_settings(&self) -> &[CompatibilitySetting] {
        &self.compatibility_settings
    }

    /// Look up a `w:compatSetting` triple by name and URI.
    pub fn compatibility_setting(&self, name: &str, uri: &str) -> Option<&CompatibilitySetting> {
        self.compatibility_settings
            .iter()
            .find(|setting| setting.name() == name && setting.uri() == uri)
    }

    /// Return the Word compatibility mode, when declared.
    pub fn compatibility_mode(&self) -> Option<u32> {
        self.compatibility_setting(COMPATIBILITY_MODE_SETTING_NAME, COMPATIBILITY_SETTING_URI)
            .and_then(|setting| setting.value().parse().ok())
    }

    /// Return the document-level footnote properties, if present.
    #[inline]
    pub fn footnote_properties(&self) -> Option<&NoteNumberingProperties<F>> {
        self.footnote_properties.as_ref()
    }

    /// Return the document-level endnote properties, if present.
    #[inline]
    pub fn endnote_properties(&self) -> Option<&NoteNumberingProperties<F>> {
        self.endnote_properties.as_ref()
    }

    /// Whether applications should recommend write protection.
    #[inline]
    pub const fn is_write_protected(&self) -> bool {
        self.write_protection
    }

    /// Return the document view mode, when specified.
    #[inline]
    pub const fn view(&self) -> Option<View> {
        self.view
    }

    /// Return the proofing completion markers, if present.
    #[inline]
    pub fn proofing_state(&self) -> Option<&ProofingState> {
        self.proofing_state.as_ref()
    }

    /// Return the default tab stop interval in twips, when specified.
    #[inline]
    pub const fn default_tab_stop_twips(&self) -> Option<u32> {
        self.default_tab_stop_twips
    }

    /// Return the theme font language defaults, if present.
    #[inline]
    pub fn theme_font_languages(&self) -> Option<&ThemeFontLanguages> {
        self.theme_font_languages.as_ref()
    }

    /// Return the theme color slot remapping, if present.
    #[inline]
    pub fn color_scheme_mapping(&self) -> Option<&ColorSchemeMapping> {
        self.color_scheme_mapping.as_ref()
    }

    /// Return the ordered Word 2010/2012 settings extensions.
    #[inline]
    pub fn extensions(&self) -> &Extensions {
        &self.extensions
    }
}

impl Settings<NoteNumberFormat> {
    /// Map the owner-local note format to a host format without reparsing XML.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn try_map_note_format<G, E>(
        self,
        mut map: impl FnMut(NoteNumberFormat) -> Result<G, E>,
    ) -> Result<Settings<G>, E> {
        let Settings {
            protected,
            protection_type,
            track_revisions,
            zoom_percent,
            compatibility_options,
            compatibility_settings,
            footnote_properties,
            endnote_properties,
            write_protection,
            view,
            proofing_state,
            default_tab_stop_twips,
            theme_font_languages,
            color_scheme_mapping,
            extensions,
        } = self;
        Ok(Settings {
            protected,
            protection_type,
            track_revisions,
            zoom_percent,
            compatibility_options,
            compatibility_settings,
            footnote_properties: footnote_properties
                .map(|value| value.try_map_format(&mut map))
                .transpose()?,
            endnote_properties: endnote_properties
                .map(|value| value.try_map_format(&mut map))
                .transpose()?,
            write_protection,
            view,
            proofing_state,
            default_tab_stop_twips,
            theme_font_languages,
            color_scheme_mapping,
            extensions,
        })
    }
}
