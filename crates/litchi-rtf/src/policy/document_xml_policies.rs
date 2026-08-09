/// Passive web-save and custom-XML document policies.
///
/// Every field preserves omission separately from explicit `0` or `1`. This
/// crate does not save web pages, validate or fetch schemas, execute
/// transforms, activate custom XML, or display validation UI.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentXmlPolicies {
    /// Explicit `relyonvmlN` policy.
    pub rely_on_vml: Option<bool>,
    /// Explicit `validatexmlN` policy.
    pub validate_custom_xml: Option<bool>,
    /// Explicit `showplaceholdtextN` policy.
    pub show_placeholder_text: Option<bool>,
    /// Explicit `ignoremixedcontentN` policy.
    pub ignore_mixed_content: Option<bool>,
    /// Explicit `saveinvalidxmlN` policy.
    pub save_invalid_xml: Option<bool>,
    /// Explicit `showxmlerrorsN` policy.
    pub show_xml_errors: Option<bool>,
}

impl DocumentXmlPolicies {
    /// Return whether all six policy controls were omitted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.rely_on_vml.is_none()
            && self.validate_custom_xml.is_none()
            && self.show_placeholder_text.is_none()
            && self.ignore_mixed_content.is_none()
            && self.save_invalid_xml.is_none()
            && self.show_xml_errors.is_none()
    }

    /// Return the effective VML policy; omission has the same effect as `0`.
    #[must_use]
    pub fn effective_rely_on_vml(&self) -> bool {
        self.rely_on_vml.unwrap_or(false)
    }

    /// Return the effective validation policy; omission requests validation when supported.
    #[must_use]
    pub fn effective_validate_custom_xml(&self) -> bool {
        self.validate_custom_xml.unwrap_or(true)
    }

    #[must_use]
    pub fn effective_show_placeholder_text(&self) -> bool {
        self.show_placeholder_text.unwrap_or(false)
    }

    #[must_use]
    pub fn effective_ignore_mixed_content(&self) -> bool {
        self.ignore_mixed_content.unwrap_or(false)
    }

    #[must_use]
    pub fn effective_save_invalid_xml(&self) -> bool {
        self.save_invalid_xml.unwrap_or(false)
    }

    #[must_use]
    pub fn effective_show_xml_errors(&self) -> bool {
        self.show_xml_errors.unwrap_or(false)
    }
}
