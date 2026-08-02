/// Passive document embedding policies from the RTF header.
///
/// These values are retained for round-tripping only. This crate does not
/// embed fonts, device data, linguistic data, handwriting, or controls data.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentEmbeddingPolicies {
    /// Explicit `donotembedsysfontN` policy.
    pub do_not_embed_system_fonts: Option<bool>,
    /// Explicit `donotembedlingdataN` policy.
    pub do_not_embed_linguistic_data: Option<bool>,
}

impl DocumentEmbeddingPolicies {
    /// Return whether both embedding-policy controls were omitted.
    pub fn is_empty(&self) -> bool {
        self.do_not_embed_system_fonts.is_none() && self.do_not_embed_linguistic_data.is_none()
    }

    /// Return the system-font policy; omission has the same effect as `1`.
    pub fn effective_do_not_embed_system_fonts(&self) -> bool {
        self.do_not_embed_system_fonts.unwrap_or(true)
    }
}
