//! Inert chartsheet extension children.

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Extension {
    pub uri: String,
    /// Canonical, namespace-aware XML for the single wildcard child of `ext`.
    pub payload_xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtensionList {
    pub extensions: Vec<Extension>,
}
