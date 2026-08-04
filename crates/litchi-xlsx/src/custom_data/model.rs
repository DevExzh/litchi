//! Package-neutral custom-data properties model.

/// One self-contained `x14:extLst` subtree, retained without interpretation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtensionList {
    pub xml: Vec<u8>,
}

/// Typed properties for one embedded custom-data storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Properties {
    pub id: String,
    pub extension_list: Option<ExtensionList>,
}
