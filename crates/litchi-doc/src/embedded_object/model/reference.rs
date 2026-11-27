//! Read-only reference to a managed DOC embedded-object field.

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Reference {
    pub storage_id: u32,
    pub storage_name: String,
    pub start_cp: u32,
    pub separator_cp: u32,
    pub end_cp: u32,
    pub data_offset: u32,
}
