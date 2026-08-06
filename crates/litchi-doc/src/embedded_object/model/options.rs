//! Authoring options for one embedded DOC object.

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct WriteOptions {
    pub storage_id: u32,
    pub instruction: String,
    /// Complete PICFAndOfficeArtData block for the Data stream.
    pub picture_data: Vec<u8>,
    /// Standalone CFB to install as `ObjectPool/_<storage_id>`.
    pub compound_file: Vec<u8>,
}

impl WriteOptions {
    pub fn new(storage_id: u32, compound_file: Vec<u8>, picture_data: Vec<u8>) -> Self {
        Self {
            storage_id,
            instruction: format!(" EMBED LITCHI_OBJECT _{storage_id} "),
            picture_data,
            compound_file,
        }
    }
}
