//! The MathType and Equation Editor equations embedded in a Word document.
//!
//! Word stores each embedded OLE object as a storage `ObjectPool/_<id>`,
//! where `<id>` is the picture location (`sprmCPicLocation`) of the run that
//! shows the object. An equation object holds its MTEF data in an
//! `Equation Native` stream; `litchi-equation` validates the stream's OLE
//! header and decodes the MTEF.

use litchi_cfb::OleFile;
use std::collections::HashMap;
use std::io::{Read, Seek};

/// The `ObjectPool` storage that holds embedded OLE objects.
const OBJECT_POOL: &str = "ObjectPool";
/// The stream of an equation object that holds its MTEF data.
const EQUATION_NATIVE: &str = "Equation Native";
/// Directory entry type of a storage (MS-CFB 2.6.1 `STGTY_STORAGE`).
const STORAGE: u8 = 1;

/// The `Equation Native` stream of every `ObjectPool` object that has one,
/// keyed by the object's storage name (`_<id>`).
///
/// # Errors
///
/// Returns the container error when the `ObjectPool` directory cannot be
/// listed or an `Equation Native` stream cannot be read. An object without
/// an `Equation Native` stream is not an equation and is skipped.
pub(crate) fn equation_native_streams<R: Read + Seek>(
    ole: &mut OleFile<R>,
) -> Result<HashMap<String, Vec<u8>>, litchi_cfb::OleError> {
    let mut streams = HashMap::new();
    if !ole.directory_exists(&[OBJECT_POOL]) {
        return Ok(streams);
    }
    let storages: Vec<String> = ole
        .list_directory_entries(&[OBJECT_POOL])?
        .into_iter()
        .filter(|entry| entry.entry_type == STORAGE && entry.name.starts_with('_'))
        .map(|entry| entry.name.clone())
        .collect();
    for storage in storages {
        match ole.open_stream(&[OBJECT_POOL, &storage, EQUATION_NATIVE]) {
            Ok(data) => {
                streams.insert(storage, data);
            },
            Err(litchi_cfb::OleError::StreamNotFound) => {},
            Err(error) => return Err(error),
        }
    }
    Ok(streams)
}

/// The picture location an `ObjectPool` storage name (`_<id>`) encodes, or
/// zero for a name that does not encode one.
pub(crate) fn storage_id(storage: &str) -> u32 {
    storage
        .strip_prefix('_')
        .and_then(|id| id.parse().ok())
        .unwrap_or(0)
}

#[cfg(test)]
mod tests {
    use super::storage_id;

    #[test]
    fn storage_names_encode_the_picture_location() {
        assert_eq!(storage_id("_1234567890"), 1_234_567_890);
        assert_eq!(storage_id("_x"), 0);
        assert_eq!(storage_id("1234"), 0);
    }
}
