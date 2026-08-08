//! Semantic owner for the legacy XLS `XML` stream ([MS-XLS] 2.1.7.22).
//!
//! The facade is intentionally prefix-free: callers work with [`MapInfo`],
//! [`Schema`], [`Map`], and [`DataBinding`] in this contextual module. Schema
//! and binding children are retained as inert opaque XML; this crate never
//! resolves schemas, opens bound files, or performs import/export work.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{parse, write};
pub use model::{
    DataBinding, LoadMode, Map, MapId, MapInfo, NamespaceDeclaration, OpaqueXml, Schema, SchemaId,
    XPath,
};

pub(crate) use validation::{
    validate as validate_info, validate_list_columns, validate_list_objects,
};
pub(crate) const STREAM_NAME: &str = "XML";

pub(crate) fn parse_stream_if_present<R: std::io::Read + std::io::Seek>(
    ole_file: &mut litchi_cfb::OleFile<R>,
) -> crate::Result<Option<MapInfo>> {
    let paths = ole_file
        .list_streams()
        .into_iter()
        .filter(|path| path.len() == 1 && path[0] == STREAM_NAME)
        .collect::<Vec<_>>();
    if paths.len() > 1 {
        return Err(crate::Error::InvalidData(
            "XML map: OLE package contains multiple XML streams".to_string(),
        ));
    }
    let Some(path) = paths.first() else {
        return Ok(None);
    };
    let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
    let data = ole_file.open_stream(&refs)?;
    parse(&data).map(Some)
}
