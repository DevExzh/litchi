//! Lossless, package-safe ODF annotation discovery and mutation.

mod codec;
mod model;
mod mutation;
mod package;
mod scan;

pub(crate) use model::AnnotationHost;
pub use model::{
    Annotation, AnnotationAnchor, AnnotationInfo, AnnotationPosition, AnnotationUpdate,
};
pub(crate) use package::{add, remove, reorder, replace, update};

use litchi_core::{Error, Result};
use litchi_odf_common::XmlStreamLimits;

pub(crate) fn annotations(
    content: &str,
    host: AnnotationHost,
    limits: &XmlStreamLimits,
) -> Result<Vec<AnnotationInfo>> {
    let scan = scan::scan(content, host, limits)?;
    scan.records
        .into_iter()
        .enumerate()
        .map(|(index, mut record)| {
            Ok(AnnotationInfo {
                index,
                annotation: record
                    .annotation
                    .take()
                    .ok_or_else(|| invalid_error("unterminated annotation"))?,
                anchor: AnnotationAnchor {
                    start: record.start_position,
                    end: record.end.map(|(_, position)| position),
                },
            })
        })
        .collect()
}

pub(crate) fn find_annotation(
    content: &str,
    host: AnnotationHost,
    name: &str,
    limits: &XmlStreamLimits,
) -> Result<Option<AnnotationInfo>> {
    if name.is_empty() {
        return invalid("annotation name cannot be empty");
    }
    Ok(annotations(content, host, limits)?
        .into_iter()
        .find(|item| item.annotation.name() == Some(name)))
}

fn bounds(index: usize, len: usize) -> Error {
    invalid_error(format!(
        "annotation index {index} is out of bounds for {len} entries"
    ))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

macro_rules! annotation_facade_methods {
    ($host:ident) => {
        /// Inspect annotations in document order without following body links.
        pub fn annotations(&self) -> litchi_core::Result<Vec<crate::AnnotationInfo>> {
            crate::package::annotation::annotations(
                self.state.content.xml_content(),
                crate::package::annotation::AnnotationHost::$host,
                &self.state.xml_limits,
            )
        }

        /// Find a uniquely named annotation.
        pub fn find_annotation(
            &self,
            name: &str,
        ) -> litchi_core::Result<Option<crate::AnnotationInfo>> {
            crate::package::annotation::find_annotation(
                self.state.content.xml_content(),
                crate::package::annotation::AnnotationHost::$host,
                name,
                &self.state.xml_limits,
            )
        }
    };
}

pub(crate) use annotation_facade_methods;

#[cfg(test)]
mod tests;
