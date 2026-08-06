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
pub(crate) use mutation::{add_xml, remove_xml, replace_xml};
pub(crate) use package::{add, remove, reorder, replace, update};

use litchi_core::{Error, Result};

pub(crate) fn annotations(content: &str, host: AnnotationHost) -> Result<Vec<AnnotationInfo>> {
    let scan = scan::scan(content, host)?;
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
) -> Result<Option<AnnotationInfo>> {
    if name.is_empty() {
        return invalid("annotation name cannot be empty");
    }
    Ok(annotations(content, host)?
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
                self.content.xml_content(),
                crate::package::annotation::AnnotationHost::$host,
            )
        }

        /// Find a uniquely named annotation.
        pub fn find_annotation(
            &self,
            name: &str,
        ) -> litchi_core::Result<Option<crate::AnnotationInfo>> {
            crate::package::annotation::find_annotation(
                self.content.xml_content(),
                crate::package::annotation::AnnotationHost::$host,
                name,
            )
        }

        /// Add a point or named-range annotation atomically.
        pub fn add_annotation(
            &mut self,
            anchor: &crate::AnnotationAnchor,
            annotation: &crate::Annotation,
        ) -> litchi_core::Result<usize> {
            let (bytes, index) = crate::package::annotation::add(
                &self.package,
                self.content.xml_content(),
                crate::package::annotation::AnnotationHost::$host,
                anchor,
                annotation,
            )?;
            let replacement = Self::from_bytes(bytes)?;
            *self = replacement;
            Ok(index)
        }

        /// Replace an annotation body and metadata while retaining its anchor.
        pub fn replace_annotation(
            &mut self,
            index: usize,
            annotation: &crate::Annotation,
        ) -> litchi_core::Result<()> {
            let bytes = crate::package::annotation::replace(
                &self.package,
                self.content.xml_content(),
                crate::package::annotation::AnnotationHost::$host,
                index,
                annotation,
            )?;
            let replacement = Self::from_bytes(bytes)?;
            *self = replacement;
            Ok(())
        }

        /// Apply a partial typed annotation metadata update.
        pub fn update_annotation(
            &mut self,
            index: usize,
            update: &crate::AnnotationUpdate,
        ) -> litchi_core::Result<()> {
            let bytes = crate::package::annotation::update(
                &self.package,
                self.content.xml_content(),
                crate::package::annotation::AnnotationHost::$host,
                index,
                update,
            )?;
            let replacement = Self::from_bytes(bytes)?;
            *self = replacement;
            Ok(())
        }

        /// Remove an annotation and its paired end marker, if any.
        pub fn remove_annotation(&mut self, index: usize) -> litchi_core::Result<()> {
            let bytes = crate::package::annotation::remove(
                &self.package,
                self.content.xml_content(),
                crate::package::annotation::AnnotationHost::$host,
                index,
            )?;
            let replacement = Self::from_bytes(bytes)?;
            *self = replacement;
            Ok(())
        }

        /// Reorder point annotations that are direct XML siblings.
        pub fn reorder_annotation(&mut self, from: usize, to: usize) -> litchi_core::Result<()> {
            let bytes = crate::package::annotation::reorder(
                &self.package,
                self.content.xml_content(),
                crate::package::annotation::AnnotationHost::$host,
                from,
                to,
            )?;
            let replacement = Self::from_bytes(bytes)?;
            *self = replacement;
            Ok(())
        }
    };
}

pub(crate) use annotation_facade_methods;

#[cfg(test)]
mod tests;
