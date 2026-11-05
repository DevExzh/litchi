//! Package rebuild and transactional annotation facade operations.

use super::mutation::{add_xml, remove_xml, reorder_xml, replace_xml};
use super::{Annotation, AnnotationAnchor, AnnotationHost, AnnotationUpdate, annotations, bounds};
use crate::core::OwnedPackage;
use crate::package::charts::rebuild_package;
use litchi_core::Result;
use litchi_odf_common::XmlStreamLimits;

pub(crate) fn add(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    anchor: &AnnotationAnchor,
    annotation: &Annotation,
    limits: &XmlStreamLimits,
) -> Result<(Vec<u8>, usize)> {
    let (updated, index) = add_xml(content, host, anchor, annotation, limits)?;
    rebuild(package, &updated).map(|bytes| (bytes, index))
}

pub(crate) fn replace(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    index: usize,
    annotation: &Annotation,
    limits: &XmlStreamLimits,
) -> Result<Vec<u8>> {
    rebuild(
        package,
        &replace_xml(content, host, index, annotation, limits)?,
    )
}

pub(crate) fn update(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    index: usize,
    update: &AnnotationUpdate,
    limits: &XmlStreamLimits,
) -> Result<Vec<u8>> {
    let items = annotations(content, host, limits)?;
    let len = items.len();
    let mut info = items
        .into_iter()
        .nth(index)
        .ok_or_else(|| bounds(index, len))?;
    if let Some(value) = &update.creator {
        info.annotation.set_creator(value.as_deref());
    }
    if let Some(value) = &update.date {
        info.annotation.set_date(value.as_deref());
    }
    if let Some(value) = &update.date_string {
        info.annotation.set_date_string(value.as_deref());
    }
    if let Some(value) = &update.initials {
        info.annotation.set_initials(value.as_deref());
    }
    if let Some(value) = update.display {
        info.annotation.set_display(value);
    }
    replace(package, content, host, index, &info.annotation, limits)
}

pub(crate) fn remove(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    index: usize,
    limits: &XmlStreamLimits,
) -> Result<Vec<u8>> {
    rebuild(package, &remove_xml(content, host, index, limits)?)
}

pub(crate) fn reorder(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    from: usize,
    to: usize,
    limits: &XmlStreamLimits,
) -> Result<Vec<u8>> {
    rebuild(package, &reorder_xml(content, host, from, to, limits)?)
}

fn rebuild(package: &OwnedPackage, content: &str) -> Result<Vec<u8>> {
    rebuild_package(
        package,
        content,
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
    )
}
