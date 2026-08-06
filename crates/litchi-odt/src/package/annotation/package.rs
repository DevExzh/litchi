//! Package rebuild and transactional annotation facade operations.

use super::mutation::{add_xml, remove_xml, reorder_xml, replace_xml};
use super::{Annotation, AnnotationAnchor, AnnotationHost, AnnotationUpdate, annotations, bounds};
use crate::core::OwnedPackage;
use crate::package::charts::rebuild_package;
use litchi_core::Result;

pub(crate) fn add(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    anchor: &AnnotationAnchor,
    annotation: &Annotation,
) -> Result<(Vec<u8>, usize)> {
    let (updated, index) = add_xml(content, host, anchor, annotation)?;
    rebuild(package, &updated).map(|bytes| (bytes, index))
}

pub(crate) fn replace(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    index: usize,
    annotation: &Annotation,
) -> Result<Vec<u8>> {
    rebuild(package, &replace_xml(content, host, index, annotation)?)
}

pub(crate) fn update(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    index: usize,
    update: &AnnotationUpdate,
) -> Result<Vec<u8>> {
    let items = annotations(content, host)?;
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
    replace(package, content, host, index, &info.annotation)
}

pub(crate) fn remove(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    index: usize,
) -> Result<Vec<u8>> {
    rebuild(package, &remove_xml(content, host, index)?)
}

pub(crate) fn reorder(
    package: &OwnedPackage,
    content: &str,
    host: AnnotationHost,
    from: usize,
    to: usize,
) -> Result<Vec<u8>> {
    rebuild(package, &reorder_xml(content, host, from, to)?)
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
