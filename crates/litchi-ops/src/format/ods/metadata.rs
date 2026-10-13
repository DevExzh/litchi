//! ODS document properties (record 0277): initial metadata at creation and
//! the `/metadata` set path.
//!
//! `set("/metadata", props)` stages `meta.xml` through the leaf's primary
//! `Edit::set_metadata`, which owns the fields ODS's retained-source
//! metadata patcher writes: `title`, `subject`, `description`, `keywords`
//! and `author`. `null` removes a field. Every other root-metadata key the
//! OLE and OOXML formats accept is a typed `unsupported_capability` naming
//! the property, and an unknown key is `unsupported_property`. A set never
//! stamps `modified` (determinism, ADR 0034 §9).

use litchi_core::{FileFormat, Metadata};
use litchi_ods::Package;

use super::write::{Commit, finish, leaf};
use super::{Ods, core_error, odf_error};
use crate::error::{Code, Error, Result};
use crate::props::{PropValue, Props};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// Build the minimal spreadsheet (one empty `Sheet1`) with creation-time
/// metadata, through the leaf builder's fresh-authoring metadata.
pub(super) fn create(metadata: &Metadata) -> Result<Package> {
    let mut builder = litchi_ods::Builder::new();
    builder.set_metadata(metadata.clone()).map_err(core_error)?;
    builder
        .add_sheet(litchi_ods::Sheet::new("Sheet1").map_err(core_error)?)
        .map_err(core_error)?;
    builder.build().map_err(odf_error)
}

/// Whether `target` is exactly the `/metadata` node.
pub(crate) fn is_target(target: &TargetArg<'_>) -> bool {
    target.path().is_some_and(|path| {
        let path = path.path();
        path.is_absolute()
            && matches!(path.steps(), [step]
                if step.predicates().is_empty() && step.name().eq_ignore_ascii_case("metadata"))
    })
}

/// The keys other formats accept on `/metadata` that ODS's metadata
/// patcher does not own.
const UNOWNED: [&str; 14] = [
    "application",
    "category",
    "company",
    "contentStatus",
    "contentType",
    "created",
    "extended.company",
    "extended.manager",
    "extended.template",
    "language",
    "lastModifiedBy",
    "lastPrinted",
    "manager",
    "modified",
];

const MORE_UNOWNED: [&str; 3] = ["revisionNumber", "template", "version"];

/// Stage `props` onto the snapshot's metadata in one leaf edit.
pub(super) fn set(file: &Ods, props: &Props<'_>, limits: &SemanticLimits) -> Result<Commit> {
    if props.is_empty() {
        return Err(Error::new(
            Code::InvalidArgument,
            "a metadata set needs at least one property",
        )
        .with_format(FileFormat::Ods));
    }
    let mut metadata = file.metadata_ref()?.clone();
    for (key, value) in props.iter() {
        let name = key.as_str();
        let field = if name.eq_ignore_ascii_case("title") {
            &mut metadata.title
        } else if name.eq_ignore_ascii_case("subject") {
            &mut metadata.subject
        } else if name.eq_ignore_ascii_case("description") {
            &mut metadata.description
        } else if name.eq_ignore_ascii_case("keywords") {
            &mut metadata.keywords
        } else if name.eq_ignore_ascii_case("author") {
            &mut metadata.author
        } else if UNOWNED
            .iter()
            .chain(MORE_UNOWNED.iter())
            .any(|key| key.eq_ignore_ascii_case(name))
            || name
                .get(..7)
                .is_some_and(|prefix| prefix.eq_ignore_ascii_case("custom."))
        {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the ODS metadata writer owns title, subject, description, keywords and author",
            )
            .with_property(name.to_owned())
            .with_format(FileFormat::Ods)
            .with_suggestion("follow-up: ODS metadata family (record 0277 gap M1)"));
        } else {
            return Err(Error::new(
                Code::UnsupportedProperty,
                "the metadata key is not a root-metadata key",
            )
            .with_property(name.to_owned())
            .with_format(FileFormat::Ods)
            .with_suggestion("use `title`, `subject`, `description`, `keywords` or `author`"));
        };
        *field = match value {
            PropValue::Null => None,
            PropValue::Text(text) => Some(text.to_string()),
            _ => {
                return Err(
                    Error::new(Code::InvalidValue, "the metadata property requires text")
                        .with_property(name.to_owned())
                        .with_format(FileFormat::Ods),
                );
            },
        };
    }
    let mut edit = file.leaf_edit()?;
    edit.set_metadata(&metadata)
        .map_err(|error| leaf(error, Code::UnsupportedCapability))?;
    finish(file, edit, props.len(), limits)
}
