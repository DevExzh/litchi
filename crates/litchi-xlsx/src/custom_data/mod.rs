//! Typed `SpreadsheetML` custom-data properties.
//!
//! [`crate::Package::custom_data`] reads the source-bound catalog;
//! [`crate::Package::edit_custom_data`] starts an [`Edit`] whose
//! [`Edit::commit`] returns a [`Commit`] holding the new package, the exact
//! [`Patch`] and the [`Diagnostics`]; [`crate::Package::apply_custom_data`]
//! replays a patch. The source package never changes.
//!
//! ```
//! use litchi_xlsx::Package;
//! use litchi_xlsx::custom_data::CustomData;
//!
//! # fn main() -> litchi_xlsx::Result<()> {
//! let package = Package::new()?;
//! let mut edit = package.edit_custom_data()?;
//! edit.upsert(CustomData::new("Storage-1", b"inert".to_vec()))?;
//! let (after, patch) = edit.commit()?.into_parts();
//! assert_eq!(after.custom_data()?.len(), 1);
//! assert!(package.custom_data()?.is_empty());
//! let restored = after.apply_custom_data(&patch.inverse())?.into_package();
//! assert!(restored.custom_data()?.is_empty());
//! # Ok(())
//! # }
//! ```
//!
//! The model and bounded XML codec are split by responsibility so callers
//! can reuse the package-neutral types without taking on OPC lifecycle code.

mod codec;
mod model;
mod package;

pub use codec::{XmlLimits, parse_properties, validate_workbook_root, write_properties};
pub use model::{CustomData, CustomDataView, ExtensionList, Properties, RemovalDisposition};
pub use package::{
    Commit, Diagnostics, Edit, EditOptions, Error, LimitKind, Limits, Part, Patch, Snapshot,
};

pub use litchi_ooxml_common::custom_data::{
    DATA_CONTENT_TYPE, DATA_RELATIONSHIP_TYPE, PROPERTIES_CONTENT_TYPE,
    PROPERTIES_RELATIONSHIP_TYPE,
};
