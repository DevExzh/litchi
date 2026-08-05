//! OpenDocument family discovery and ergonomic umbrella APIs.
//!
//! Family implementations live in dedicated crates so consumers can select a
//! small dependency and memory footprint. This crate owns only the common
//! vocabulary, inert format detection, and explicit family facades.

#![forbid(unsafe_code)]

pub use litchi_odf_common::detect;

/// Dedicated OpenDocument Database implementation.
#[cfg(feature = "odb")]
pub use litchi_odb as odb;
/// Dedicated OpenDocument Chart implementation.
#[cfg(feature = "odc")]
pub use litchi_odc as odc;
/// Dedicated OpenDocument Formula implementation.
#[cfg(feature = "formula")]
pub use litchi_odf_formula as formula;
/// Dedicated OpenDocument Drawing implementation.
#[cfg(feature = "odg")]
pub use litchi_odg as odg;
/// Dedicated OpenDocument Image implementation.
#[cfg(feature = "odi")]
pub use litchi_odi as odi;
/// Dedicated OpenDocument Master implementation.
#[cfg(feature = "odm")]
pub use litchi_odm as odm;
/// Dedicated OpenDocument Presentation implementation.
#[cfg(feature = "odp")]
pub use litchi_odp as odp;
/// Dedicated OpenDocument Spreadsheet implementation.
#[cfg(feature = "ods")]
pub use litchi_ods as ods;
/// Dedicated OpenDocument Text implementation.
#[cfg(feature = "odt")]
pub use litchi_odt as odt;
/// Dedicated OpenDocument Web Template implementation.
#[cfg(feature = "oth")]
pub use litchi_oth as oth;
