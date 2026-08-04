//! OpenDocument family discovery and ergonomic umbrella APIs.
//!
//! Family implementations live in dedicated crates so consumers can select a
//! small dependency and memory footprint. This crate owns only the common
//! vocabulary, inert format detection, and explicit family facades.

#![forbid(unsafe_code)]

pub use litchi_odf_common::{annotation, constants, coordinates, core, datatype, namespace};

/// ODF family detection without constructing a document model.
pub use litchi_odt::detect;

/// Dedicated OpenDocument Presentation implementation.
pub use litchi_odp as odp;
/// Dedicated OpenDocument Spreadsheet implementation.
pub use litchi_ods as ods;
/// Dedicated OpenDocument Text implementation.
pub use litchi_odt as odt;

pub use litchi_odf_common::core::{
    AutoReloadMetadata, DocumentStatistics, HyperlinkBehaviourMetadata, Manifest, ManifestChecksum,
    ManifestChecksumAlgorithm, ManifestEncryption, ManifestEncryptionAlgorithm, ManifestEntry,
    ManifestKeyDerivation, ManifestStartKeyGeneration, Metadata, OdfEncryptionCipher,
    OdfEncryptionKdf, OdfEncryptionProfile, OdfEncryptionStartKey, OdfStructure, OwnedPackage,
    PackageWriter, TemplateMetadata, UserDefinedMetadata, UserDefinedValueType,
};

pub use litchi_odp::{Presentation, PresentationBuilder};
pub use litchi_ods::{Spreadsheet, SpreadsheetBuilder};
pub use litchi_odt::{Document, DocumentBuilder, MutableDocument};
