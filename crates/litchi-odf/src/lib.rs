//! OpenDocument family discovery and ergonomic umbrella APIs.
//!
//! Family implementations live in dedicated crates so consumers can select a
//! small dependency and memory footprint. This crate owns only the common
//! vocabulary, inert format detection, and explicit family facades.

#![forbid(unsafe_code)]

pub use litchi_odf_common::detect;
pub use litchi_odf_common::{annotation, constants, coordinates, core, datatype, namespace, rdf};

/// Dedicated OpenDocument Database implementation.
pub use litchi_odb as odb;
/// Dedicated OpenDocument Chart implementation.
pub use litchi_odc as odc;
/// Dedicated OpenDocument Drawing implementation.
pub use litchi_odg as odg;
/// Dedicated OpenDocument Image implementation.
pub use litchi_odi as odi;
/// Dedicated OpenDocument Master implementation.
pub use litchi_odm as odm;
/// Dedicated OpenDocument Presentation implementation.
pub use litchi_odp as odp;
/// Dedicated OpenDocument Spreadsheet implementation.
pub use litchi_ods as ods;
/// Dedicated OpenDocument Text implementation.
pub use litchi_odt as odt;
/// Dedicated OpenDocument Web Template implementation.
pub use litchi_oth as oth;

pub use litchi_odf_common::core::{
    AutoReloadMetadata, DocumentStatistics, HyperlinkBehaviourMetadata, Manifest, ManifestChecksum,
    ManifestChecksumAlgorithm, ManifestEncryption, ManifestEncryptionAlgorithm, ManifestEntry,
    ManifestKeyDerivation, ManifestStartKeyGeneration, Metadata, OdfEncryptionCipher,
    OdfEncryptionKdf, OdfEncryptionProfile, OdfEncryptionStartKey, OdfStructure, OwnedPackage,
    PackageWriter, TemplateMetadata, UserDefinedMetadata, UserDefinedValueType,
};

pub use litchi_odb::{Database, DatabaseBuilder};
pub use litchi_odc::{Chart, ChartBuilder};
pub use litchi_odg::{Drawing, DrawingBuilder};
pub use litchi_odi::{Image, ImageBuilder};
pub use litchi_odm::{Master, MasterBuilder};
pub use litchi_odp::{Presentation, PresentationBuilder};
pub use litchi_ods::{Spreadsheet, SpreadsheetBuilder};
pub use litchi_odt::{Document, DocumentBuilder, MutableDocument};
pub use litchi_oth::{WebTemplate, WebTemplateBuilder};
