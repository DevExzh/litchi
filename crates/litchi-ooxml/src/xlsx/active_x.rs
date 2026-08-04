//! Compatibility adapter for the canonical XLSX ActiveX persistence codec.
//!
//! SpreadsheetML control metadata, ActiveX descriptor XML, bounded limits, and
//! the inert OPC graph implementation live in `litchi_xlsx::active_x`.
//! This module preserves the historical OOXML-facing names and maps owner
//! failures back to `OoxmlError`. ActiveX payloads remain opaque.

use crate::error::{OoxmlError, Result};
use litchi_opc::{OpcPackage, PackURI};

use litchi_xlsx::active_x as owner;

pub use owner::{
    ActiveXFont, ActiveXPicture, ActiveXProperty, ActiveXPropertyObject, ControlProperties, Marker,
    ObjectAnchor, OpaqueActiveXBinary, OpaqueActiveXPreviewImage, Persistence, WorksheetControl,
};

/// The worksheet `controls` collection.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WorksheetControls {
    pub controls: Vec<WorksheetControl>,
}

impl WorksheetControls {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        owner::WorksheetControls::parse(xml)
            .map(Self::from_owner)
            .map_err(map_owner_error)
    }

    pub fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        self.to_owner().to_xml(strict).map_err(map_owner_error)
    }

    fn to_owner(&self) -> owner::WorksheetControls {
        owner::WorksheetControls {
            controls: self.controls.clone(),
        }
    }

    fn from_owner(value: owner::WorksheetControls) -> Self {
        Self {
            controls: value.controls,
        }
    }
}

/// The inert ActiveX descriptor/property XML document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ActiveXDescriptor {
    pub class_id: String,
    pub license: Option<String>,
    pub persistence: Persistence,
    pub relationship_id: Option<String>,
    pub properties: Vec<ActiveXProperty>,
}

impl ActiveXDescriptor {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        owner::ActiveXDescriptor::parse(xml)
            .map(Self::from_owner)
            .map_err(map_owner_error)
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        self.to_owner().to_xml().map_err(map_owner_error)
    }

    fn to_owner(&self) -> owner::ActiveXDescriptor {
        owner::ActiveXDescriptor {
            class_id: self.class_id.clone(),
            license: self.license.clone(),
            persistence: self.persistence,
            relationship_id: self.relationship_id.clone(),
            properties: self.properties.clone(),
        }
    }

    fn from_owner(value: owner::ActiveXDescriptor) -> Self {
        Self {
            class_id: value.class_id,
            license: value.license,
            persistence: value.persistence,
            relationship_id: value.relationship_id,
            properties: value.properties,
        }
    }
}

/// An inert ActiveX control descriptor and its opaque package resources.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LoadedActiveXControl {
    pub control: WorksheetControl,
    pub descriptor_uri: PackURI,
    pub descriptor: ActiveXDescriptor,
    pub binaries: Vec<OpaqueActiveXBinary>,
    pub preview: Option<OpaqueActiveXPreviewImage>,
}

/// A complete inert ActiveX graph attached to one worksheet.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ActiveXControlSet {
    pub controls: Vec<LoadedActiveXControl>,
}

impl ActiveXControlSet {
    fn to_owner(&self) -> owner::ActiveXControlSet {
        owner::ActiveXControlSet {
            controls: self
                .controls
                .iter()
                .map(LoadedActiveXControl::to_owner)
                .collect(),
        }
    }

    fn from_owner(value: owner::ActiveXControlSet) -> Self {
        Self {
            controls: value
                .controls
                .into_iter()
                .map(LoadedActiveXControl::from_owner)
                .collect(),
        }
    }
}

impl LoadedActiveXControl {
    fn to_owner(&self) -> owner::LoadedActiveXControl {
        owner::LoadedActiveXControl {
            control: self.control.clone(),
            descriptor_uri: self.descriptor_uri.clone(),
            descriptor: self.descriptor.to_owner(),
            binaries: self.binaries.clone(),
            preview: self.preview.clone(),
        }
    }

    fn from_owner(value: owner::LoadedActiveXControl) -> Self {
        Self {
            control: value.control,
            descriptor_uri: value.descriptor_uri,
            descriptor: ActiveXDescriptor::from_owner(value.descriptor),
            binaries: value.binaries,
            preview: value.preview,
        }
    }
}

/// Replace the direct worksheet `controls` collection while preserving
/// unrelated worksheet bytes.
pub fn replace_worksheet_controls_xml(xml: &[u8], controls: &WorksheetControls) -> Result<Vec<u8>> {
    owner::replace_worksheet_controls_xml(xml, &controls.to_owner()).map_err(map_owner_error)
}

/// Load one worksheet's complete inert ActiveX graph.
pub fn load_from_worksheet(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
) -> Result<ActiveXControlSet> {
    owner::load_from_worksheet(package, worksheet_uri)
        .map(ActiveXControlSet::from_owner)
        .map_err(map_owner_error)
}

/// Store a complete inert ActiveX graph on a worksheet that has no controls.
pub fn store_on_worksheet(
    package: &mut OpcPackage,
    worksheet_uri: &PackURI,
    value: &ActiveXControlSet,
) -> Result<()> {
    owner::store_on_worksheet(package, worksheet_uri, &value.to_owner()).map_err(map_owner_error)
}

/// Atomically replace one worksheet's complete inert ActiveX graph.
pub fn replace_on_worksheet(
    package: &mut OpcPackage,
    worksheet_uri: &PackURI,
    value: &ActiveXControlSet,
) -> Result<()> {
    owner::replace_on_worksheet(package, worksheet_uri, &value.to_owner()).map_err(map_owner_error)
}

/// Remove one worksheet's complete inert ActiveX graph.
pub fn remove_from_worksheet(package: &mut OpcPackage, worksheet_uri: &PackURI) -> Result<bool> {
    owner::remove_from_worksheet(package, worksheet_uri).map_err(map_owner_error)
}

fn map_owner_error(error: litchi_xlsx::Error) -> OoxmlError {
    match error {
        litchi_xlsx::Error::Package(error) => OoxmlError::Opc(error),
        litchi_xlsx::Error::MarkupCompatibility(error) => OoxmlError::from(error),
        litchi_xlsx::Error::Xml(error) => OoxmlError::Xml(error.to_string()),
        litchi_xlsx::Error::Common(error) => match error {
            litchi_ooxml_common::Error::ContentType { expected, actual } => {
                OoxmlError::InvalidContentType {
                    expected,
                    got: actual,
                }
            },
            litchi_ooxml_common::Error::Relationship(message) => {
                OoxmlError::InvalidRelationship(message)
            },
            litchi_ooxml_common::Error::Xml(message) => OoxmlError::Xml(message),
            litchi_ooxml_common::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
            other => OoxmlError::Common(other),
        },
        litchi_xlsx::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
        litchi_xlsx::Error::Allocation { resource, source } => {
            OoxmlError::Allocation { resource, source }
        },
        other => OoxmlError::Xlsx(other),
    }
}
