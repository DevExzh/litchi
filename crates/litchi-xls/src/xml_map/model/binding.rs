//! Data-binding model.

use super::identity::validate_string;
use super::opaque::OpaqueXml;
use super::schema::NamespaceDeclaration;
use crate::{Error, Result};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum LoadMode {
    None,
    Normal,
    DelayLoad,
    Asynchronous,
    ObjectModel,
}

impl LoadMode {
    #[must_use]
    pub const fn code(self) -> u32 {
        match self {
            Self::None => 0,
            Self::Normal => 1,
            Self::DelayLoad => 2,
            Self::Asynchronous => 3,
            Self::ObjectModel => 4,
        }
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn from_code(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::Normal),
            2 => Ok(Self::DelayLoad),
            3 => Ok(Self::Asynchronous),
            4 => Ok(Self::ObjectModel),
            _ => Err(invalid("DataBindingLoadMode must be in 0..=4")),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataBinding {
    pub(super) data_binding_name: Option<String>,
    pub(super) file_binding: String,
    pub(super) file_binding_name: Option<String>,
    pub(super) load_mode: LoadMode,
    pub(super) namespaces: Vec<NamespaceDeclaration>,
    pub(super) payload: Option<OpaqueXml>,
}

impl DataBinding {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn try_new(file_binding: impl Into<String>, load_mode: LoadMode) -> Result<Self> {
        Self::from_parts(
            None,
            file_binding.into(),
            None,
            load_mode.code().to_string(),
            None,
            Vec::new(),
        )
    }

    pub(crate) fn from_parts(
        data_binding_name: Option<String>,
        file_binding: String,
        file_binding_name: Option<String>,
        load_mode: String,
        payload: Option<OpaqueXml>,
        namespaces: Vec<(String, String)>,
    ) -> Result<Self> {
        let data_binding_name = data_binding_name
            .map(|value| validate_string(value, 65_535, "DataBindingName", false))
            .transpose()?;
        let file_binding = validate_string(file_binding, 65_535, "FileBinding", false)?;
        if matches!(file_binding.as_str(), "true" | "false") {
            return Err(invalid("FileBinding must not be true or false"));
        }
        let file_binding_name = file_binding_name
            .map(|value| validate_string(value, 65_535, "FileBindingName", false))
            .transpose()?;
        let load_mode =
            LoadMode::from_code(load_mode.parse::<u32>().map_err(|_error| {
                invalid("DataBindingLoadMode is not an unsigned decimal value")
            })?)?;
        let namespaces = namespaces
            .into_iter()
            .map(|(prefix, uri)| NamespaceDeclaration::try_new(prefix, uri))
            .collect::<Result<Vec<_>>>()?;
        Ok(Self {
            data_binding_name,
            file_binding,
            file_binding_name,
            load_mode,
            namespaces,
            payload,
        })
    }

    #[must_use]
    pub fn data_binding_name(&self) -> Option<&str> {
        self.data_binding_name.as_deref()
    }

    #[must_use]
    pub fn file_binding(&self) -> &str {
        &self.file_binding
    }

    #[must_use]
    pub fn file_binding_name(&self) -> Option<&str> {
        self.file_binding_name.as_deref()
    }

    #[must_use]
    pub const fn load_mode(&self) -> LoadMode {
        self.load_mode
    }

    #[must_use]
    pub fn namespaces(&self) -> &[NamespaceDeclaration] {
        &self.namespaces
    }

    #[must_use]
    pub fn payload(&self) -> Option<&OpaqueXml> {
        self.payload.as_ref()
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidData(format!("XML map: {}", message.into()))
}
