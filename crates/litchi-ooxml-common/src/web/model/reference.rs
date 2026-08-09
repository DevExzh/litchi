use super::super::Result;
use super::super::codec::{invalid, require_nonempty};
use super::super::validation::{
    validate_binding, validate_extension_list, validate_store_reference,
};
use super::{ExtKind, ExtList};
/// Catalog provider type from MS-OWEXML section 2.2.5.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Store {
    Omex,
    #[default]
    SharePointCatalog,
    SharePointApp,
    Exchange,
    /// File-system provider. Author references with [`Reference::file`].
    FileSystem,
    Registry,
    ExchangeCatalog,
    WopiCatalog,
}

impl Store {
    #[must_use]
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Omex => "OMEX",
            Self::SharePointCatalog => "SPCatalog",
            Self::SharePointApp => "SPApp",
            Self::Exchange => "Exchange",
            Self::FileSystem => "FileSystem",
            Self::Registry => "Registry",
            Self::ExchangeCatalog => "ExCatalog",
            Self::WopiCatalog => "WOPICatalog",
        }
    }

    pub(in crate::web) fn parse(value: &str) -> Result<Self> {
        match value {
            "OMEX" => Ok(Self::Omex),
            "SPCatalog" => Ok(Self::SharePointCatalog),
            "SPApp" => Ok(Self::SharePointApp),
            "Exchange" => Ok(Self::Exchange),
            "FileSystem" => Ok(Self::FileSystem),
            "Registry" => Ok(Self::Registry),
            "ExCatalog" => Ok(Self::ExchangeCatalog),
            "WOPICatalog" => Ok(Self::WopiCatalog),
            _ => invalid(format!("invalid web extension storeType '{value}'")),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    pub(in crate::web) id: String,
    pub(in crate::web) version: String,
    pub(in crate::web) location: Option<String>,
    pub(in crate::web) store: Store,
    pub(in crate::web) extension_list: Option<ExtList>,
}

impl Reference {
    /// Create a validated reference for a catalog-backed provider.
    ///
    /// File-system references require a location and must be created with
    /// [`Self::file`]. Keeping the location in the constructor prevents the
    /// safe model from representing the store-less form rejected by Office.
    pub fn new(id: impl Into<String>, version: impl Into<String>, store: Store) -> Result<Self> {
        if store == Store::FileSystem {
            return invalid(
                "FileSystem references require Reference::file(id, version, location)".into(),
            );
        }
        let value = Self {
            id: id.into(),
            version: version.into(),
            location: None,
            store,
            extension_list: None,
        };
        validate_store_reference(&value)?;
        Ok(value)
    }

    /// Create a validated file-system reference with its required location.
    pub fn file(
        id: impl Into<String>,
        version: impl Into<String>,
        location: impl Into<String>,
    ) -> Result<Self> {
        let value = Self {
            id: id.into(),
            version: version.into(),
            location: Some(location.into()),
            store: Store::FileSystem,
            extension_list: None,
        };
        validate_store_reference(&value)?;
        Ok(value)
    }

    /// Add or replace the optional provider-specific location.
    pub fn location(mut self, value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        require_nonempty("reference location", &value)?;
        self.location = Some(value);
        Ok(self)
    }

    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    #[must_use]
    pub fn version(&self) -> &str {
        &self.version
    }

    #[must_use]
    pub const fn store(&self) -> Store {
        self.store
    }

    #[must_use]
    pub fn location_name(&self) -> Option<&str> {
        self.location.as_deref()
    }

    #[must_use]
    pub const fn ext(&self) -> Option<&ExtList> {
        self.extension_list.as_ref()
    }

    pub fn set_ext(&mut self, extension: ExtList) -> Result<&mut Self> {
        validate_extension_list(Some(&extension), &[ExtKind::AddIn])?;
        self.extension_list = Some(extension);
        Ok(self)
    }

    pub fn with_ext(mut self, extension: ExtList) -> Result<Self> {
        self.set_ext(extension)?;
        Ok(self)
    }

    pub fn clear_ext(&mut self) -> Option<ExtList> {
        self.extension_list.take()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Property {
    pub(in crate::web) name: String,
    pub(in crate::web) value: String,
}

impl Property {
    pub fn new(name: impl Into<String>, value: impl Into<String>) -> Result<Self> {
        let value = Self {
            name: name.into(),
            value: value.into(),
        };
        require_nonempty("property name", &value.name)?;
        Ok(value)
    }

    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// Binding data shape with forward-compatible retention of newer values.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BindingKind {
    Matrix,
    Table,
    Text,
    Other(String),
}

impl BindingKind {
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Matrix => "matrix",
            Self::Table => "table",
            Self::Text => "text",
            Self::Other(value) => value,
        }
    }

    pub(in crate::web) fn parse(value: &str) -> Result<Self> {
        require_nonempty("binding type", value)?;
        Ok(match value {
            "matrix" => Self::Matrix,
            "table" => Self::Table,
            "text" => Self::Text,
            value => Self::Other(value.to_owned()),
        })
    }
}

impl AsRef<str> for BindingKind {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Binding {
    pub(in crate::web) id: String,
    pub(in crate::web) kind: BindingKind,
    pub(in crate::web) app_ref: String,
    pub(in crate::web) extension_list: Option<ExtList>,
}

impl Binding {
    pub fn new(
        id: impl Into<String>,
        kind: impl AsRef<str>,
        app_ref: impl Into<String>,
    ) -> Result<Self> {
        let value = Self {
            id: id.into(),
            kind: BindingKind::parse(kind.as_ref())?,
            app_ref: app_ref.into(),
            extension_list: None,
        };
        validate_binding(&value)?;
        Ok(value)
    }

    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    #[must_use]
    pub const fn kind(&self) -> &BindingKind {
        &self.kind
    }

    #[must_use]
    pub fn kind_name(&self) -> &str {
        self.kind.as_str()
    }

    #[must_use]
    pub fn app_ref(&self) -> &str {
        &self.app_ref
    }

    #[must_use]
    pub const fn ext(&self) -> Option<&ExtList> {
        self.extension_list.as_ref()
    }

    pub fn set_ext(&mut self, extension: ExtList) -> Result<&mut Self> {
        validate_extension_list(Some(&extension), &[ExtKind::AddIn])?;
        self.extension_list = Some(extension);
        Ok(self)
    }

    pub fn with_ext(mut self, extension: ExtList) -> Result<Self> {
        self.set_ext(extension)?;
        Ok(self)
    }

    pub fn clear_ext(&mut self) -> Option<ExtList> {
        self.extension_list.take()
    }
}
