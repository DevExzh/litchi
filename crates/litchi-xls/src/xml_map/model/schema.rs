//! Schema identity and namespace preservation.

use super::identity::{SchemaId, validate_string};
use super::opaque::OpaqueXml;
use crate::{Error, Result};

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct NamespaceDeclaration {
    prefix: String,
    uri: String,
}

impl NamespaceDeclaration {
    pub fn try_new(prefix: impl Into<String>, uri: impl Into<String>) -> Result<Self> {
        let prefix = prefix.into();
        if !prefix.is_empty() && !valid_prefix(&prefix) {
            return Err(invalid("namespace declaration has an invalid prefix"));
        }
        let uri = validate_string(uri.into(), 65_535, "namespace URI", true)?;
        if !prefix.is_empty() && uri.is_empty() {
            return Err(invalid("prefixed namespace declaration has an empty URI"));
        }
        Ok(Self { prefix, uri })
    }

    pub fn prefix(&self) -> &str {
        &self.prefix
    }

    pub fn uri(&self) -> &str {
        &self.uri
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Schema {
    pub(super) id: SchemaId,
    pub(super) schema_ref: Option<Vec<SchemaId>>,
    pub(super) namespace: Option<String>,
    pub(super) namespaces: Vec<NamespaceDeclaration>,
    pub(super) payload: OpaqueXml,
}

impl Schema {
    pub fn try_new(id: SchemaId, payload: OpaqueXml) -> Result<Self> {
        Self::from_parts(id.as_str().to_string(), None, None, Vec::new(), payload)
    }

    pub(crate) fn from_parts(
        id: String,
        schema_ref: Option<String>,
        namespace: Option<String>,
        namespaces: Vec<(String, String)>,
        payload: OpaqueXml,
    ) -> Result<Self> {
        let id = SchemaId::new(id)?;
        let schema_ref = schema_ref
            .map(|value| parse_schema_refs(&value))
            .transpose()?;
        let namespace = namespace
            .map(|value| validate_string(value, 65_535, "schema Namespace", true))
            .transpose()?;
        let namespaces = namespaces
            .into_iter()
            .map(|(prefix, uri)| NamespaceDeclaration::try_new(prefix, uri))
            .collect::<Result<Vec<_>>>()?;
        Ok(Self {
            id,
            schema_ref,
            namespace,
            namespaces,
            payload,
        })
    }

    pub fn id(&self) -> &SchemaId {
        &self.id
    }

    pub fn schema_references(&self) -> Option<&[SchemaId]> {
        self.schema_ref.as_deref()
    }

    pub fn namespace(&self) -> Option<&str> {
        self.namespace.as_deref()
    }

    pub fn namespaces(&self) -> &[NamespaceDeclaration] {
        &self.namespaces
    }

    pub fn payload(&self) -> &OpaqueXml {
        &self.payload
    }
}

fn parse_schema_refs(value: &str) -> Result<Vec<SchemaId>> {
    value
        .split_whitespace()
        .map(|item| SchemaId::new(item.to_string()))
        .collect()
}

fn valid_prefix(value: &str) -> bool {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    (first.is_ascii_alphabetic() || first == '_')
        && chars.all(|character| {
            character.is_ascii_alphanumeric() || matches!(character, '_' | '-' | '.')
        })
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidData(format!("XML map: {}", message.into()))
}
