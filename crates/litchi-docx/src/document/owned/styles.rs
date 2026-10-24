//! Style definitions staged with a main-document edit (record 0319).
//!
//! Each verb reads the staged styles part through
//! [`crate::styles::catalog::Catalog`], splices one `w:style` and stages the
//! part's exact replacement as a graph change (0251b); a document without a
//! styles part gains `styles.xml` and its relationship. Every other byte of
//! the part is kept as authored.

use std::sync::Arc;

use super::super::{Error, Refusal, Result};
use super::Edit;
use crate::Package;
use crate::styles::catalog::{Catalog, NewStyle, StyleChange};

const STYLES_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const STRICT_STYLES_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/styles";
const STYLES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD_NAMESPACE: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

impl Edit<Package> {
    /// The styles part as the staged changes leave it, or `None` when the
    /// document has none.
    ///
    /// # Errors
    ///
    /// Returns a package-graph error, or a malformed-part error.
    pub fn style_catalog(&self) -> Result<Option<Catalog>> {
        let Some((name, _)) = self.related_part(styles_relationship(self.is_strict()?))? else {
            return Ok(None);
        };
        let Some(payload) = self.part_payload(&name)? else {
            return Ok(None);
        };
        Ok(Some(Catalog::from_xml(payload).map_err(Error::Document)?))
    }

    /// Add `style` to the styles part, creating the part when the document
    /// has none.
    ///
    /// ```
    /// use litchi_docx::styles::Type;
    /// use litchi_docx::styles::catalog::{NewStyle, StyleChange};
    ///
    /// let package = litchi_docx::Builder::new()?.build()?;
    /// let mut edit = package.edit()?;
    /// edit.add_style(
    ///     &NewStyle::new("Accent", Type::Paragraph)
    ///         .with_change(StyleChange::default().with_name("Accent")),
    /// )?;
    /// let next = edit.commit()?.into_package();
    /// let catalog = next.style_catalog()?.unwrap();
    /// assert_eq!(catalog.get("Accent").unwrap().name(), Some("Accent"));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidPropertyValue`] for an empty or taken
    /// identifier, or a typed error for a malformed part.
    pub fn add_style(&mut self, style: &NewStyle) -> Result<&mut Self> {
        let strict = self.is_strict()?;
        if style.id().is_empty() {
            return Err(invalid_value());
        }
        match self.style_part()? {
            Some((name, catalog)) => {
                if catalog.get(style.id()).is_some() {
                    return Err(invalid_value());
                }
                let updated = catalog.with_added(style).map_err(Error::Document)?;
                self.replace_part(&name, updated)?;
            },
            None => {
                let namespace = if strict {
                    STRICT_WORD_NAMESPACE
                } else {
                    WORD_NAMESPACE
                };
                let empty = format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{namespace}"></w:styles>"#
                );
                let catalog =
                    Catalog::from_xml(Arc::new(empty.into_bytes())).map_err(Error::Document)?;
                let xml = catalog.with_added(style).map_err(Error::Document)?;
                self.stage_related_part(
                    styles_relationship(strict),
                    ("", "styles", ".xml"),
                    true,
                    STYLES_CONTENT_TYPE,
                    xml,
                )?;
            },
        }
        Ok(self)
    }

    /// Apply `change` to the style `id`. An empty change, or one that
    /// leaves the style as it reads, stages nothing.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown style, a typed
    /// refusal for a property change that cannot be written, or a typed
    /// error for a malformed part.
    pub fn set_style(&mut self, id: &str, change: &StyleChange) -> Result<&mut Self> {
        let (name, catalog) = self.style_part()?.ok_or_else(not_found)?;
        if catalog.get(id).is_none() {
            return Err(not_found());
        }
        if let Some(updated) = catalog.with_changed(id, change).map_err(Error::Document)? {
            self.replace_part(&name, updated)?;
        }
        Ok(self)
    }

    /// Remove the style `id` from the styles part. References to it in the
    /// document are kept; Word reads them as the type's default style.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown style, or a typed
    /// error for a malformed part.
    pub fn remove_style(&mut self, id: &str) -> Result<&mut Self> {
        let (name, catalog) = self.style_part()?.ok_or_else(not_found)?;
        if catalog.get(id).is_none() {
            return Err(not_found());
        }
        let updated = catalog.with_removed(id).map_err(Error::Document)?;
        self.replace_part(&name, updated)?;
        Ok(self)
    }

    /// The styles part's name and catalog, staged changes included.
    fn style_part(&self) -> Result<Option<(String, Catalog)>> {
        let Some((name, _)) = self.related_part(styles_relationship(self.is_strict()?))? else {
            return Ok(None);
        };
        let payload = self
            .part_payload(&name)?
            .ok_or_else(|| crate::Error::PartNotFound(name.clone()))?;
        Ok(Some((
            name,
            Catalog::from_xml(payload).map_err(Error::Document)?,
        )))
    }
}

const fn styles_relationship(strict: bool) -> &'static str {
    if strict {
        STRICT_STYLES_RELATIONSHIP
    } else {
        STYLES_RELATIONSHIP
    }
}

const fn invalid_value() -> Error {
    Error::Refused {
        position: 0,
        reason: Refusal::InvalidPropertyValue,
    }
}

const fn not_found() -> Error {
    Error::Refused {
        position: 0,
        reason: Refusal::BlockNotFound,
    }
}
