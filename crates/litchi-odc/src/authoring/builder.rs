//! Chart package authoring.

use super::{ChartClass, Definition, serialize_content_with_limits};
use litchi_core::{Error, Result};
use litchi_odf_common::{compact_xml, core::PackageWriter};

#[derive(Clone, Debug)]
struct Resource {
    path: String,
    media_type: String,
    bytes: Vec<u8>,
}

/// Detached typed builder for a standalone chart package.
#[derive(Clone, Debug)]
pub struct Builder {
    definition: Definition,
    limits: crate::Limits,
    styles_xml: Option<String>,
    resources: Vec<Resource>,
}

impl Builder {
    #[must_use]
    pub fn new() -> Self {
        Self {
            definition: Definition::new(ChartClass::line()),
            limits: crate::Limits::default(),
            styles_xml: None,
            resources: Vec::new(),
        }
    }

    /// Supply the typed chart definition to publish.
    #[must_use]
    pub fn with_definition(mut self, definition: Definition) -> Self {
        self.definition = definition;
        self
    }

    #[must_use]
    pub fn definition(&self) -> &Definition {
        &self.definition
    }

    pub fn definition_mut(&mut self) -> &mut Definition {
        &mut self.definition
    }

    /// Retain caller-selected limits for validation and package publication.
    #[must_use]
    pub fn with_limits(mut self, limits: crate::Limits) -> Self {
        self.limits = limits;
        self
    }

    /// Add or replace a fresh package `styles.xml` payload.
    #[must_use]
    pub fn with_styles_xml(mut self, styles_xml: impl Into<String>) -> Self {
        self.styles_xml = Some(styles_xml.into());
        self
    }

    /// Add one inert package-local resource with an explicit manifest type.
    #[must_use]
    pub fn with_resource(
        mut self,
        path: impl Into<String>,
        media_type: impl Into<String>,
        bytes: impl Into<Vec<u8>>,
    ) -> Self {
        self.resources.push(Resource {
            path: path.into(),
            media_type: media_type.into(),
            bytes: bytes.into(),
        });
        self
    }

    /// Serialize the definition into a validated chart package.
    ///
    /// # Errors
    ///
    /// Returns an error if the definition fails serialization or validation,
    /// or if the package cannot be written.
    pub fn build(self) -> Result<Vec<u8>> {
        if self.resources.len() > self.limits.max_resources() {
            return Err(Error::InvalidFormat(
                "ODC resource count exceeds the caller-selected limit".into(),
            ));
        }
        let content_xml = serialize_content_with_limits(&self.definition, self.limits)?;
        package_content(
            &content_xml,
            self.styles_xml.as_deref(),
            &self.resources,
            self.limits,
        )
    }
}

impl Default for Builder {
    fn default() -> Self {
        Self::new()
    }
}

fn package_content(
    content_xml: &str,
    styles_xml: Option<&str>,
    resources: &[Resource],
    limits: crate::Limits,
) -> Result<Vec<u8>> {
    let compact_limits = compact_xml::Limits::new(limits.max_content_bytes(), limits.max_depth())
        .map_err(Error::from)?;
    compact_xml::validate_with_limits(content_xml.as_bytes(), compact_limits)?;
    crate::codec::validate(content_xml)?;
    if let Some(styles) = styles_xml {
        compact_xml::validate_with_limits(styles.as_bytes(), compact_limits)?;
        crate::codec::validate_styles(styles, limits)?;
    }
    let mut writer = PackageWriter::new_bounded(limits.max_package_bytes());
    writer.set_mimetype(crate::package::MIMETYPE)?;
    writer.add_file("content.xml", content_xml.as_bytes())?;
    if let Some(styles) = styles_xml {
        writer.add_file("styles.xml", styles.as_bytes())?;
    }
    for resource in resources {
        validate_media_type(&resource.media_type)?;
        crate::package::validate_authored_resource(
            &resource.path,
            &resource.media_type,
            &resource.bytes,
            compact_limits,
        )?;
        writer.add_file_with_media_type(&resource.path, &resource.bytes, &resource.media_type)?;
    }
    let bytes = writer.finish_to_bounded_bytes()?;
    crate::package::Snapshot::from_bytes_with_limits(bytes, limits)
        .map(crate::package::Snapshot::into_bytes)
}

fn validate_media_type(media_type: &str) -> Result<()> {
    if media_type.is_empty()
        || media_type.len() > 1_024
        || !media_type.is_ascii()
        || !media_type.contains('/')
        || media_type
            .bytes()
            .any(|byte| byte.is_ascii_control() || byte.is_ascii_whitespace())
    {
        return Err(Error::InvalidFormat(
            "ODC resource media type is invalid".into(),
        ));
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "test code panics on unexpected errors to keep assertions concise"
)]
mod tests {
    use super::{ChartClass, Definition, package_content};
    use litchi_core::{Error, xml::CompactnessKind};

    #[test]
    fn noncompact_serialized_content_is_rejected_before_publication() {
        let content = crate::serialize_content(&Definition::new(ChartClass::line())).unwrap();
        let noncompact = content.replacen("><", ">\n<", 1);
        assert!(matches!(
            package_content(&noncompact, None, &[], crate::Limits::default()).unwrap_err(),
            Error::XmlCompactness {
                kind: CompactnessKind::FormattingWhitespace,
                ..
            }
        ));
    }
}
