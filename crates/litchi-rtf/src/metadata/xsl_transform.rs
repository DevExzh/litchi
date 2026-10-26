use std::borrow::Cow;

/// Maximum encoded UTF-8 size accepted for a custom XSL transform location.
pub const MAX_DOCUMENT_XSL_TRANSFORM_LOCATION_BYTES: usize = 65_536;

/// Passive requested intent represented by the RTF `usexform` flag.
///
/// This records source metadata only. It never causes a transform to be
/// resolved or executed, including when a transform location is present.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum DocumentXslTransformUsage {
    #[default]
    NotRequested,
    Requested,
}

impl DocumentXslTransformUsage {
    #[must_use]
    pub fn is_requested(self) -> bool {
        matches!(self, Self::Requested)
    }
}

/// Passive location metadata from the RTF `xform` destination.
///
/// This value is inert: parsing or setting it never resolves, opens, downloads,
/// or executes the referenced transform.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentXslTransform<'a> {
    pub location: Cow<'a, str>,
}

impl<'a> DocumentXslTransform<'a> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(location: Cow<'a, str>) -> crate::Result<Self> {
        let transform = Self { location };
        transform.validate()?;
        Ok(transform)
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> crate::Result<()> {
        if self.location.is_empty() {
            return Err(crate::Error::MalformedDocument(
                "RTF XSL transform location must not be empty".to_string(),
            ));
        }
        if self.location.len() > MAX_DOCUMENT_XSL_TRANSFORM_LOCATION_BYTES {
            return Err(crate::Error::MalformedDocument(
                "RTF XSL transform location exceeds the resource limit".to_string(),
            ));
        }
        if self
            .location
            .chars()
            .any(|character| matches!(character, '\0' | '\r' | '\n'))
        {
            return Err(crate::Error::MalformedDocument(
                "RTF XSL transform location contains a forbidden character".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> DocumentXslTransform<'static> {
        DocumentXslTransform {
            location: Cow::Owned(self.location.into_owned()),
        }
    }
}
