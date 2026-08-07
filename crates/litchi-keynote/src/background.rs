//! Archive-free Keynote slide-background values.

use crate::{Error, Result};

pub use litchi_iwa_common::shape::fill::{Angle, Gradient, Kind, Stop};

/// The semantic fill of a Keynote slide background.
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq)]
pub enum Background {
    /// No fill is applied.
    None,
    /// A single validated color fills the slide.
    Solid(litchi_iwa_common::color::Rgba),
    /// A validated native gradient fills the slide.
    Gradient(Gradient),
    /// A bounded native fill payload unknown to this version of the crate.
    Opaque(Opaque),
}

/// A lossless, non-empty native fill payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Opaque(Box<[u8]>);

impl Opaque {
    /// Retain a non-empty native payload without interpreting its fields.
    ///
    /// # Errors
    ///
    /// Returns [`Error::EmptyBackgroundPayload`] when `payload` is empty.
    pub fn new(payload: impl Into<Box<[u8]>>) -> Result<Self> {
        let bytes = payload.into();
        if bytes.is_empty() {
            return Err(Error::EmptyBackgroundPayload);
        }
        Ok(Self(bytes))
    }

    /// Copy a borrowed native payload into an opaque semantic value.
    ///
    /// # Errors
    ///
    /// Returns [`Error::EmptyBackgroundPayload`] when `payload` is empty.
    pub fn from_slice(payload: &[u8]) -> Result<Self> {
        Self::new(payload.to_vec().into_boxed_slice())
    }

    /// Borrow the exact native bytes retained by this value.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.0
    }

    /// Consume the value and return its exact native bytes.
    #[must_use]
    pub fn into_bytes(self) -> Box<[u8]> {
        self.0
    }
}

#[cfg(test)]
mod tests {
    use super::{Background, Error, Opaque};

    #[test]
    fn opaque_payload_is_lossless_and_non_empty() -> Result<(), Error> {
        let opaque = Opaque::from_slice(&[0x0a, 0xff])?;
        assert_eq!(opaque.as_bytes(), [0x0a, 0xff]);
        assert_eq!(
            Background::Opaque(opaque.clone()),
            Background::Opaque(opaque)
        );
        assert_eq!(
            Opaque::from_slice(&[]).err(),
            Some(Error::EmptyBackgroundPayload)
        );
        Ok(())
    }
}
