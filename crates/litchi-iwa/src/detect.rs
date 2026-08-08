//! Facade adapters for the independent iWork detector.
//!
//! The detector implementation and its typed error live in
//! the litchi-iwa-detect leaf. This module preserves the established
//! litchi_iwa::detect API and maps leaf failures into the facade error.

use crate::application::Application;
use std::{
    io::{Read, Seek},
    path::Path,
};

pub use litchi_iwa_detect::{Format, Limits};

impl From<litchi_iwa_detect::Error> for crate::Error {
    fn from(error: litchi_iwa_detect::Error) -> Self {
        match error {
            litchi_iwa_detect::Error::Io(error) => Self::Io(error),
            litchi_iwa_detect::Error::IwaCore(error) => Self::from(error),
            litchi_iwa_detect::Error::IwaCommon(error) => Self::IwaCommon(error),
            litchi_iwa_detect::Error::InvalidFormat(message) => Self::InvalidFormat(message),
            litchi_iwa_detect::Error::Archive(message) => {
                Self::Archive(format!("iWork archive ingress: {message}"))
            },
            litchi_iwa_detect::Error::LimitExceeded {
                kind,
                observed,
                maximum,
            } => Self::Bundle(format!(
                "iWork detection {kind} limit exceeded: observed {observed}, maximum {maximum}"
            )),
            litchi_iwa_detect::Error::InvalidLimits => {
                Self::Bundle("invalid iWork detection limits".to_owned())
            },
            litchi_iwa_detect::Error::Allocation { amount } => Self::Bundle(format!(
                "iWork detection allocation failed for {amount} units"
            )),
            litchi_iwa_detect::Error::Encrypted => Self::InvalidFormat(
                "password-protected iWork documents are not supported".to_owned(),
            ),
            litchi_iwa_detect::Error::SourceChanged => {
                Self::Bundle("iWork source changed while it was being captured".to_owned())
            },
            _ => Self::Bundle("unrecognized iWork detection failure".to_owned()),
        }
    }
}

/// Detect an iWork application from complete packaged bytes.
pub fn bytes(value: &[u8]) -> crate::Result<Option<Format>> {
    litchi_iwa_detect::bytes(value).map_err(Into::into)
}

/// Detect an iWork application using caller-selected resource ceilings.
pub fn bytes_with_limits(value: &[u8], limits: Limits) -> crate::Result<Option<Format>> {
    litchi_iwa_detect::bytes_with_limits(value, limits).map_err(Into::into)
}

/// Detect an iWork application from a seekable stream.
pub fn reader<R: Read + Seek>(value: &mut R) -> crate::Result<Option<Format>> {
    litchi_iwa_detect::reader(value).map_err(Into::into)
}

/// Detect an iWork application from a seekable stream under explicit limits.
pub fn reader_with_limits<R: Read + Seek>(
    value: &mut R,
    limits: Limits,
) -> crate::Result<Option<Format>> {
    litchi_iwa_detect::reader_with_limits(value, limits).map_err(Into::into)
}

/// Detect a packaged iWork file or a legacy directory bundle.
pub fn path(value: impl AsRef<Path>) -> crate::Result<Option<Format>> {
    litchi_iwa_detect::path(value).map_err(Into::into)
}

/// Detect a packaged file or legacy directory bundle under explicit limits.
pub fn path_with_limits(value: impl AsRef<Path>, limits: Limits) -> crate::Result<Option<Format>> {
    litchi_iwa_detect::path_with_limits(value, limits).map_err(Into::into)
}

pub(crate) fn detect_application_from_document(payload: &[u8]) -> Option<Application> {
    litchi_iwa_detect::detect_application_from_document(payload).map(|format| match format {
        Format::Pages => Application::Pages,
        Format::Keynote => Application::Keynote,
        Format::Numbers => Application::Numbers,
    })
}

#[cfg(test)]
mod tests {
    #[test]
    fn expanded_detector_failures_remain_compatible_with_the_migration_host() {
        let limit = crate::Error::from(litchi_iwa_detect::Error::LimitExceeded {
            kind: litchi_iwa_detect::LimitKind::EntryBytes,
            observed: 9,
            maximum: 8,
        });
        assert!(
            matches!(limit, crate::Error::Bundle(message) if message.contains("9") && message.contains("8"))
        );

        let encrypted = crate::Error::from(litchi_iwa_detect::Error::Encrypted);
        assert!(matches!(encrypted, crate::Error::InvalidFormat(_)));

        let changed = crate::Error::from(litchi_iwa_detect::Error::SourceChanged);
        assert!(matches!(changed, crate::Error::Bundle(message) if message.contains("changed")));
    }
}
