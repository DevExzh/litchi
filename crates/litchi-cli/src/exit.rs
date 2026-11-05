//! Process exit statuses: one per family of `litchi_ops::Code`, the same
//! families as the Python exception classes
//! (`crates/pyo3-litchi/src/errors.rs`). The crate documentation lists the
//! codes of each status.

use litchi_ops::Code;

/// Success.
pub(crate) const OK: u8 = 0;
/// A failure outside the named families.
pub(crate) const OTHER: u8 = 1;
/// The argument parser refused the command line.
pub(crate) const USAGE: u8 = 2;
/// A path or selector failed to parse or resolve.
pub(crate) const PATH: u8 = 3;
/// A property, value or argument was refused.
pub(crate) const PROPERTY: u8 = 4;
/// The format cannot perform the request.
pub(crate) const UNSUPPORTED: u8 = 5;
/// Reading or writing a file failed.
pub(crate) const IO: u8 = 6;
/// The input document is malformed.
pub(crate) const CORRUPT: u8 = 7;
/// The document is encrypted, signed or edit-protected.
pub(crate) const PROTECTED: u8 = 8;
/// A best-effort batch had failed items, or a `--unsupported warn` write
/// skipped properties.
pub(crate) const PARTIAL: u8 = 9;

/// The exit status of a failure with `code`.
pub(crate) const fn of(code: Code) -> u8 {
    match code {
        Code::NotFound
        | Code::PathNotFound
        | Code::ZeroMatches
        | Code::OutOfRange
        | Code::RangeTargetNotFound
        | Code::InvalidPath
        | Code::InvalidSelector
        | Code::InvalidXpath
        | Code::UnsupportedPath
        | Code::BareSelectorRejected => PATH,
        Code::InvalidValue
        | Code::UnsupportedProperty
        | Code::MissingProperty
        | Code::DuplicateName
        | Code::InvalidArgument
        | Code::MissingArgument
        | Code::InvalidJson
        | Code::CorruptBatch => PROPERTY,
        Code::UnsupportedFormat
        | Code::UnsupportedCapability
        | Code::UnsupportedElement
        | Code::UnsupportedType
        | Code::UnsupportedCommand => UNSUPPORTED,
        Code::FileNotFound
        | Code::FileLocked
        | Code::FileExists
        | Code::FileRequired
        | Code::IoError => IO,
        Code::CorruptFile | Code::InvalidInput => CORRUPT,
        Code::DocumentProtected => PROTECTED,
        _ => OTHER,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_code_has_a_documented_status() {
        for &code in Code::ALL {
            let status = of(code);
            assert!(
                matches!(status, OTHER | PATH..=PROTECTED),
                "{code} maps to {status}"
            );
        }
        assert_eq!(of(Code::PathNotFound), PATH);
        assert_eq!(of(Code::UnsupportedProperty), PROPERTY);
        assert_eq!(of(Code::UnsupportedCapability), UNSUPPORTED);
        assert_eq!(of(Code::FileNotFound), IO);
        assert_eq!(of(Code::CorruptFile), CORRUPT);
        assert_eq!(of(Code::DocumentProtected), PROTECTED);
        assert_eq!(of(Code::InternalError), OTHER);
        assert_ne!(PARTIAL, USAGE);
    }
}
