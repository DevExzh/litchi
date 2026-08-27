//! Package-error mapping for the standalone XLSB formula codec.

use crate::package::error::Error;

impl From<crate::formula::Error> for Error {
    fn from(error: crate::formula::Error) -> Self {
        match error {
            crate::formula::Error::InvalidFormula(message) => Self::InvalidFormula(message),
            crate::formula::Error::InvalidCellReference(reference) => {
                Self::InvalidCellReference(reference)
            },
            crate::formula::Error::InvalidLength { expected, found } => {
                Self::InvalidLength { expected, found }
            },
            crate::formula::Error::UnsupportedFeature(feature) => Self::UnsupportedFeature(feature),
            crate::formula::Error::UnresolvedDependency(dependency) => {
                Self::UnresolvedDependency(dependency)
            },
            crate::formula::Error::Encoding(message) => Self::Encoding(message),
        }
    }
}
