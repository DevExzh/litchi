//! Typed object opacity and reflection effects for ordinary shapes.

mod native;
mod style;

pub(crate) use style::{reset_shape_effects, set_shape_effects, shape_effects};

impl From<litchi_iwa_common::shape::effects::Error> for crate::Error {
    fn from(error: litchi_iwa_common::shape::effects::Error) -> Self {
        Self::ParseError(error.to_string())
    }
}
