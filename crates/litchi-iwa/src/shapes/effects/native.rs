//! Native protobuf conversion for shape effects.

use crate::Result;
use crate::protobuf::tsd;

use litchi_iwa_common::shape::effects::{Opacity, Reflection, ReflectionOpacity};

pub(super) fn opacity_from_native(value: f32) -> Result<Opacity> {
    Ok(Opacity::new(value)?)
}

pub(super) const fn opacity_to_native(opacity: Opacity) -> f32 {
    opacity.value()
}

pub(super) fn reflection_from_native(native: &tsd::ReflectionArchive) -> Result<Reflection> {
    native.opacity.map_or(Ok(Reflection::Disabled), |opacity| {
        Ok(Reflection::Enabled(ReflectionOpacity::new(opacity)?))
    })
}

pub(super) fn reflection_to_native(reflection: Reflection) -> tsd::ReflectionArchive {
    tsd::ReflectionArchive {
        opacity: match reflection {
            Reflection::Disabled => None,
            Reflection::Enabled(opacity) => Some(opacity.value()),
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_reflection_distinguishes_disabled_and_enabled() {
        assert_eq!(
            reflection_from_native(&reflection_to_native(Reflection::Disabled)).unwrap(),
            Reflection::Disabled
        );
        let enabled = Reflection::Enabled(ReflectionOpacity::new(0.4).unwrap());
        assert_eq!(
            reflection_from_native(&reflection_to_native(enabled)).unwrap(),
            enabled
        );
    }
}
