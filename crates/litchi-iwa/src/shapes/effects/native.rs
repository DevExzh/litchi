//! Native protobuf conversion for shape effects.

use crate::Result;
use crate::protobuf::tsd;

use super::{ShapeOpacity, ShapeReflection, ShapeReflectionOpacity};

pub(super) fn opacity_from_native(value: f32) -> Result<ShapeOpacity> {
    ShapeOpacity::new(value)
}

pub(super) const fn opacity_to_native(opacity: ShapeOpacity) -> f32 {
    opacity.get()
}

pub(super) fn reflection_from_native(native: &tsd::ReflectionArchive) -> Result<ShapeReflection> {
    native
        .opacity
        .map_or(Ok(ShapeReflection::Disabled), |opacity| {
            Ok(ShapeReflection::Enabled(ShapeReflectionOpacity::new(
                opacity,
            )?))
        })
}

pub(super) fn reflection_to_native(reflection: ShapeReflection) -> tsd::ReflectionArchive {
    tsd::ReflectionArchive {
        opacity: match reflection {
            ShapeReflection::Disabled => None,
            ShapeReflection::Enabled(opacity) => Some(opacity.get()),
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_reflection_distinguishes_disabled_and_enabled() {
        assert_eq!(
            reflection_from_native(&reflection_to_native(ShapeReflection::Disabled)).unwrap(),
            ShapeReflection::Disabled
        );
        let enabled = ShapeReflection::Enabled(ShapeReflectionOpacity::new(0.4).unwrap());
        assert_eq!(
            reflection_from_native(&reflection_to_native(enabled)).unwrap(),
            enabled
        );
    }
}
