//! Native lighting-package builders for the Chart inspector's fixed presets.

use super::Chart3dLightingStyle;
use crate::protobuf::tsch;

const BLACK: [f32; 3] = [0.0, 0.0, 0.0];
const WHITE: [f32; 3] = [1.0, 1.0, 1.0];
const HOMOGENEOUS_COLOR: f32 = 1.0;
const HOMOGENEOUS_DIRECTION: f32 = 0.0;
const ATTENUATION: [f32; 3] = [0.0, 0.0, 1.0];

#[derive(Clone, Copy)]
enum CoordinateSpace {
    Scene = 0,
    Camera = 1,
}

impl CoordinateSpace {
    const fn native(self) -> u32 {
        self as u32
    }
}

#[derive(Clone, Copy)]
enum LightSource {
    Point([f32; 3]),
    Directional([f32; 3]),
    Spot {
        position: [f32; 3],
        direction: [f32; 3],
        cutoff: f32,
        dropoff: f32,
    },
}

#[derive(Clone, Copy)]
struct LightSpec {
    name: &'static str,
    diffuse: [f32; 3],
    intensity: f32,
    coordinate_space: CoordinateSpace,
    source: LightSource,
}

impl LightSpec {
    const fn point(name: &'static str, intensity: f32, position: [f32; 3]) -> Self {
        Self {
            name,
            diffuse: WHITE,
            intensity,
            coordinate_space: CoordinateSpace::Camera,
            source: LightSource::Point(position),
        }
    }

    const fn scene_point(
        name: &'static str,
        diffuse: [f32; 3],
        intensity: f32,
        position: [f32; 3],
    ) -> Self {
        Self {
            name,
            diffuse,
            intensity,
            coordinate_space: CoordinateSpace::Scene,
            source: LightSource::Point(position),
        }
    }

    const fn colored_point(
        name: &'static str,
        diffuse: [f32; 3],
        intensity: f32,
        position: [f32; 3],
    ) -> Self {
        Self {
            name,
            diffuse,
            intensity,
            coordinate_space: CoordinateSpace::Camera,
            source: LightSource::Point(position),
        }
    }

    const fn directional(name: &'static str, intensity: f32, direction: [f32; 3]) -> Self {
        Self {
            name,
            diffuse: WHITE,
            intensity,
            coordinate_space: CoordinateSpace::Camera,
            source: LightSource::Directional(direction),
        }
    }

    const fn spot(
        name: &'static str,
        intensity: f32,
        position: [f32; 3],
        direction: [f32; 3],
        cutoff: f32,
        dropoff: f32,
    ) -> Self {
        Self {
            name,
            diffuse: WHITE,
            intensity,
            coordinate_space: CoordinateSpace::Camera,
            source: LightSource::Spot {
                position,
                direction,
                cutoff,
                dropoff,
            },
        }
    }
}

const SOFT_LIGHT: &[LightSpec] = &[
    LightSpec::point("Fill Center", 0.4, [11.0, 50.0, 100.0]),
    LightSpec::point("Directional Key", 1.0, [-50.0, 90.0, 100.0]),
    LightSpec::point("Fill Right", 0.4, [100.0, 0.0, 0.0]),
    LightSpec::point("Fill Left", 0.1, [-100.0, 0.0, 0.0]),
];

const SOFT_FILL_KEY_DIRECTION: [f32; 3] = [
    f32::from_bits(0x3f30_9fb8),
    f32::from_bits(0xbf2f_c828),
    f32::from_bits(0xbe6a_a197),
];
const SOFT_FILL_BOUNCE_COLOR: [f32; 3] = [0.625, 0.625, 0.625];
const SOFT_FILL: &[LightSpec] = &[
    LightSpec::directional("Key", 0.8, SOFT_FILL_KEY_DIRECTION),
    LightSpec::point("Center Fill", 0.5, [5.0, 10.0, 25.0]),
    LightSpec::point("Right Fill", 0.25, [100.0, 0.0, -100.0]),
    LightSpec::scene_point("Bounce", SOFT_FILL_BOUNCE_COLOR, 0.6, [-10.0, -10.0, 10.0]),
    LightSpec::point("Top", 0.7, [10.0, 50.0, 0.0]),
    LightSpec::point("Left Fill", 0.15, [-100.0, 0.0, -100.0]),
];

const MEDIUM_CENTER_SPOT_POSITION: [f32; 3] = [
    f32::from_bits(0x41f0_0101),
    f32::from_bits(0x4220_0091),
    f32::from_bits(0x41f0_00f4),
];
const MEDIUM_CENTER_SPOT_DIRECTION: [f32; 3] = [
    f32::from_bits(0x3ec6_5c31),
    f32::from_bits(0xbf14_2826),
    f32::from_bits(0xbf37_b56b),
];
const MEDIUM_CENTER_EDGE_COLOR: [f32; 3] = [
    f32::from_bits(0x3f66_0ee6),
    1.0,
    f32::from_bits(0x3f7f_f800),
];
const MEDIUM_CENTER: &[LightSpec] = &[
    LightSpec::point("Fill Center", 0.3, [-12.0, 12.0, 30.0]),
    LightSpec::spot(
        "Directional Key",
        0.9,
        MEDIUM_CENTER_SPOT_POSITION,
        MEDIUM_CENTER_SPOT_DIRECTION,
        3.2,
        0.0,
    ),
    LightSpec::colored_point(
        "Edge Right",
        MEDIUM_CENTER_EDGE_COLOR,
        0.2,
        [50.0, 20.0, 0.0],
    ),
    LightSpec::colored_point(
        "Edge Left",
        MEDIUM_CENTER_EDGE_COLOR,
        0.2,
        [-50.0, 20.0, 0.0],
    ),
];

const MEDIUM_SIDE_TOP_COLOR: [f32; 3] = [
    f32::from_bits(0x3f52_d1d3),
    f32::from_bits(0x3f52_d1d3),
    1.0,
];
const MEDIUM_SIDE_HIGH_COLOR: [f32; 3] = [
    f32::from_bits(0x3f74_c568),
    f32::from_bits(0x3f78_3cb4),
    1.0,
];
const MEDIUM_RIGHT: &[LightSpec] = &[
    LightSpec::colored_point("Top Light", MEDIUM_SIDE_TOP_COLOR, 0.1, [12.0, 100.0, 0.0]),
    LightSpec::colored_point(
        "Fill Center High",
        MEDIUM_SIDE_HIGH_COLOR,
        0.9,
        [70.0, 90.0, 100.0],
    ),
    LightSpec::point("Fill Left", 0.2, [-100.0, 0.0, 40.0]),
    LightSpec::point("Fill Center Low", 0.5, [70.0, -90.0, 100.0]),
];
const MEDIUM_LEFT: &[LightSpec] = &[
    LightSpec::colored_point("Top Light", MEDIUM_SIDE_TOP_COLOR, 0.1, [12.0, 100.0, 0.0]),
    LightSpec::colored_point(
        "Fill Center High",
        MEDIUM_SIDE_HIGH_COLOR,
        0.9,
        [-50.0, 90.0, 100.0],
    ),
    LightSpec::point("Fill Right", 0.2, [100.0, 0.0, 40.0]),
    LightSpec::point("Fill Center Low", 0.5, [-50.0, -90.0, 100.0]),
];

const GLOSSY_KEY_DIRECTION: [f32; 3] = [
    f32::from_bits(0x3ed6_5731),
    f32::from_bits(0xbea4_d916),
    f32::from_bits(0xbf59_62e6),
];
const GLOSSY: &[LightSpec] = &[
    LightSpec::directional("Directional Key", 0.7, GLOSSY_KEY_DIRECTION),
    LightSpec::point("Center", 0.3, [11.0, 0.0, 50.0]),
    LightSpec::point("Edge Fill", 0.1, [100.0, 0.0, -50.0]),
    LightSpec::point("Bounce", 0.2, [11.0, -100.0, 100.0]),
    LightSpec::point("Top", 0.3, [11.0, 50.0, 0.0]),
];

pub(super) fn native_lighting_package(
    style: Chart3dLightingStyle,
) -> Option<tsch::Chart3DLightingPackageArchive> {
    let lights = match style {
        Chart3dLightingStyle::Default => return None,
        Chart3dLightingStyle::SoftLight => SOFT_LIGHT,
        Chart3dLightingStyle::SoftFill => SOFT_FILL,
        Chart3dLightingStyle::MediumCenter => MEDIUM_CENTER,
        Chart3dLightingStyle::MediumRight => MEDIUM_RIGHT,
        Chart3dLightingStyle::MediumLeft => MEDIUM_LEFT,
        Chart3dLightingStyle::Glossy => GLOSSY,
    };
    Some(tsch::Chart3DLightingPackageArchive {
        name: style.native_name().to_owned(),
        lights: lights.iter().copied().map(native_light).collect(),
    })
}

fn native_light(spec: LightSpec) -> tsch::Chart3DLightArchive {
    let mut light = tsch::Chart3DLightArchive {
        name: spec.name.to_owned(),
        ambient_color: vector(BLACK, HOMOGENEOUS_COLOR),
        diffuse_color: vector(spec.diffuse, HOMOGENEOUS_COLOR),
        specular_color: vector(spec.diffuse, HOMOGENEOUS_COLOR),
        intensity: spec.intensity,
        attenuation: vector(ATTENUATION, HOMOGENEOUS_DIRECTION),
        coordinate_space: spec.coordinate_space.native(),
        enabled: true,
        point_light: None,
        directional_light: None,
        spot_light: None,
    };
    match spec.source {
        LightSource::Point(position) => {
            light.point_light = Some(tsch::Chart3DPointLightArchive {
                position: vector(position, HOMOGENEOUS_DIRECTION),
            });
        },
        LightSource::Directional(direction) => {
            light.directional_light = Some(tsch::Chart3DDirectionalLightArchive {
                direction: vector(direction, HOMOGENEOUS_DIRECTION),
            });
        },
        LightSource::Spot {
            position,
            direction,
            cutoff,
            dropoff,
        } => {
            light.spot_light = Some(tsch::Chart3DSpotLightArchive {
                position: vector(position, HOMOGENEOUS_DIRECTION),
                direction: vector(direction, HOMOGENEOUS_DIRECTION),
                cutoff,
                dropoff,
            });
        },
    }
    light
}

const fn vector(values: [f32; 3], w: f32) -> tsch::Chart3DVectorArchive {
    tsch::Chart3DVectorArchive {
        x: values[0],
        y: values[1],
        z: values[2],
        w,
    }
}
