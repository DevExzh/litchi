//! Layered, namespace-aware `OpenDocument` configuration settings.
//!
//! The owner facade exposes the typed configuration vocabulary while keeping
//! XML decoding, package/flat dispatch, and regression coverage in focused
//! modules.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{
    ConfigItem, ConfigMap, ConfigMapEntry, ConfigNode, ConfigSet, ConfigValue, Settings,
};

pub(crate) use package::{parse_flat, parse_package};
