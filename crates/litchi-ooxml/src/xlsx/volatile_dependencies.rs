//! Compatibility access to the canonical XLSX volatile-dependencies owner.

pub use litchi_xlsx::volatile_dependencies::{
    VolatileDependencies, VolatileDependenciesConformance, VolatileDependencyType, VolatileMain,
    VolatileReference, VolatileTopic, VolatileType, VolatileValue, load_from_package,
    load_from_package_with_conformance, remove_from_package, store_in_package,
};
