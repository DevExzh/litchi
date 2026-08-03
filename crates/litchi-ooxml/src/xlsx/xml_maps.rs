//! Compatibility access to the canonical XLSX Custom XML Maps owner.

pub use litchi_xlsx::xml_maps::{
    XmlMap, XmlMapConformance, XmlMapDataBinding, XmlMapInfo, XmlMapSchema, load_from_package,
    load_from_package_with_conformance, remove_from_package, store_in_package,
};
