//! PPT's private compatibility facade for shared OfficeArt wire helpers.
//!
//! The binary vocabulary is owned by [`litchi_odraw`]. This module only keeps
//! the existing crate-private helper paths used by PPT fixtures and semantic
//! tests; no OfficeArt wire structs or constants are defined here.

pub(crate) use litchi_odraw::write::prop_value;

#[cfg(test)]
pub(crate) use litchi_odraw::write::{
    PropertyBuilder, ShapeBuilder, raw_atom as write_atom, raw_container as write_container,
    record_type,
};
#[cfg(test)]
pub(crate) use litchi_odraw::write::{child_anchor as write_child_anchor, spgr as write_spgr};

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn property_builder_distinguishes_blip_and_complex_flags() {
        let mut builder = PropertyBuilder::new();
        builder.add_blip_id(0x0104, 7);
        builder.add_complex(0x0145, &[1, 2, 3, 4]);

        let mut bytes = Vec::new();
        builder.write(&mut bytes).unwrap();

        assert_eq!(builder.size().unwrap(), 24);
        assert_eq!(u16::from_le_bytes([bytes[8], bytes[9]]), 0x4104);
        assert_eq!(i32::from_le_bytes(bytes[10..14].try_into().unwrap()), 7);
        assert_eq!(u16::from_le_bytes([bytes[14], bytes[15]]), 0x8145);
        assert_eq!(i32::from_le_bytes(bytes[16..20].try_into().unwrap()), 4);
        assert_eq!(&bytes[20..], &[1, 2, 3, 4]);
    }
}
