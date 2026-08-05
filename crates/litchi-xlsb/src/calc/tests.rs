use std::mem::size_of;

use crate::raw::Writer;

use super::{LEN, Opts, Props, read, write};

#[test]
fn defaults_round_trip_exactly() {
    let expected = Props::default();
    let mut payload = Vec::new();
    write(&expected, &mut Writer::new(&mut payload)).unwrap();

    assert_eq!(payload.len(), LEN);
    assert_eq!(read(&payload).unwrap(), expected);
}

#[test]
fn representation_has_small_structural_bounds() {
    assert_eq!(size_of::<Opts>(), 2);
    assert!(size_of::<Props>() <= 32);
}
