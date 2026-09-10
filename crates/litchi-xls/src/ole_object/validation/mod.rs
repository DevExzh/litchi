//! Validation layer for authored and parsed Obj/control metadata.

mod controls;
mod obj;

pub(crate) use obj::validate_picture_formula;
