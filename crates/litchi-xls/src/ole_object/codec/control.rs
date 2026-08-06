//! Control-payload codec boundary.
//!
//! The concrete payload routines currently remain in the Obj dispatcher so
//! one wire walk owns all unknown-subrecord preservation. This module exposes
//! the XLUnicodeString decoder to the semantic/control layer without making
//! controls executable or normalizing their bytes.
