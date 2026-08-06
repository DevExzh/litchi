//! Package-facing emission for List12 and following worksheet records.
//!
//! List12 wire records, future-record placement, and placement validation are
//! kept separate while these inherent methods remain available through the
//! existing `ListObject` facade.

mod future;
mod records;
mod validation;
