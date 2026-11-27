//! TxMasterStyleAtom builder (MS-PPT 2.9.45)
//!
//! Constructs text master style atoms with proper formatting structures
//! using zerocopy for binary serialization.

mod codec;
mod semantic;
#[cfg(test)]
mod validation;

pub use codec::*;
pub use semantic::*;
