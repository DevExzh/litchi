//! Temporary host exports for the canonical PPTX protection owner.

pub use litchi_pptx::presentation_properties::metadata::protection::{
    Algorithm as CryptoAlgorithm, Settings as Protection, Slide as SlideProtection,
    Type as ProtectionType, Verifier as ModifyVerifier,
};
