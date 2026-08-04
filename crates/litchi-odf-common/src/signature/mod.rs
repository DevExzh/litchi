//! Inert OpenDocument package signatures and trust-neutral verification.

mod crypto;
mod model;

pub use crypto::{
    CanonicalizationAlgorithm, DocumentSigner, SignatureAlgorithm, SignatureValidity,
    SignatureVerification,
};
pub use model::{DigitalSignature, DigitalSignatures, SignatureReference};

pub(crate) use crypto::{sign_package, verify_package};
pub(crate) use model::{DOCUMENT_SIGNATURE_PATH, MACRO_SIGNATURE_PATH, parse_signature_container};
