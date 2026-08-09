//! Inert `OpenDocument` package digital signatures and trust-neutral verification.

mod crypto;
mod model;

#[allow(
    clippy::module_name_repetitions,
    reason = "These established public API names accurately distinguish digital-signature records."
)]
pub use crypto::{
    CanonicalizationAlgorithm, DocumentSigner, SignatureAlgorithm, SignatureValidity,
    SignatureVerification,
};
#[allow(
    clippy::module_name_repetitions,
    reason = "These established public API names accurately distinguish digital-signature records."
)]
pub use model::{DigitalSignature, DigitalSignatures, SignatureReference};

pub(crate) use crypto::{sign_package, verify_package};
#[allow(
    clippy::module_name_repetitions,
    reason = "The package paths and parser are explicitly scoped to signature handling."
)]
pub(crate) use model::{DOCUMENT_SIGNATURE_PATH, MACRO_SIGNATURE_PATH, parse_signature_container};
