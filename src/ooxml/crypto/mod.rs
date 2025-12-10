pub mod agile;
mod ole_encrypted_package;
pub mod standard2007;

pub use agile::encrypt_ooxml_package_agile;
pub use standard2007::encrypt_ooxml_package_standard_2007;

/// OOXML encryption helpers (Standard 2007, Agile).
///
/// This module is compiled only when the `ooxml_encryption` feature is enabled.
/// The public APIs here are intentionally minimal and format-agnostic so they can
/// be reused by DOCX, PPTX, and XLSX packages.
fn password_to_utf16le(password: &str) -> Vec<u8> {
    let mut buf = Vec::with_capacity(password.len() * 2);
    for ch in password.encode_utf16() {
        let bytes = ch.to_le_bytes();
        buf.push(bytes[0]);
        buf.push(bytes[1]);
    }
    buf
}
