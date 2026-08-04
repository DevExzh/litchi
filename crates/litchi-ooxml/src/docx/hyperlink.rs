//! Compatibility exports for the canonical WordprocessingML hyperlink owner.
//!
//! The value model and XML extraction live in [`litchi_docx`]. The DOCX host
//! retains this module so existing `litchi_ooxml::docx` paths remain stable.

pub use litchi_docx::hyperlink::Hyperlink;
