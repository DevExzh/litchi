//! Shape text extraction belongs to the document-level text traversal.
//!
//! The former private extractor only returned shape accessibility metadata and
//! duplicated archive decoding without reaching referenced text storage. The
//! format document readers now use the canonical text traversal instead.
