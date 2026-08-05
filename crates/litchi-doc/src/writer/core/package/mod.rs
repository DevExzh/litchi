//! FIB-backed DOC stream assembly and compound-file packaging facade.
//
// The implementation is divided by responsibility: semantic story/table
// construction lives in the semantic module, while stream finalization and
// OLE2 packaging live in the package module.

mod package;
mod semantic;
