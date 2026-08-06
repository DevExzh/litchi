//! XLSB formula-text compiler facade.

mod ast;
mod builtin;
mod codec;
mod compiler;
mod model;
mod parser;
mod references;
mod validation;

pub use compiler::Compiler;
pub(crate) use model::{CompilationContext, DefinedName};
pub(crate) use validation::{FORMULA_ERRORS, excel_name_eq};

#[cfg(test)]
pub(super) use builtin::has_builtin_function;
