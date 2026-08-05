//! Typed, inert ODF spreadsheet content-validation metadata.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::parse_content_validations;
pub use model::{
    ContentValidation, ContentValidationPart, ContentValidations, ValidationCellAddress,
    ValidationCondition, ValidationDisplayList, ValidationEventListeners, ValidationFailure,
    ValidationMessage, ValidationMessageType, ValidationParagraph,
};
