//! Typed, bounded `MsoEnvelopeCLSID` metadata from a DOC table stream.

mod codec;
mod model;
mod validation;

pub use model::{
    Attachment, Envelope, FollowUpStatus, Importance, MSO_ENVELOPE_CLSID, Message, Payload,
    PropertyValue, RecipientCollection, RecipientProperties, RecipientProperty, SecurityFlags,
    Sensitivity, Text, Version,
};

#[cfg(test)]
mod tests;
