//! Strict, inert parsing of `PowerPoint` 9 `EnvelopeData9Atom` records.
//!
//! The known Office mail-envelope CLSID is decoded according to MS-OSHARED.
//! Other CLSIDs remain bounded opaque payloads. Nothing in this module sends
//! mail, opens attachments, invokes a mail client, or evaluates embedded data.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{
    EnvelopeData, EnvelopePayload, MSO_ENVELOPE_CLSID, MsoAttachment, MsoEnvelope, MsoEnvelopeText,
    MsoEnvelopeVersion, MsoFollowUpStatus, MsoImportance, MsoPropertyValue, MsoRecipientCollection,
    MsoRecipientProperties, MsoRecipientProperty, MsoSecurityFlags, MsoSensitivity,
};
