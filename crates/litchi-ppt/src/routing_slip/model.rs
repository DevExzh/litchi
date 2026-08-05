//! Contextual routing-slip values and the semantic document model.

use crate::package::{PptError, Result};

/// A bounded legacy printable-ANSI routing-slip string.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Text {
    pub(crate) bytes: Vec<u8>,
}

impl Text {
    /// Build a routing-slip string from its printable-ANSI bytes.
    pub fn from_ansi_bytes(bytes: Vec<u8>) -> Result<Self> {
        if bytes.contains(&0) {
            return Err(PptError::Corrupted(
                "routing-slip text contains an embedded NUL".to_string(),
            ));
        }
        if bytes.len() > usize::from(u16::MAX) - 1 {
            return Err(PptError::Corrupted(
                "routing-slip text exceeds the 16-bit length limit".to_string(),
            ));
        }
        Ok(Self { bytes })
    }

    /// Borrow the original printable-ANSI bytes.
    pub fn as_ansi_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Decode the legacy bytes lossily as one-byte characters.
    pub fn to_string_lossy(&self) -> String {
        self.bytes.iter().map(|&byte| char::from(byte)).collect()
    }
}

/// An originator or recipient entry in a routing slip.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Address {
    pub text: Text,
    /// The undefined byte following the address terminator, retained for
    /// lossless round trips.
    pub trailing_undefined: u8,
}

impl Address {
    /// Create an address with the specification's conventional undefined byte.
    pub fn new(text: Text) -> Self {
        Self {
            text,
            trailing_undefined: 0,
        }
    }
}

/// The addressee currently processing the document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CurrentRecipient {
    /// The originator before routing starts.
    OriginatorBeforeRouting,
    /// A one-based recipient position.
    Recipient(u32),
    /// The originator after every recipient has processed the document.
    OriginatorAfterRouting,
}

/// The typed document-level routing-slip record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Slip {
    pub originator: Address,
    pub recipients: Vec<Address>,
    pub current_recipient: CurrentRecipient,
    pub subject: Text,
    pub message: Text,
    pub one_after_another: bool,
    pub return_when_done: bool,
    pub track_status: bool,
    pub document_routed: bool,
    pub cycle_completed: bool,
    /// Undefined `unused1` bytes represented as their native little-endian
    /// value. They are retained and never interpreted.
    pub unused1: u32,
    /// Undefined `unused2` bytes represented as their native little-endian
    /// value. They are retained and never interpreted.
    pub unused2: u32,
    /// The variable-length `unused3` tail, including every source byte after
    /// the meaningful routing-slip payload.
    pub trailing_undefined: Vec<u8>,
}

impl Slip {
    /// Create a new routing slip in its pre-routing state.
    pub fn new(
        originator: Address,
        recipients: Vec<Address>,
        subject: Text,
        message: Text,
    ) -> Result<Self> {
        if recipients.len() > u32::MAX as usize {
            return Err(PptError::Corrupted(
                "routing slip has too many recipients".to_string(),
            ));
        }
        Ok(Self {
            originator,
            recipients,
            current_recipient: CurrentRecipient::OriginatorBeforeRouting,
            subject,
            message,
            one_after_another: false,
            return_when_done: false,
            track_status: false,
            document_routed: false,
            cycle_completed: false,
            unused1: 0,
            unused2: 0,
            trailing_undefined: vec![0; 8],
        })
    }
}
