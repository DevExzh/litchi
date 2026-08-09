//! Semantic ownership of one source-preserving toolbar control.

use std::borrow::Cow;

use super::validation;
use super::{ControlHeader, Data, Error, TextIcon, WString};

/// The variable bytes following a [`ControlHeader`].
///
/// The shared `Data` variant owns the common `TBCGeneralInfo` fields and keeps
/// its format-specific tail opaque. `Opaque` is used for command prefixes,
/// unsupported control-specific structures, and payloads that a host format
/// has not yet bounded. No variant executes or interprets a command.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Body<'a> {
    /// A control with no variable body, such as an `ActiveX` header-only form.
    Empty,
    /// A decoded common `TBCData` body with an opaque specific tail.
    Data(Data<'a>),
    /// Bytes retained without semantic interpretation.
    Opaque(Cow<'a, [u8]>),
}

impl<'a> Body<'a> {
    /// Construct an empty control body.
    #[must_use]
    pub const fn empty() -> Self {
        Self::Empty
    }

    /// Construct a decoded common control-data body.
    #[must_use]
    pub fn data(value: Data<'a>) -> Self {
        Self::Data(value)
    }

    /// Construct an opaque body without interpreting its bytes.
    pub fn opaque(value: impl Into<Cow<'a, [u8]>>) -> Self {
        Self::Opaque(value.into())
    }

    /// Return the decoded common data, when supported.
    #[must_use]
    pub const fn data_ref(&self) -> Option<&Data<'a>> {
        match self {
            Self::Data(value) => Some(value),
            Self::Empty | Self::Opaque(_) => None,
        }
    }

    /// Return opaque bytes, when this body is not decoded.
    #[must_use]
    pub fn opaque_bytes(&self) -> Option<&[u8]> {
        match self {
            Self::Opaque(value) => Some(value),
            Self::Empty | Self::Data(_) => None,
        }
    }

    /// Return the exact serialized body bytes.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        match self {
            Self::Empty => Vec::new(),
            Self::Data(value) => value.to_bytes(),
            Self::Opaque(value) => value.to_vec(),
        }
    }

    pub(crate) fn into_owned(self) -> Body<'static> {
        match self {
            Self::Empty => Body::Empty,
            Self::Data(value) => Body::Data(value.into_owned()),
            Self::Opaque(value) => Body::Opaque(Cow::Owned(value.into_owned())),
        }
    }
}

/// A bounded, inert `[MS-OSHARED]` toolbar control.
///
/// `prefix` is deliberately opaque because DOC and XLS place different
/// command structures between `TBCHeader` and `TBCData`. The body preserves
/// unsupported control-specific structures byte-for-byte. The common layer
/// therefore supplies safe metadata edits without pretending to own a host
/// format's record boundaries or command vocabulary.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Control<'a> {
    header: ControlHeader,
    prefix: Cow<'a, [u8]>,
    body: Body<'a>,
}

impl<'a> Control<'a> {
    /// Construct a control with no opaque command prefix.
    ///
    /// # Errors
    ///
    /// Returns an error if `header` and `body` do not form a valid authored
    /// control.
    pub fn new(header: ControlHeader, body: Body<'a>) -> Result<Self, Error> {
        Self::from_parts(header, &[][..], body)
    }

    /// Construct a control while retaining its host-specific prefix bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the fields do not form a valid authored control.
    pub fn from_parts(
        header: ControlHeader,
        prefix: impl Into<Cow<'a, [u8]>>,
        body: Body<'a>,
    ) -> Result<Self, Error> {
        let value = Self {
            header,
            prefix: prefix.into(),
            body,
        };
        validation::validate_authored(&value)?;
        Ok(value)
    }

    /// Borrow the fixed `TBCHeader` projection.
    #[must_use]
    pub const fn header(&self) -> &ControlHeader {
        &self.header
    }

    /// Borrow the opaque host-specific prefix.
    #[must_use]
    pub fn prefix(&self) -> &[u8] {
        &self.prefix
    }

    /// Borrow the common or opaque variable body.
    #[must_use]
    pub const fn body(&self) -> &Body<'a> {
        &self.body
    }

    /// Borrow common control metadata, when the body is decoded.
    #[must_use]
    pub const fn data(&self) -> Option<&Data<'a>> {
        self.body.data_ref()
    }

    /// Return the shared text/icon visibility mode.
    #[must_use]
    pub const fn text_icon(&self) -> TextIcon {
        self.header.specifics().text_icon()
    }

    /// Return the custom text when the common body is decoded and present.
    #[must_use]
    pub const fn custom_text(&self) -> Option<&WString<'a>> {
        match self.data() {
            Some(data) => data.general().custom_text(),
            None => None,
        }
    }

    /// Return the description when the common body is decoded and present.
    #[must_use]
    pub const fn description(&self) -> Option<&WString<'a>> {
        match self.data() {
            Some(data) => data.general().description(),
            None => None,
        }
    }

    /// Return the tooltip when the common body is decoded and present.
    #[must_use]
    pub const fn tooltip(&self) -> Option<&WString<'a>> {
        match self.data() {
            Some(data) => data.general().tooltip(),
            None => None,
        }
    }

    /// Return the exact serialized control bytes.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        let body = self.body.to_bytes();
        let mut output =
            Vec::with_capacity(self.header.to_bytes().len() + self.prefix.len() + body.len());
        output.extend_from_slice(&self.header.to_bytes());
        output.extend_from_slice(&self.prefix);
        output.extend_from_slice(&body);
        output
    }

    /// Move all borrowed projections into an owned control.
    #[must_use]
    pub fn into_owned(self) -> Control<'static> {
        Control {
            header: self.header,
            prefix: Cow::Owned(self.prefix.into_owned()),
            body: self.body.into_owned(),
        }
    }

    pub(crate) fn from_decoded(
        header: ControlHeader,
        prefix: &'a [u8],
        body: Body<'a>,
    ) -> Result<Self, Error> {
        let value = Self {
            header,
            prefix: Cow::Borrowed(prefix),
            body,
        };
        validation::validate_decoded(&value)?;
        Ok(value)
    }

    pub(crate) fn from_edited(
        header: ControlHeader,
        prefix: Cow<'static, [u8]>,
        body: Body<'static>,
    ) -> Result<Self, Error> {
        let value = Self {
            header,
            prefix,
            body,
        };
        validation::validate_edited(&value)?;
        Ok(value)
    }
}
