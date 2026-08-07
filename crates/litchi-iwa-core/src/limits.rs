use crate::error::{Error, LimitKind, Result};

/// Hard and caller-selectable bounds for one decompressed IWA archive.
///
/// The type is intentionally copyable and contains no allocation. Every
/// parser entry point validates it before it reads a header or reserves a
/// metadata collection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::struct_field_names,
    reason = "The max_* vocabulary makes the resource-budget fields self-documenting."
)]
pub struct Limits {
    max_archive_bytes: usize,
    max_objects: usize,
    max_messages: usize,
    max_messages_per_object: usize,
    max_object_bytes: usize,
    max_message_bytes: usize,
    max_header_bytes: usize,
    max_metadata_items: usize,
}

impl Limits {
    /// Hard ceiling for one decompressed IWA component.
    pub const MAX_ARCHIVE_BYTES: usize = 512 * 1024 * 1024;
    /// Hard ceiling for the number of objects in one component.
    pub const MAX_OBJECTS: usize = 100_000;
    /// Hard ceiling for the number of messages in one component.
    pub const MAX_MESSAGES: usize = 1_000_000;
    /// Hard ceiling for messages described by one object header.
    pub const MAX_MESSAGES_PER_OBJECT: usize = 100_000;
    /// Hard ceiling for one object, including its header and payload.
    pub const MAX_OBJECT_BYTES: usize = 512 * 1024 * 1024;
    /// Hard ceiling for one message payload.
    pub const MAX_MESSAGE_BYTES: usize = 512 * 1024 * 1024;
    /// Hard ceiling for one encoded `TSP.ArchiveInfo` header.
    pub const MAX_HEADER_BYTES: usize = 16 * 1024 * 1024;
    /// Hard ceiling for repeated metadata items in one object header.
    pub const MAX_METADATA_ITEMS: usize = 1_000_000;

    /// Tighten the aggregate archive-byte budget.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero, exceeds the hard ceiling, or
    /// would violate a cross-field limit invariant.
    pub fn with_archive_bytes(mut self, value: usize) -> Result<Self> {
        check(LimitKind::ArchiveBytes, value, Self::MAX_ARCHIVE_BYTES)?;
        self.max_archive_bytes = value;
        self.max_object_bytes = self.max_object_bytes.min(value);
        self.max_message_bytes = self.max_message_bytes.min(value);
        self.max_header_bytes = self.max_header_bytes.min(value);
        self.validate()
    }

    /// Tighten the object-count budget.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero, exceeds the hard ceiling, or
    /// would violate a cross-field limit invariant.
    pub fn with_objects(mut self, value: usize) -> Result<Self> {
        check(LimitKind::Objects, value, Self::MAX_OBJECTS)?;
        self.max_objects = value;
        self.validate()
    }

    /// Tighten the aggregate message-count budget.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero, exceeds the hard ceiling, or
    /// would violate a cross-field limit invariant.
    pub fn with_messages(mut self, value: usize) -> Result<Self> {
        check(LimitKind::Messages, value, Self::MAX_MESSAGES)?;
        self.max_messages = value;
        self.max_messages_per_object = self.max_messages_per_object.min(value);
        self.validate()
    }

    /// Tighten the per-object message-count budget.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero, exceeds the hard ceiling, or
    /// would violate a cross-field limit invariant.
    pub fn with_messages_per_object(mut self, value: usize) -> Result<Self> {
        check(
            LimitKind::MessagesPerObject,
            value,
            Self::MAX_MESSAGES_PER_OBJECT,
        )?;
        self.max_messages_per_object = value;
        self.validate()
    }

    /// Tighten the per-object wire-byte budget.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero, exceeds the hard ceiling, or
    /// would violate a cross-field limit invariant.
    pub fn with_object_bytes(mut self, value: usize) -> Result<Self> {
        check(LimitKind::ObjectBytes, value, Self::MAX_OBJECT_BYTES)?;
        self.max_object_bytes = value;
        self.max_message_bytes = self.max_message_bytes.min(value);
        self.max_header_bytes = self.max_header_bytes.min(value);
        self.validate()
    }

    /// Tighten the per-message payload-byte budget.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero, exceeds the hard ceiling, or
    /// would violate a cross-field limit invariant.
    pub fn with_message_bytes(mut self, value: usize) -> Result<Self> {
        check(LimitKind::MessageBytes, value, Self::MAX_MESSAGE_BYTES)?;
        self.max_message_bytes = value;
        self.validate()
    }

    /// Tighten the encoded-header-byte budget.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero, exceeds the hard ceiling, or
    /// would violate a cross-field limit invariant.
    pub fn with_header_bytes(mut self, value: usize) -> Result<Self> {
        check(LimitKind::HeaderBytes, value, Self::MAX_HEADER_BYTES)?;
        self.max_header_bytes = value;
        self.validate()
    }

    /// Tighten the repeated metadata-item budget in one object header.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is zero or exceeds the hard ceiling.
    pub fn with_metadata_items(mut self, value: usize) -> Result<Self> {
        check(LimitKind::MetadataItems, value, Self::MAX_METADATA_ITEMS)?;
        self.max_metadata_items = value;
        self.validate()
    }

    #[must_use]
    pub const fn max_archive_bytes(self) -> usize {
        self.max_archive_bytes
    }

    #[must_use]
    pub const fn max_objects(self) -> usize {
        self.max_objects
    }

    #[must_use]
    pub const fn max_messages(self) -> usize {
        self.max_messages
    }

    #[must_use]
    pub const fn max_messages_per_object(self) -> usize {
        self.max_messages_per_object
    }

    #[must_use]
    pub const fn max_object_bytes(self) -> usize {
        self.max_object_bytes
    }

    #[must_use]
    pub const fn max_message_bytes(self) -> usize {
        self.max_message_bytes
    }

    #[must_use]
    pub const fn max_header_bytes(self) -> usize {
        self.max_header_bytes
    }

    #[must_use]
    pub const fn max_metadata_items(self) -> usize {
        self.max_metadata_items
    }

    /// Validate cross-field invariants and the format hard ceilings.
    ///
    /// # Errors
    ///
    /// Returns an error if any configured budget is zero or if a narrower
    /// budget exceeds the aggregate budget that contains it.
    pub fn validate(self) -> Result<Self> {
        if self.max_archive_bytes == 0
            || self.max_objects == 0
            || self.max_messages == 0
            || self.max_messages_per_object == 0
            || self.max_object_bytes == 0
            || self.max_message_bytes == 0
            || self.max_header_bytes == 0
            || self.max_metadata_items == 0
        {
            return Err(Error::invalid_limits("all IWA limits must be non-zero"));
        }
        if self.max_messages_per_object > self.max_messages {
            return Err(Error::invalid_limits(
                "per-object message limit exceeds aggregate message limit",
            ));
        }
        if self.max_object_bytes > self.max_archive_bytes {
            return Err(Error::invalid_limits(
                "object byte limit exceeds archive byte limit",
            ));
        }
        if self.max_message_bytes > self.max_object_bytes {
            return Err(Error::invalid_limits(
                "message byte limit exceeds object byte limit",
            ));
        }
        if self.max_header_bytes > self.max_object_bytes {
            return Err(Error::invalid_limits(
                "header byte limit exceeds object byte limit",
            ));
        }
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_archive_bytes: Self::MAX_ARCHIVE_BYTES,
            max_objects: Self::MAX_OBJECTS,
            max_messages: Self::MAX_MESSAGES,
            max_messages_per_object: Self::MAX_MESSAGES_PER_OBJECT,
            max_object_bytes: Self::MAX_OBJECT_BYTES,
            max_message_bytes: Self::MAX_MESSAGE_BYTES,
            max_header_bytes: Self::MAX_HEADER_BYTES,
            max_metadata_items: Self::MAX_METADATA_ITEMS,
        }
    }
}

fn check(kind: LimitKind, value: usize, maximum: usize) -> Result<()> {
    if value == 0 {
        return Err(Error::invalid_limits("limits must be non-zero"));
    }
    if value > maximum {
        return Err(Error::limit(kind, value, maximum));
    }
    Ok(())
}
