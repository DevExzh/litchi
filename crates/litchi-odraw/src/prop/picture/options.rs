use super::{Error, LimitKind, Result};

/// Finite admission and publication ceilings for one picture property table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    property_bytes: usize,
    properties: usize,
    name_units: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self::new()
    }
}

impl Limits {
    pub(super) const fn from_admitted(
        property_bytes: usize,
        properties: usize,
        name_units: usize,
    ) -> Self {
        Self {
            property_bytes,
            properties,
            name_units,
        }
    }
    /// Uses the previous finite property-table and picture-name ceilings.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            property_bytes: 16 * 1024 * 1024,
            properties: 0x0fff,
            name_units: 4096,
        }
    }

    /// Bounds the property-record body, excluding its eight-byte wire header.
    pub fn property_bytes(mut self, maximum: usize) -> Result<Self> {
        validate(
            LimitKind::PropertyBytes,
            maximum,
            (u32::MAX as usize).min(usize::MAX - 8),
        )?;
        self.property_bytes = maximum;
        Ok(self)
    }

    /// Bounds the descriptor count; zero admits only an empty table.
    pub fn properties(mut self, maximum: usize) -> Result<Self> {
        validate(LimitKind::Properties, maximum, 0x0fff)?;
        self.properties = maximum;
        Ok(self)
    }

    /// Bounds UTF-16 units excluding the terminating NUL.
    pub fn name_units(mut self, maximum: usize) -> Result<Self> {
        validate(LimitKind::NameUnits, maximum, (i32::MAX as usize / 2) - 1)?;
        self.name_units = maximum;
        Ok(self)
    }

    #[must_use]
    pub const fn property_bytes_value(self) -> usize {
        self.property_bytes
    }

    #[must_use]
    pub const fn properties_value(self) -> usize {
        self.properties
    }

    #[must_use]
    pub const fn name_units_value(self) -> usize {
        self.name_units
    }

    pub(super) fn constrained_by(self, other: Self) -> Self {
        Self {
            property_bytes: self.property_bytes.min(other.property_bytes),
            properties: self.properties.min(other.properties),
            name_units: self.name_units.min(other.name_units),
        }
    }
}

fn validate(kind: LimitKind, value: usize, maximum: usize) -> Result<()> {
    if value > maximum {
        return Err(Error::InvalidLimit {
            kind,
            value: value as u64,
            maximum: maximum as u64,
        });
    }
    Ok(())
}

/// Policy retained by an immutable picture package.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct OpenOptions {
    limits: Limits,
}

impl OpenOptions {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            limits: Limits::new(),
        }
    }

    #[must_use]
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = limits;
        self
    }

    #[must_use]
    pub const fn limits_value(self) -> Limits {
        self.limits
    }
}

/// Optional tighter policy for one edit; the default inherits ingress limits.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    limits: Option<Limits>,
}

impl EditOptions {
    #[must_use]
    pub const fn new() -> Self {
        Self { limits: None }
    }

    #[must_use]
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = Some(limits);
        self
    }

    #[must_use]
    pub const fn limits_value(self) -> Option<Limits> {
        self.limits
    }
}
