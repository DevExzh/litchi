use super::{HYPERLINKS, LINK_BASE, codec};
use crate::property_set::model::try_clone_property_set;
use crate::property_set::{CodePage, Section, USER_DEFINED_PROPERTIES_FMTID, Value};
use litchi_cfb::OleError;

const MIN_NAMED_PROPERTY_ID: u32 = 2;
const MAX_NAMED_PROPERTY_ID: u32 = 0x00ff_ffff;

/// Bounded, caller-configurable limits for the inert hyperlink property blobs.
///
/// These bound this typed overlay's secondary decoding and allocation after the
/// generic [`Section`] parser has already read a property stream. They are not
/// limits on the initial generic `VT_BLOB` property allocation. Use the builder
/// when document-specific policy is stricter than the safe defaults.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    blob_bytes: usize,
    link_count: usize,
    string_units: usize,
    total_utf16_units: usize,
}

impl Limits {
    /// Creates a builder initialized to conservative safe defaults.
    #[must_use]
    pub const fn builder() -> LimitsBuilder {
        LimitsBuilder {
            blob_bytes: Self::DEFAULT.blob_bytes,
            link_count: Self::DEFAULT.link_count,
            string_units: Self::DEFAULT.string_units,
            total_utf16_units: Self::DEFAULT.total_utf16_units,
        }
    }

    /// Maximum encoded bytes accepted for either reserved BLOB payload.
    #[must_use]
    pub const fn max_blob_bytes(self) -> usize {
        self.blob_bytes
    }

    /// Maximum hyperlink records accepted in `_PID_HLINKS`.
    #[must_use]
    pub const fn max_links(self) -> usize {
        self.link_count
    }

    /// Maximum UTF-16 code units, including a terminating NUL, per string.
    #[must_use]
    pub const fn max_string_units(self) -> usize {
        self.string_units
    }

    /// Maximum decoded UTF-16 code units across one typed property read.
    #[must_use]
    pub const fn max_total_utf16_units(self) -> usize {
        self.total_utf16_units
    }

    const DEFAULT: Self = Self {
        blob_bytes: 16 * 1024 * 1024,
        link_count: 100_000,
        string_units: 65_536,
        total_utf16_units: 1_000_000,
    };
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Builder for [`Limits`]. Every cap must be nonzero.
#[derive(Debug, Clone, Copy)]
pub struct LimitsBuilder {
    blob_bytes: usize,
    link_count: usize,
    string_units: usize,
    total_utf16_units: usize,
}

impl LimitsBuilder {
    /// Changes the encoded BLOB cap.
    #[must_use]
    pub const fn max_blob_bytes(mut self, value: usize) -> Self {
        self.blob_bytes = value;
        self
    }

    /// Changes the hyperlink-record cap.
    #[must_use]
    pub const fn max_links(mut self, value: usize) -> Self {
        self.link_count = value;
        self
    }

    /// Changes the cap for one UTF-16 string, including its NUL terminator.
    #[must_use]
    pub const fn max_string_units(mut self, value: usize) -> Self {
        self.string_units = value;
        self
    }

    /// Changes the aggregate UTF-16 decode cap.
    #[must_use]
    pub const fn max_total_utf16_units(mut self, value: usize) -> Self {
        self.total_utf16_units = value;
        self
    }

    /// Validates and returns the configured limits.
    ///
    /// # Errors
    ///
    /// Returns an error if any configured cap is zero.
    pub fn build(self) -> Result<Limits, OleError> {
        if self.blob_bytes == 0 {
            return Err(invalid(
                "user-defined hyperlink maximum blob bytes must be nonzero",
            ));
        }
        if self.link_count == 0 {
            return Err(invalid(
                "user-defined hyperlink maximum link count must be nonzero",
            ));
        }
        if self.string_units == 0 {
            return Err(invalid(
                "user-defined hyperlink maximum string units must be nonzero",
            ));
        }
        if self.total_utf16_units == 0 {
            return Err(invalid(
                "user-defined hyperlink maximum aggregate string units must be nonzero",
            ));
        }
        Ok(Limits {
            blob_bytes: self.blob_bytes,
            link_count: self.link_count,
            string_units: self.string_units,
            total_utf16_units: self.total_utf16_units,
        })
    }
}

/// An immutable, inert `_PID_LINKBASE` value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LinkBase {
    value: String,
}

impl LinkBase {
    /// Creates a base URL text value without resolving or otherwise using it.
    ///
    /// # Errors
    ///
    /// Returns an error if `value` contains a NUL character.
    pub fn new(value: impl Into<String>) -> Result<Self, OleError> {
        let text = value.into();
        validate_text(&text, "link base")?;
        Ok(Self { value: text })
    }

    /// Borrows the stored base text.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// One inert `VtHyperlink` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Hyperlink {
    stored_hash: u32,
    app: i32,
    office_art: i32,
    info: i32,
    target: String,
    location: String,
}

impl Hyperlink {
    /// Creates a record and computes the MS-OSHARED 2.4.2 SHOULD hash.
    ///
    /// # Errors
    ///
    /// Returns an error if either inert text field contains a NUL character.
    pub fn new(
        app: i32,
        office_art: i32,
        info: i32,
        target: impl Into<String>,
        location: impl Into<String>,
    ) -> Result<Self, OleError> {
        let target_text = target.into();
        let location_text = location.into();
        validate_text(&target_text, "hyperlink target")?;
        validate_text(&location_text, "hyperlink location")?;
        Ok(Self {
            stored_hash: hyperlink_hash(&target_text, &location_text),
            app,
            office_art,
            info,
            target: target_text,
            location: location_text,
        })
    }

    pub(crate) fn from_wire(
        stored_hash: u32,
        app: i32,
        office_art: i32,
        info: i32,
        target: String,
        location: String,
    ) -> Self {
        Self {
            stored_hash,
            app,
            office_art,
            info,
            target,
            location,
        }
    }

    /// The hash as it was stored in the document, retained losslessly on read.
    #[must_use]
    pub const fn stored_hash(&self) -> u32 {
        self.stored_hash
    }

    /// Calculates the MS-OSHARED 2.4.2 hash from the inert target and location.
    #[must_use]
    pub fn calculated_hash(&self) -> u32 {
        hyperlink_hash(&self.target, &self.location)
    }

    /// Checks whether the retained stored hash matches the calculated SHOULD hash.
    #[must_use]
    pub fn hash_matches(&self) -> bool {
        self.stored_hash == self.calculated_hash()
    }

    /// Implementation-specific application value.
    #[must_use]
    pub const fn app(&self) -> i32 {
        self.app
    }

    /// Inert `OfficeArt` shape identifier, or zero when not shape-bound.
    #[must_use]
    pub const fn office_art(&self) -> i32 {
        self.office_art
    }

    /// Implementation-specific link state value.
    #[must_use]
    pub const fn info(&self) -> i32 {
        self.info
    }

    /// The inert hyperlink target text.
    #[must_use]
    pub fn target(&self) -> &str {
        &self.target
    }

    /// The inert hyperlink location text.
    #[must_use]
    pub fn location(&self) -> &str {
        &self.location
    }
}

/// An immutable, inert `_PID_HLINKS` collection.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Hyperlinks {
    links: Vec<Hyperlink>,
}

impl Hyperlinks {
    /// Creates an inert hyperlink collection. Limits are applied on read/write.
    #[must_use]
    pub fn new(links: Vec<Hyperlink>) -> Self {
        Self { links }
    }

    /// Borrows all records in their stored order.
    #[must_use]
    pub fn links(&self) -> &[Hyperlink] {
        &self.links
    }

    /// Iterates records in their stored order.
    #[must_use]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Hyperlink> {
        self.links.iter()
    }

    /// Returns the number of records.
    #[must_use]
    pub fn len(&self) -> usize {
        self.links.len()
    }

    /// Whether the list contains no records.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.links.is_empty()
    }
}

/// Lazy typed reads over exactly one `FMTID_UserDefinedProperties` section.
#[derive(Debug, Clone, Copy)]
pub struct Properties<'a> {
    section: &'a Section,
    limits: Limits,
}

impl<'a> Properties<'a> {
    /// Creates a lazy view using [`Limits::default`].
    ///
    /// # Errors
    ///
    /// Returns an error if `section` is not a `UserDefinedProperties` section.
    pub fn new(section: &'a Section) -> Result<Self, OleError> {
        Self::with_limits(section, Limits::default())
    }

    /// Creates a lazy view using caller-supplied validated limits.
    ///
    /// # Errors
    ///
    /// Returns an error if `section` is not a `UserDefinedProperties` section.
    pub fn with_limits(section: &'a Section, limits: Limits) -> Result<Self, OleError> {
        validate_section(section)?;
        Ok(Self { section, limits })
    }

    /// Borrows the complete section, including unknown properties.
    #[must_use]
    pub const fn section(&self) -> &'a Section {
        self.section
    }

    /// Returns the limits used for subsequent lazy reads.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Reads `_PID_LINKBASE` only when it is accessed.
    ///
    /// # Errors
    ///
    /// Returns an error if the named property is not a BLOB or its payload
    /// violates the configured decoding limits.
    pub fn link_base(&self) -> Result<Option<LinkBase>, OleError> {
        self.named_blob(LINK_BASE)?
            .map(|value| codec::decode_link_base(value, self.limits))
            .transpose()
    }

    /// Reads `_PID_HLINKS` only when it is accessed.
    ///
    /// # Errors
    ///
    /// Returns an error if the named property is not a BLOB or its payload
    /// violates the configured decoding limits.
    pub fn hyperlinks(&self) -> Result<Option<Hyperlinks>, OleError> {
        self.named_blob(HYPERLINKS)?
            .map(|value| codec::decode_hyperlinks(value, self.limits))
            .transpose()
    }

    fn named_blob(&self, name: &str) -> Result<Option<&'a [u8]>, OleError> {
        match self.section.find_named(name) {
            None => Ok(None),
            Some((_, Value::Blob(value))) => Ok(Some(value)),
            Some(_) => Err(invalid(format!("{name} must be a VT_BLOB property"))),
        }
    }
}

/// Typed, transactional-in-effect mutations for the reserved named properties.
///
/// Values are completely encoded and bounded before their section is changed,
/// so a failed typed edit leaves the source section untouched.
pub struct Edit<'a> {
    section: &'a mut Section,
    limits: Limits,
}

impl<'a> Edit<'a> {
    /// Creates an editor using [`Limits::default`].
    ///
    /// # Errors
    ///
    /// Returns an error if `section` is not a `UserDefinedProperties` section or
    /// an existing reserved property has an invalid identifier.
    pub fn new(section: &'a mut Section) -> Result<Self, OleError> {
        Self::with_limits(section, Limits::default())
    }

    /// Creates an editor using caller-supplied validated limits.
    ///
    /// # Errors
    ///
    /// Returns an error if `section` is not a `UserDefinedProperties` section or
    /// an existing reserved property has an invalid identifier.
    pub fn with_limits(section: &'a mut Section, limits: Limits) -> Result<Self, OleError> {
        validate_edit_section(section)?;
        Ok(Self { section, limits })
    }

    /// Borrows the complete editable section.
    #[must_use]
    pub const fn section(&self) -> &Section {
        self.section
    }

    /// Sets `_PID_LINKBASE`, preserving an existing PID/name spelling/order.
    ///
    /// # Errors
    ///
    /// Returns an error if the value cannot be encoded within the configured
    /// limits or the section cannot accept the replacement.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "the public edit API intentionally accepts temporary validated values by value"
    )]
    pub fn set_link_base(&mut self, value: LinkBase) -> Result<(), OleError> {
        let encoded_value = Value::Blob(codec::encode_link_base(&value, self.limits)?);
        self.set_named(LINK_BASE, encoded_value)
    }

    /// Removes `_PID_LINKBASE` if present.
    pub fn remove_link_base(&mut self) -> bool {
        self.remove_named(LINK_BASE)
    }

    /// Sets `_PID_HLINKS`, preserving an existing PID/name spelling/order.
    ///
    /// # Errors
    ///
    /// Returns an error if the value cannot be encoded within the configured
    /// limits or the section cannot accept the replacement.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "the public edit API intentionally accepts temporary validated values by value"
    )]
    pub fn set_hyperlinks(&mut self, value: Hyperlinks) -> Result<(), OleError> {
        let encoded_value = Value::Blob(codec::encode_hyperlinks(&value, self.limits)?);
        self.set_named(HYPERLINKS, encoded_value)
    }

    /// Removes `_PID_HLINKS` if present.
    pub fn remove_hyperlinks(&mut self) -> bool {
        self.remove_named(HYPERLINKS)
    }

    fn set_named(&mut self, name: &str, value: Value) -> Result<(), OleError> {
        let mut draft = try_clone_property_set(self.section)?;
        if draft.page().is_none() {
            draft.set_page(CodePage::WINDOWS_1252);
        }
        if let Some((identifier, _)) = draft.find_named(name) {
            draft.update(identifier, value)?;
        } else {
            let identifier = unused_named_identifier(&draft)?;
            draft.add_named(identifier, name.to_owned(), value)?;
        }
        *self.section = draft;
        Ok(())
    }

    fn remove_named(&mut self, name: &str) -> bool {
        let Some((identifier, _)) = self.section.find_named(name) else {
            return false;
        };
        self.section.remove(identifier).is_some()
    }
}

fn invalid(message: impl Into<String>) -> OleError {
    OleError::InvalidFormat(message.into())
}

fn validate_section(section: &Section) -> Result<(), OleError> {
    if section.format_identifier != USER_DEFINED_PROPERTIES_FMTID {
        return Err(invalid(
            "hyperlink properties require the UserDefinedProperties section",
        ));
    }
    Ok(())
}

fn validate_edit_section(section: &Section) -> Result<(), OleError> {
    validate_section(section)?;
    for name in [LINK_BASE, HYPERLINKS] {
        if let Some((identifier, _)) = section.find_named(name)
            && identifier > MAX_NAMED_PROPERTY_ID
        {
            return Err(invalid(format!(
                "{name} property identifier {identifier} exceeds the UserDefinedProperties effective maximum"
            )));
        }
    }
    Ok(())
}

fn unused_named_identifier(section: &Section) -> Result<u32, OleError> {
    for identifier in MIN_NAMED_PROPERTY_ID..=MAX_NAMED_PROPERTY_ID {
        if section.property(identifier).is_none() {
            return Ok(identifier);
        }
    }
    Err(invalid(
        "no unused user-defined property identifiers remain",
    ))
}

fn validate_text(value: &str, field: &str) -> Result<(), OleError> {
    if value.contains('\0') {
        return Err(invalid(format!("{field} must not contain NUL")));
    }
    Ok(())
}

fn hyperlink_hash(target: &str, location: &str) -> u32 {
    unicode_hash(target) ^ unicode_hash(location)
}

fn unicode_hash(value: &str) -> u32 {
    let mut hash = 0u32;
    let mut first = None;
    let mut remaining = 255usize;
    for character in value.chars() {
        let mut unit_buffer = [0; 2];
        let encoded_units = character.encode_utf16(&mut unit_buffer);
        if encoded_units.len() > remaining {
            // MS-OSHARED counts WCHARs. A surrogate cut at the exact boundary
            // cannot be lowercased as a scalar, so retain that code unit.
            hash_unit(&mut hash, &mut first, encoded_units[0]);
            break;
        }
        remaining -= encoded_units.len();
        for lowered in character.to_lowercase() {
            for unit in lowered.encode_utf16(&mut [0; 2]) {
                hash_unit(&mut hash, &mut first, *unit);
            }
        }
        if remaining == 0 {
            break;
        }
    }
    // The required trailing NUL only completes an odd pair and is zero.
    hash ^ first.map_or(0, u32::from)
}

fn hash_unit(hash: &mut u32, first: &mut Option<u16>, unit: u16) {
    if let Some(previous) = first.take() {
        *hash ^= u32::from(previous) ^ (u32::from(unit) << 16);
    } else {
        *first = Some(unit);
    }
}
