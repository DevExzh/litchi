use super::{HYPERLINKS, LINK_BASE, codec};
use crate::property_set::model::try_clone_property_set;
use crate::property_set::{CodePage, Section, USER_DEFINED_PROPERTIES_FMTID, Value};
use litchi_cfb::OleError;

const MIN_NAMED_PROPERTY_ID: u32 = 2;
const MAX_NAMED_PROPERTY_ID: u32 = 0x00ff_ffff;

fn invalid(message: impl Into<String>) -> OleError {
    OleError::InvalidFormat(message.into())
}

/// Bounded, caller-configurable limits for the inert hyperlink property blobs.
///
/// These bound this typed overlay's secondary decoding and allocation after the
/// generic [`Section`] parser has already read a property stream. They are not
/// limits on the initial generic `VT_BLOB` property allocation. Use the builder
/// when document-specific policy is stricter than the safe defaults.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_blob_bytes: usize,
    max_links: usize,
    max_string_units: usize,
    max_total_utf16_units: usize,
}

impl Limits {
    /// Creates a builder initialized to conservative safe defaults.
    pub const fn builder() -> LimitsBuilder {
        LimitsBuilder {
            max_blob_bytes: Self::DEFAULT.max_blob_bytes,
            max_links: Self::DEFAULT.max_links,
            max_string_units: Self::DEFAULT.max_string_units,
            max_total_utf16_units: Self::DEFAULT.max_total_utf16_units,
        }
    }

    /// Maximum encoded bytes accepted for either reserved BLOB payload.
    pub const fn max_blob_bytes(self) -> usize {
        self.max_blob_bytes
    }

    /// Maximum hyperlink records accepted in `_PID_HLINKS`.
    pub const fn max_links(self) -> usize {
        self.max_links
    }

    /// Maximum UTF-16 code units, including a terminating NUL, per string.
    pub const fn max_string_units(self) -> usize {
        self.max_string_units
    }

    /// Maximum decoded UTF-16 code units across one typed property read.
    pub const fn max_total_utf16_units(self) -> usize {
        self.max_total_utf16_units
    }

    const DEFAULT: Self = Self {
        max_blob_bytes: 16 * 1024 * 1024,
        max_links: 100_000,
        max_string_units: 65_536,
        max_total_utf16_units: 1_000_000,
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
    max_blob_bytes: usize,
    max_links: usize,
    max_string_units: usize,
    max_total_utf16_units: usize,
}

impl LimitsBuilder {
    /// Changes the encoded BLOB cap.
    pub const fn max_blob_bytes(mut self, value: usize) -> Self {
        self.max_blob_bytes = value;
        self
    }

    /// Changes the hyperlink-record cap.
    pub const fn max_links(mut self, value: usize) -> Self {
        self.max_links = value;
        self
    }

    /// Changes the cap for one UTF-16 string, including its NUL terminator.
    pub const fn max_string_units(mut self, value: usize) -> Self {
        self.max_string_units = value;
        self
    }

    /// Changes the aggregate UTF-16 decode cap.
    pub const fn max_total_utf16_units(mut self, value: usize) -> Self {
        self.max_total_utf16_units = value;
        self
    }

    /// Validates and returns the configured limits.
    pub fn build(self) -> Result<Limits, OleError> {
        if self.max_blob_bytes == 0 {
            return Err(invalid(
                "user-defined hyperlink maximum blob bytes must be nonzero",
            ));
        }
        if self.max_links == 0 {
            return Err(invalid(
                "user-defined hyperlink maximum link count must be nonzero",
            ));
        }
        if self.max_string_units == 0 {
            return Err(invalid(
                "user-defined hyperlink maximum string units must be nonzero",
            ));
        }
        if self.max_total_utf16_units == 0 {
            return Err(invalid(
                "user-defined hyperlink maximum aggregate string units must be nonzero",
            ));
        }
        Ok(Limits {
            max_blob_bytes: self.max_blob_bytes,
            max_links: self.max_links,
            max_string_units: self.max_string_units,
            max_total_utf16_units: self.max_total_utf16_units,
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
    pub fn new(value: impl Into<String>) -> Result<Self, OleError> {
        let value = value.into();
        validate_text(&value, "link base")?;
        Ok(Self { value })
    }

    /// Borrows the stored base text.
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
    pub fn new(
        app: i32,
        office_art: i32,
        info: i32,
        target: impl Into<String>,
        location: impl Into<String>,
    ) -> Result<Self, OleError> {
        let target = target.into();
        let location = location.into();
        validate_text(&target, "hyperlink target")?;
        validate_text(&location, "hyperlink location")?;
        Ok(Self {
            stored_hash: hyperlink_hash(&target, &location),
            app,
            office_art,
            info,
            target,
            location,
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
    pub const fn stored_hash(&self) -> u32 {
        self.stored_hash
    }

    /// Calculates the MS-OSHARED 2.4.2 hash from the inert target and location.
    pub fn calculated_hash(&self) -> u32 {
        hyperlink_hash(&self.target, &self.location)
    }

    /// Checks whether the retained stored hash matches the calculated SHOULD hash.
    pub fn hash_matches(&self) -> bool {
        self.stored_hash == self.calculated_hash()
    }

    /// Implementation-specific application value.
    pub const fn app(&self) -> i32 {
        self.app
    }

    /// Inert OfficeArt shape identifier, or zero when not shape-bound.
    pub const fn office_art(&self) -> i32 {
        self.office_art
    }

    /// Implementation-specific link state value.
    pub const fn info(&self) -> i32 {
        self.info
    }

    /// The inert hyperlink target text.
    pub fn target(&self) -> &str {
        &self.target
    }

    /// The inert hyperlink location text.
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
    pub fn new(links: Vec<Hyperlink>) -> Self {
        Self { links }
    }

    /// Borrows all records in their stored order.
    pub fn links(&self) -> &[Hyperlink] {
        &self.links
    }

    /// Iterates records in their stored order.
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Hyperlink> {
        self.links.iter()
    }

    /// Returns the number of records.
    pub fn len(&self) -> usize {
        self.links.len()
    }

    /// Whether the list contains no records.
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
    pub fn new(section: &'a Section) -> Result<Self, OleError> {
        Self::with_limits(section, Limits::default())
    }

    /// Creates a lazy view using caller-supplied validated limits.
    pub fn with_limits(section: &'a Section, limits: Limits) -> Result<Self, OleError> {
        validate_section(section)?;
        Ok(Self { section, limits })
    }

    /// Borrows the complete section, including unknown properties.
    pub const fn section(&self) -> &'a Section {
        self.section
    }

    /// Returns the limits used for subsequent lazy reads.
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Reads `_PID_LINKBASE` only when it is accessed.
    pub fn link_base(&self) -> Result<Option<LinkBase>, OleError> {
        self.named_blob(LINK_BASE)?
            .map(|value| codec::decode_link_base(value, self.limits))
            .transpose()
    }

    /// Reads `_PID_HLINKS` only when it is accessed.
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
    pub fn new(section: &'a mut Section) -> Result<Self, OleError> {
        Self::with_limits(section, Limits::default())
    }

    /// Creates an editor using caller-supplied validated limits.
    pub fn with_limits(section: &'a mut Section, limits: Limits) -> Result<Self, OleError> {
        validate_edit_section(section)?;
        Ok(Self { section, limits })
    }

    /// Borrows the complete editable section.
    pub const fn section(&self) -> &Section {
        self.section
    }

    /// Sets `_PID_LINKBASE`, preserving an existing PID/name spelling/order.
    pub fn set_link_base(&mut self, value: LinkBase) -> Result<(), OleError> {
        let value = Value::Blob(codec::encode_link_base(&value, self.limits)?);
        self.set_named(LINK_BASE, value)
    }

    /// Removes `_PID_LINKBASE` if present.
    pub fn remove_link_base(&mut self) -> bool {
        self.remove_named(LINK_BASE)
    }

    /// Sets `_PID_HLINKS`, preserving an existing PID/name spelling/order.
    pub fn set_hyperlinks(&mut self, value: Hyperlinks) -> Result<(), OleError> {
        let value = Value::Blob(codec::encode_hyperlinks(&value, self.limits)?);
        self.set_named(HYPERLINKS, value)
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
        if let Some((identifier, _)) = section.find_named(name) {
            if identifier > MAX_NAMED_PROPERTY_ID {
                return Err(invalid(format!(
                    "{name} property identifier {identifier} exceeds the UserDefinedProperties effective maximum"
                )));
            }
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
        let mut original = [0; 2];
        let original = character.encode_utf16(&mut original);
        if original.len() > remaining {
            // MS-OSHARED counts WCHARs. A surrogate cut at the exact boundary
            // cannot be lowercased as a scalar, so retain that code unit.
            hash_unit(&mut hash, &mut first, original[0]);
            break;
        }
        remaining -= original.len();
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
