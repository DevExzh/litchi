//! `PowerPoint` 9 picture-bullet collection parsing.

use super::package::{Error, RecordLimits, Result};
use super::records::{Record, RecordParseSession};
use crate::consts::RecordType;

/// The highest `BlipEntityAtom` instance allowed by MS-PPT.
pub const MAX_PICTURE_BULLET_INDEX: u16 = 0x80;

/// Finite limits for picture-bullet parsing and serialization.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PictureBulletLimits {
    /// Maximum encoded `BlipCollection9Container` payload.
    pub max_collection_bytes: usize,
    /// Maximum number of `BlipEntityAtom` children.
    pub max_bullets: usize,
    /// Maximum encoded OfficeArt BLIP/FBSE record per child.
    pub max_officeart_record_bytes: usize,
    /// Maximum number of discoverable `TextPFException9` references.
    pub max_references: usize,
    /// Shared limits for the strict PPT child-record parser.
    pub records: RecordLimits,
}

impl Default for PictureBulletLimits {
    fn default() -> Self {
        Self {
            max_collection_bytes: 256 * 1024 * 1024 - 8,
            max_bullets: usize::from(MAX_PICTURE_BULLET_INDEX) + 1,
            max_officeart_record_bytes: 256 * 1024 * 1024 - 18,
            max_references: 1_000_000,
            records: RecordLimits::default(),
        }
    }
}

impl PictureBulletLimits {
    fn validate(self) -> Result<()> {
        if self.max_collection_bytes == 0
            || self.max_bullets == 0
            || self.max_bullets > usize::from(MAX_PICTURE_BULLET_INDEX) + 1
            || self.max_officeart_record_bytes < 8
            || self.max_officeart_record_bytes > self.max_collection_bytes
            || self.max_references == 0
            || self.records.max_input_bytes == 0
            || self.records.max_records == 0
            || self.records.max_record_bytes < 8
            || self.records.max_record_payload_bytes < 2
            || self.records.max_copied_payload_bytes < 2
        {
            return Err(Error::ResourceLimit(
                "picture-bullet limits are empty or exceed the MS-PPT envelope".to_string(),
            ));
        }
        Ok(())
    }

    fn collection_record_limits(self) -> RecordLimits {
        RecordLimits {
            max_input_bytes: self.records.max_input_bytes.min(self.max_collection_bytes),
            max_records: self.records.max_records.min(self.max_bullets),
            max_record_bytes: self.records.max_record_bytes.min(
                self.max_officeart_record_bytes
                    .checked_add(10)
                    .unwrap_or(usize::MAX),
            ),
            max_record_payload_bytes: self.records.max_record_payload_bytes.min(
                self.max_officeart_record_bytes
                    .checked_add(2)
                    .unwrap_or(usize::MAX),
            ),
            max_copied_payload_bytes: self
                .records
                .max_copied_payload_bytes
                .min(self.max_collection_bytes),
            ..self.records
        }
    }
}

/// Preferred native picture format for a `PowerPoint` 9 bullet.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum PictureBulletType {
    /// Windows Enhanced Metafile.
    Emf = 0x02,
    /// Windows Metafile.
    Wmf = 0x03,
    /// JPEG image.
    Jpeg = 0x05,
    /// PNG image.
    Png = 0x06,
}

impl TryFrom<u8> for PictureBulletType {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self> {
        match value {
            0x02 => Ok(Self::Emf),
            0x03 => Ok(Self::Wmf),
            0x05 => Ok(Self::Jpeg),
            0x06 => Ok(Self::Png),
            _ => Err(Error::Corrupted(
                "BlipEntityAtom has an invalid winBlipType".to_string(),
            )),
        }
    }
}

impl PictureBulletType {
    const fn kind(self) -> litchi_odraw::image::Kind {
        match self {
            Self::Emf => litchi_odraw::image::Kind::Emf,
            Self::Wmf => litchi_odraw::image::Kind::Wmf,
            Self::Jpeg => litchi_odraw::image::Kind::Jpeg,
            Self::Png => litchi_odraw::image::Kind::Png,
        }
    }
}

/// One picture bullet from a `BlipEntityAtom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PictureBullet {
    /// Zero-based bullet index referenced by `TextPFException9`.
    pub index: u16,
    /// Preferred native picture type.
    pub picture_type: PictureBulletType,
    /// Undefined byte preserved from the atom.
    pub unused: u8,
    /// Complete embedded `OfficeArt` BLIP or FBSE record, including its header.
    pub officeart_record: Vec<u8>,
}

impl PictureBullet {
    /// Parses the embedded `OfficeArt` image without copying its file data.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn blip(&self) -> Result<litchi_odraw::image::Blip<'_>> {
        self.blip_with_delay(None)
    }

    /// Parses the image, optionally resolving a delay-loaded FBSE.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn blip_with_delay<'data>(
        &'data self,
        delay: Option<&'data [u8]>,
    ) -> Result<litchi_odraw::image::Blip<'data>> {
        use litchi_odraw::image::{Blip, Context, Delay, Entry};

        let (record, consumed) = litchi_odraw::Record::parse(&self.officeart_record, 0)
            .map_err(|error| Error::Corrupted(format!("Invalid picture-bullet BLIP: {error}")))?;
        if consumed != self.officeart_record.len() {
            return Err(Error::Corrupted(
                "Picture-bullet BLIP was only partially parsed".to_string(),
            ));
        }
        let blip = if record.kind() == litchi_odraw::RecordKind::Bse {
            let entry = Entry::parse(record)?;
            let context = delay.map_or_else(Context::new, |data| {
                Context::new().with_delay(Delay::new(data))
            });
            entry.resolve(context)?.ok_or_else(|| {
                Error::Corrupted("Picture-bullet FBSE is an empty slot".to_string())
            })?
        } else {
            Blip::from_record(record)?
        };
        if !picture_type_matches(self.picture_type, blip.kind()) {
            return Err(Error::Corrupted(
                "Picture-bullet preferred and stored BLIP types disagree".to_string(),
            ));
        }
        Ok(blip)
    }

    /// Validate this bullet and return its complete `BlipEntityAtom` record.
    ///
    /// The embedded OfficeArt bytes are copied verbatim.  A delayed FBSE is
    /// checked through its stored preferred kind; resolving a delayed store
    /// is available separately through [`Self::blip_with_delay`].
    ///
    /// # Errors
    ///
    /// Returns an error if the index, embedded record, or preferred kind is
    /// outside the MS-PPT picture-bullet grammar.
    pub fn to_record(&self) -> Result<Record> {
        self.to_record_with_limits(PictureBulletLimits::default())
    }

    /// Validate and serialize this bullet under explicit finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the bullet exceeds `limits` or its OfficeArt
    /// record is malformed.
    pub fn to_record_with_limits(&self, limits: PictureBulletLimits) -> Result<Record> {
        limits.validate()?;
        validate_bullet_index(self.index, limits)?;
        validate_officeart_record(
            self.picture_type,
            &self.officeart_record,
            limits.max_officeart_record_bytes,
        )?;
        let payload_len =
            self.officeart_record.len().checked_add(2).ok_or_else(|| {
                Error::ResourceLimit("picture-bullet payload size overflow".into())
            })?;
        let data = picture_bullet_payload(self.picture_type, self.unused, &self.officeart_record)?;
        debug_assert_eq!(data.len(), payload_len);
        Ok(Record {
            record_type: RecordType::BlipEntity9Atom,
            record_type_raw: RecordType::BlipEntity9Atom.as_u16(),
            version: 0,
            instance: self.index,
            data_length: u32::try_from(data.len()).map_err(|_err| {
                Error::ResourceLimit("picture-bullet payload exceeds u32".into())
            })?,
            data: data.into(),
            children: Vec::new(),
        })
    }

    /// Validate and serialize this bullet after resolving an optional delay
    /// store for a delayed FBSE.
    ///
    /// # Errors
    ///
    /// Returns an error when the delayed store is absent or cannot resolve a
    /// non-embedded FBSE, or when ordinary bullet validation fails.
    pub fn to_record_with_delay(&self, delay: Option<&[u8]>) -> Result<Record> {
        let record = self.to_record()?;
        self.blip_with_delay(delay).map(|_blip| record)
    }
}

/// Parsed `BlipCollection9Container`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PictureBulletCollection {
    /// Picture bullets in record order.
    pub bullets: Vec<PictureBullet>,
}

impl PictureBulletCollection {
    /// Parse a `BlipCollection9Container` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(record: &Record) -> Result<Self> {
        Self::parse_with_limits(record, PictureBulletLimits::default())
    }

    /// Parse a `BlipCollection9Container` record under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read, is malformed, or exceeds
    /// `limits`.
    pub fn parse_with_limits(record: &Record, limits: PictureBulletLimits) -> Result<Self> {
        limits.validate()?;
        if record.record_type != RecordType::BlipCollection9
            || record.version != 0x0f
            || record.instance != 0
        {
            return Err(Error::Corrupted(
                "BlipCollection9Container has an invalid record header".to_string(),
            ));
        }
        if usize::try_from(record.data_length).ok() != Some(record.data.len()) {
            return Err(Error::Corrupted(
                "BlipCollection9Container length does not match its payload".into(),
            ));
        }
        if record.data.len() > limits.max_collection_bytes {
            return Err(Error::ResourceLimit(
                "picture-bullet collection exceeds its byte limit".to_string(),
            ));
        }
        let mut parser =
            RecordParseSession::new(limits.collection_record_limits(), record.data.len())?;
        let children = parser.parse_sequence(&record.data, "picture-bullet collection", 0)?;
        if children.len() > limits.max_bullets {
            return Err(Error::ResourceLimit(
                "picture-bullet collection exceeds its child limit".to_string(),
            ));
        }
        let mut bullets = Vec::new();
        bullets
            .try_reserve_exact(children.len())
            .map_err(|_err| Error::AllocationFailed("picture-bullet collection"))?;
        for child in children {
            if child.record_type != RecordType::BlipEntity9Atom
                || child.version != 0
                || child.instance > MAX_PICTURE_BULLET_INDEX
            {
                return Err(Error::Corrupted(
                    "Picture-bullet collection has an invalid child record".to_string(),
                ));
            }
            if usize::try_from(child.data_length).ok() != Some(child.data.len()) {
                return Err(Error::Corrupted(
                    "BlipEntityAtom length does not match its payload".into(),
                ));
            }
            if bullets
                .iter()
                .any(|bullet: &PictureBullet| bullet.index == child.instance)
            {
                return Err(Error::Corrupted(
                    "Picture-bullet collection has a duplicate index".to_string(),
                ));
            }
            bullets.push(parse_picture_bullet_with_limits(&child, limits)?);
        }
        Ok(Self { bullets })
    }

    /// Discover the single `PowerPoint` 9 picture-bullet collection below `root`.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_from(root: &Record) -> Result<Option<Self>> {
        Self::parse_from_with_limits(root, PictureBulletLimits::default())
    }

    /// Discover the single `PowerPoint` 9 picture-bullet collection below
    /// `root` under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input is malformed or more than one owner is
    /// present.
    pub fn parse_from_with_limits(
        root: &Record,
        limits: PictureBulletLimits,
    ) -> Result<Option<Self>> {
        limits.validate()?;
        let mut result = None;
        for record in root.versioned_binary_tag_records_with_limits(9, limits.records)? {
            if record.record_type != RecordType::BlipCollection9 {
                continue;
            }
            if result
                .replace(Self::parse_with_limits(&record, limits)?)
                .is_some()
            {
                return Err(Error::Corrupted(
                    "Record tree contains multiple picture-bullet collections".to_string(),
                ));
            }
        }
        Ok(result)
    }

    /// Validate and build the lossless `BlipCollection9Container` record.
    ///
    /// # Errors
    ///
    /// Returns an error if a bullet is malformed, indices are duplicated, or
    /// the collection exceeds `limits`.
    pub fn to_record(&self) -> Result<Record> {
        self.to_record_with_limits(PictureBulletLimits::default())
    }

    /// Validate and build the collection under explicit finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the collection cannot be represented by the
    /// `BlipCollection9Container` grammar.
    pub fn to_record_with_limits(&self, limits: PictureBulletLimits) -> Result<Record> {
        limits.validate()?;
        if self.bullets.len() > limits.max_bullets {
            return Err(Error::ResourceLimit(
                "picture-bullet collection exceeds its child limit".to_string(),
            ));
        }
        let mut children = Vec::new();
        children
            .try_reserve_exact(self.bullets.len())
            .map_err(|_err| Error::AllocationFailed("picture-bullet child records"))?;
        let mut seen = [false; 129];
        let mut payload_size = 0usize;
        for bullet in &self.bullets {
            let index = usize::from(bullet.index);
            validate_bullet_index(bullet.index, limits)?;
            if seen[index] {
                return Err(Error::Corrupted(
                    "picture-bullet collection has a duplicate index".to_string(),
                ));
            }
            seen[index] = true;
            let child = bullet.to_record_with_limits(limits)?;
            payload_size = payload_size
                .checked_add(8)
                .and_then(|value| value.checked_add(child.data.len()))
                .ok_or_else(|| {
                    Error::ResourceLimit("picture-bullet collection size overflow".into())
                })?;
            if payload_size > limits.max_collection_bytes {
                return Err(Error::ResourceLimit(
                    "picture-bullet collection exceeds its byte limit".to_string(),
                ));
            }
            children.push(child);
        }
        let data = encode_sequence(&children, limits.max_collection_bytes)?;
        if data.len() != payload_size {
            return Err(Error::Corrupted(
                "picture-bullet collection size accounting mismatch".into(),
            ));
        }
        Ok(Record {
            record_type: RecordType::BlipCollection9,
            record_type_raw: RecordType::BlipCollection9.as_u16(),
            version: 0x0f,
            instance: 0,
            data_length: u32::try_from(data.len()).map_err(|_err| {
                Error::ResourceLimit("picture-bullet collection exceeds u32".into())
            })?,
            data: data.into(),
            children,
        })
    }

    /// Return a complete canonical PPT record byte sequence.
    ///
    /// # Errors
    ///
    /// Returns an error if validation or bounded serialization fails.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.to_bytes_with_limits(PictureBulletLimits::default())
    }

    /// Return the collection bytes under explicit finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error if validation or bounded serialization fails.
    pub fn to_bytes_with_limits(&self, limits: PictureBulletLimits) -> Result<Vec<u8>> {
        let record = self.to_record_with_limits(limits)?;
        let maximum = limits
            .max_collection_bytes
            .checked_add(8)
            .ok_or_else(|| Error::ResourceLimit("picture-bullet record size overflow".into()))?;
        encode_record(&record, maximum)
    }

    /// Resolve a `bulletBlipRef`; `-1` is the null reference.
    #[must_use]
    pub fn get(&self, reference: i16) -> Option<&PictureBullet> {
        let index = u16::try_from(reference).ok()?;
        self.bullets.iter().find(|bullet| bullet.index == index)
    }
}

fn parse_picture_bullet_with_limits(
    record: &Record,
    limits: PictureBulletLimits,
) -> Result<PictureBullet> {
    if record.data.len() < 10 {
        return Err(Error::Corrupted("BlipEntityAtom is truncated".to_string()));
    }
    let picture_type = PictureBulletType::try_from(record.data[0])?;
    let unused = record.data[1];
    let officeart_record = &record.data[2..];
    validate_officeart_record(
        picture_type,
        officeart_record,
        limits.max_officeart_record_bytes,
    )?;

    let mut owned_officeart = Vec::new();
    owned_officeart
        .try_reserve_exact(officeart_record.len())
        .map_err(|_err| Error::AllocationFailed("picture-bullet OfficeArt record"))?;
    owned_officeart.extend_from_slice(officeart_record);

    Ok(PictureBullet {
        index: record.instance,
        picture_type,
        unused,
        officeart_record: owned_officeart,
    })
}

fn validate_bullet_index(index: u16, _limits: PictureBulletLimits) -> Result<()> {
    if index > MAX_PICTURE_BULLET_INDEX {
        return Err(Error::InvalidFormat(format!(
            "picture-bullet index {index} exceeds the supported range"
        )));
    }
    Ok(())
}

fn validate_officeart_record(
    picture_type: PictureBulletType,
    officeart_record: &[u8],
    maximum: usize,
) -> Result<()> {
    if officeart_record.len() < 8 || officeart_record.len() > maximum {
        return Err(if officeart_record.len() > maximum {
            Error::ResourceLimit("picture-bullet OfficeArt record exceeds its byte limit".into())
        } else {
            Error::Corrupted("picture-bullet OfficeArt record is truncated".into())
        });
    }
    let (image_record, consumed) = litchi_odraw::Record::parse(officeart_record, 0)?;
    if consumed != officeart_record.len() {
        return Err(Error::Corrupted(
            "Picture-bullet OfficeArt record has an invalid size".to_string(),
        ));
    }
    let kind = if image_record.kind() == litchi_odraw::RecordKind::Bse {
        litchi_odraw::image::Entry::parse(image_record)?.kind()?
    } else {
        litchi_odraw::image::Blip::from_record(image_record)?.kind()
    };
    if !picture_type_matches(picture_type, kind) {
        return Err(Error::Corrupted(
            "Picture-bullet preferred and stored BLIP types disagree".to_string(),
        ));
    }
    Ok(())
}

fn picture_bullet_payload(
    picture_type: PictureBulletType,
    unused: u8,
    officeart_record: &[u8],
) -> Result<Vec<u8>> {
    let mut data = Vec::new();
    data.try_reserve_exact(
        officeart_record
            .len()
            .checked_add(2)
            .ok_or_else(|| Error::ResourceLimit("picture-bullet payload size overflow".into()))?,
    )
    .map_err(|_err| Error::AllocationFailed("picture-bullet payload"))?;
    data.push(picture_type as u8);
    data.push(unused);
    data.extend_from_slice(officeart_record);
    Ok(data)
}

fn encode_sequence(records: &[Record], maximum: usize) -> Result<Vec<u8>> {
    let size = records.iter().try_fold(0usize, |total, record| {
        total
            .checked_add(encoded_size(record, maximum)?)
            .ok_or_else(|| Error::ResourceLimit("picture-bullet sequence overflow".into()))
    })?;
    if size > maximum {
        return Err(Error::ResourceLimit(
            "picture-bullet sequence exceeds its byte limit".into(),
        ));
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(size)
        .map_err(|_err| Error::AllocationFailed("picture-bullet sequence"))?;
    for record in records {
        bytes.extend_from_slice(&encode_record(record, maximum)?);
    }
    Ok(bytes)
}

fn encoded_size(record: &Record, maximum: usize) -> Result<usize> {
    let payload = if record.children.is_empty() {
        record.data.len()
    } else {
        record.children.iter().try_fold(0usize, |total, child| {
            total
                .checked_add(encoded_size(child, maximum)?)
                .ok_or_else(|| Error::ResourceLimit("picture-bullet sequence overflow".into()))
        })?
    };
    let total = payload
        .checked_add(8)
        .ok_or_else(|| Error::ResourceLimit("picture-bullet record size overflow".into()))?;
    if total > maximum {
        return Err(Error::ResourceLimit(
            "picture-bullet record exceeds its byte limit".into(),
        ));
    }
    Ok(total)
}

fn encode_record(record: &Record, maximum: usize) -> Result<Vec<u8>> {
    let payload = if record.children.is_empty() {
        record.data.as_slice()
    } else {
        &encode_sequence(&record.children, maximum)?
    };
    let total = payload
        .len()
        .checked_add(8)
        .ok_or_else(|| Error::ResourceLimit("picture-bullet record size overflow".into()))?;
    if total > maximum {
        return Err(Error::ResourceLimit(
            "picture-bullet record exceeds its byte limit".into(),
        ));
    }
    let length = u32::try_from(payload.len())
        .map_err(|_err| Error::ResourceLimit("picture-bullet record exceeds u32".into()))?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(total)
        .map_err(|_err| Error::AllocationFailed("picture-bullet record"))?;
    bytes.extend_from_slice(&((record.instance << 4) | (record.version & 0x0f)).to_le_bytes());
    bytes.extend_from_slice(&record.record_type_raw.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(payload);
    Ok(bytes)
}

fn picture_type_matches(
    picture_type: PictureBulletType,
    stored_type: litchi_odraw::image::Kind,
) -> bool {
    stored_type == picture_type.kind()
        || (picture_type == PictureBulletType::Jpeg
            && stored_type == litchi_odraw::image::Kind::CmykJpeg)
}

fn is_pp9_name(record: &Record) -> bool {
    if record.record_type != RecordType::CString {
        return false;
    }
    let units = record
        .data
        .as_chunks::<2>()
        .0
        .iter()
        .map(|bytes| u16::from_le_bytes(*bytes));
    units.eq("___PPT9".encode_utf16())
}

fn collect_picture_bullet_references(
    root: &Record,
    limits: PictureBulletLimits,
) -> Result<Vec<u16>> {
    limits.validate()?;
    let mut references = Vec::new();
    collect_direct_references(root, &mut references, limits.max_references)?;
    for record in root.versioned_binary_tag_records_with_limits(9, limits.records)? {
        collect_atom_references(&record, &mut references, limits.max_references)?;
    }
    Ok(references)
}

fn collect_source_picture_bullet_references(
    source: &PictureBulletSnapshot,
    document: &Record,
    collection: &PictureBulletCollection,
    limits: PictureBulletLimits,
) -> Result<Vec<u16>> {
    let mut references = collect_picture_bullet_references(document, limits)?;
    validate_admitted_picture_bullet_references(&references, collection)?;
    let mut package = crate::Package::from_reader(std::io::Cursor::new(source.bytes()))?;
    let presentation = package.presentation_with_limits(limits.records)?;
    let shape_limits = crate::ShapeProgrammableTagLimits {
        max_style_payload_bytes: limits.max_collection_bytes,
        max_style_runs: limits.max_references,
        ..crate::ShapeProgrammableTagLimits::default()
    };
    use crate::odraw::ShapeExt;
    for record in presentation
        .parser
        .try_find_records_ref()?
        .into_iter()
        .filter(|record| record.record_type == RecordType::PPDrawing)
    {
        let drawing = crate::odraw::parse(&record.data)?;
        for shape in &drawing {
            if let Some(tags) = shape.programmable_tags_with_limits(shape_limits)?
                && let Some(style) = tags.powerpoint9()
            {
                for run in &style.runs {
                    if let Some(reference) = run.paragraph.bullet_blip_ref {
                        if reference == -1 {
                            continue;
                        }
                        if references.len() >= limits.max_references {
                            return Err(Error::ResourceLimit(
                                "picture-bullet reference table exceeds its limit".into(),
                            ));
                        }
                        references
                            .try_reserve(1)
                            .map_err(|_err| Error::AllocationFailed("picture-bullet references"))?;
                        references.push(validate_picture_bullet_reference(reference)?);
                    }
                }
            }
        }
    }
    validate_admitted_picture_bullet_references(&references, collection)?;
    Ok(references)
}

fn validate_admitted_picture_bullet_references(
    references: &[u16],
    collection: &PictureBulletCollection,
) -> Result<()> {
    for &reference in references {
        if reference > MAX_PICTURE_BULLET_INDEX
            || !collection
                .bullets
                .iter()
                .any(|bullet| bullet.index == reference)
        {
            return Err(Error::InvalidFormat(format!(
                "TextPFException9 references missing picture-bullet index {reference}"
            )));
        }
    }
    Ok(())
}

fn validate_picture_bullet_reference(reference: i16) -> Result<u16> {
    let reference = u16::try_from(reference).map_err(|_err| {
        Error::Corrupted("TextPFException9 contains an invalid bullet reference".into())
    })?;
    if reference > MAX_PICTURE_BULLET_INDEX {
        return Err(Error::InvalidFormat(format!(
            "TextPFException9 references unsupported picture-bullet index {reference}"
        )));
    }
    Ok(reference)
}

fn collect_direct_references(
    record: &Record,
    references: &mut Vec<u16>,
    maximum: usize,
) -> Result<()> {
    collect_atom_references(record, references, maximum)?;
    for child in &record.children {
        collect_direct_references(child, references, maximum)?;
    }
    Ok(())
}

fn collect_atom_references(
    record: &Record,
    references: &mut Vec<u16>,
    maximum: usize,
) -> Result<()> {
    let mut push = |reference: Option<i16>| -> Result<()> {
        let Some(reference) = reference else {
            return Ok(());
        };
        if reference == -1 {
            return Ok(());
        }
        if references.len() >= maximum {
            return Err(Error::ResourceLimit(
                "picture-bullet reference table exceeds its limit".into(),
            ));
        }
        references
            .try_reserve(1)
            .map_err(|_err| Error::AllocationFailed("picture-bullet references"))?;
        references.push(validate_picture_bullet_reference(reference)?);
        Ok(())
    };
    match record.record_type {
        RecordType::StyleTextProp9Atom => {
            let style = crate::text_extensions::TextStyleExtension9::parse(&record.data)?;
            for run in style.runs {
                push(run.paragraph.bullet_blip_ref)?;
            }
        },
        RecordType::TextMasterStyle9Atom => {
            let style = crate::text_extensions::TextMasterStyleExtension9::parse(
                &record.data,
                record.instance,
            )?;
            for level in style.levels {
                push(level.paragraph.bullet_blip_ref)?;
            }
        },
        RecordType::TextDefaults9Atom => {
            let defaults = crate::text_extensions::TextDefaultsExtension9::parse(&record.data)?;
            push(defaults.paragraph.bullet_blip_ref)?;
        },
        _ => {},
    }
    Ok(())
}

fn rewrite_picture_bullet_collection(
    root: &mut Record,
    replacement: &Record,
    limits: PictureBulletLimits,
) -> Result<()> {
    let mut matches = 0usize;
    rewrite_picture_bullet_records(root, replacement, limits, &mut matches)?;
    match matches {
        1 => Ok(()),
        0 => Err(Error::InvalidFormat(
            "picture-bullet collection is absent; owner creation is not losslessly provable".into(),
        )),
        _ => Err(Error::Corrupted(
            "document contains multiple picture-bullet collections".into(),
        )),
    }
}

fn rewrite_picture_bullet_records(
    record: &mut Record,
    replacement: &Record,
    limits: PictureBulletLimits,
    matches: &mut usize,
) -> Result<()> {
    if record.record_type == RecordType::ProgBinaryTag {
        let Some(name) = record.children.first() else {
            return Err(Error::Corrupted(
                "ProgBinaryTag is missing its CString owner".into(),
            ));
        };
        if is_pp9_name(name) {
            let mut blob_position = None;
            for (position, child) in record.children.iter().enumerate() {
                if child.record_type != RecordType::BinaryTagData {
                    continue;
                }
                if blob_position.replace(position).is_some() {
                    return Err(Error::Corrupted(
                        "___PPT9 binary tag does not have exactly one BinaryTagData".into(),
                    ));
                }
            }
            let Some(blob_position) = blob_position else {
                return Err(Error::Corrupted(
                    "___PPT9 binary tag does not have exactly one BinaryTagData".into(),
                ));
            };
            rewrite_picture_bullet_blob(
                &mut record.children[blob_position],
                replacement,
                limits,
                matches,
            )?;
        }
    }
    for child in &mut record.children {
        rewrite_picture_bullet_records(child, replacement, limits, matches)?;
    }
    Ok(())
}

fn rewrite_picture_bullet_blob(
    blob: &mut Record,
    replacement: &Record,
    limits: PictureBulletLimits,
    matches: &mut usize,
) -> Result<()> {
    let mut parser = RecordParseSession::new(limits.records, blob.data.len())?;
    let records = parser.parse_sequence(&blob.data, "___PPT9 BinaryTagData", 0)?;
    if encode_sequence(&records, limits.max_collection_bytes)?.as_slice() != blob.data.as_slice() {
        return Err(Error::InvalidFormat(
            "___PPT9 BinaryTagData is not losslessly representable".into(),
        ));
    }
    let mut rewritten = Vec::new();
    rewritten
        .try_reserve_exact(records.len())
        .map_err(|_err| Error::AllocationFailed("picture-bullet tag records"))?;
    let mut found = false;
    for record in records {
        if record.record_type == RecordType::BlipCollection9 {
            if found {
                return Err(Error::Corrupted(
                    "___PPT9 contains multiple picture-bullet collections".into(),
                ));
            }
            found = true;
            *matches = matches.checked_add(1).ok_or_else(|| {
                Error::ResourceLimit("picture-bullet owner count overflow".into())
            })?;
            rewritten.push(clone_record_checked(replacement)?);
        } else {
            rewritten.push(record);
        }
    }
    if found {
        let data = encode_sequence(&rewritten, limits.max_collection_bytes)?;
        blob.data_length = u32::try_from(data.len())
            .map_err(|_err| Error::ResourceLimit("___PPT9 tag exceeds u32".into()))?;
        blob.data = data.into();
    }
    Ok(())
}

fn clone_record_checked(record: &Record) -> Result<Record> {
    let mut data = Vec::new();
    data.try_reserve_exact(record.data.len())
        .map_err(|_err| Error::AllocationFailed("picture-bullet record clone"))?;
    data.extend_from_slice(&record.data);
    let mut children = Vec::new();
    children
        .try_reserve_exact(record.children.len())
        .map_err(|_err| Error::AllocationFailed("picture-bullet child-record clone"))?;
    for child in &record.children {
        children.push(clone_record_checked(child)?);
    }
    Ok(Record {
        record_type: record.record_type,
        record_type_raw: record.record_type_raw,
        version: record.version,
        instance: record.instance,
        data_length: record.data_length,
        data: data.into(),
        children,
    })
}

/// Whole-CFB limits captured by a picture-bullet snapshot and its patches.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PictureBulletPackageLimits {
    /// Maximum source and published CFB size.
    pub max_source_bytes: usize,
    /// Semantic picture-bullet limits.
    pub picture_bullets: PictureBulletLimits,
}

impl Default for PictureBulletPackageLimits {
    fn default() -> Self {
        Self {
            max_source_bytes: 512 * 1024 * 1024,
            picture_bullets: PictureBulletLimits::default(),
        }
    }
}

/// Compact revision token for exact-source picture-bullet patches.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PictureBulletRevision(u64);

impl PictureBulletRevision {
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }

    fn from_bytes(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }
}

/// One staged picture-bullet mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PictureBulletChangeKind {
    /// Replaced an existing bullet without changing its index.
    Replace,
    /// Added a new bullet at an explicit unused index.
    Append,
    /// Removed an unreferenced bullet.
    Remove,
}

/// Compact mutation descriptor.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PictureBulletChange {
    kind: PictureBulletChangeKind,
    index: u16,
}

impl PictureBulletChange {
    #[must_use]
    pub const fn kind(self) -> PictureBulletChangeKind {
        self.kind
    }

    #[must_use]
    pub const fn index(self) -> u16 {
        self.index
    }
}

/// Immutable whole-CFB picture-bullet snapshot bound to the live document owner.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PictureBulletSnapshot {
    bytes: std::sync::Arc<[u8]>,
    document: std::sync::Arc<[u8]>,
    document_record: std::sync::Arc<Record>,
    picture_bullets: Option<std::sync::Arc<PictureBulletCollection>>,
    document_persist_id: u32,
    limits: PictureBulletPackageLimits,
}

struct AdmittedPictureBulletSource {
    document_persist_id: u32,
    document: Vec<u8>,
    document_record: Record,
    picture_bullets: Option<PictureBulletCollection>,
}

impl PictureBulletSnapshot {
    /// Snapshot a borrowed CFB source under default limits.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::parse_with_limits(bytes, PictureBulletPackageLimits::default())
    }

    /// Snapshot a borrowed CFB source under explicit limits.
    pub fn parse_with_limits(bytes: &[u8], limits: PictureBulletPackageLimits) -> Result<Self> {
        validate_picture_bullet_source_len(bytes.len(), limits)?;
        // Admit the borrowed bytes first. The owned CFB copy is retained only
        // after the live owner, record tree, and picture-bullet limits pass,
        // and the admitted state is transferred without reparsing.
        let admitted = admit_picture_bullet_source(bytes, limits)?;
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(bytes.len())
            .map_err(|_err| Error::AllocationFailed("picture-bullet snapshot source"))?;
        owned.extend_from_slice(bytes);
        Ok(Self::from_admitted(
            std::sync::Arc::from(owned),
            limits,
            admitted,
        ))
    }

    /// Snapshot an owned CFB source under default limits.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, PictureBulletPackageLimits::default())
    }

    /// Snapshot an owned CFB source under explicit limits.
    pub fn from_bytes_with_limits(
        bytes: Vec<u8>,
        limits: PictureBulletPackageLimits,
    ) -> Result<Self> {
        validate_picture_bullet_source_len(bytes.len(), limits)?;
        Self::from_arc(std::sync::Arc::from(bytes), limits)
    }

    fn from_arc(bytes: std::sync::Arc<[u8]>, limits: PictureBulletPackageLimits) -> Result<Self> {
        validate_picture_bullet_source_len(bytes.len(), limits)?;
        let admitted = admit_picture_bullet_source(&bytes, limits)?;
        Ok(Self::from_admitted(bytes, limits, admitted))
    }

    fn from_admitted(
        bytes: std::sync::Arc<[u8]>,
        limits: PictureBulletPackageLimits,
        admitted: AdmittedPictureBulletSource,
    ) -> Self {
        let picture_bullets = admitted.picture_bullets.map(std::sync::Arc::new);
        Self {
            bytes,
            document: std::sync::Arc::from(admitted.document),
            document_record: std::sync::Arc::new(admitted.document_record),
            picture_bullets,
            document_persist_id: admitted.document_persist_id,
            limits,
        }
    }
}

fn admit_picture_bullet_source(
    bytes: &[u8],
    limits: PictureBulletPackageLimits,
) -> Result<AdmittedPictureBulletSource> {
    limits.picture_bullets.validate()?;
    let (document_persist_id, source) =
        crate::embedded::object::Editor::inspect_live_document(bytes)?;
    let record_limits = limits.picture_bullets.records;
    let (record, consumed) = Record::parse_strict_with_limits(&source, 0, record_limits)?;
    if consumed != source.len() || record.record_type != RecordType::Document {
        return Err(Error::Corrupted(
            "live document persist owner is not one complete DocumentContainer".into(),
        ));
    }
    let encoded = encode_record(&record, record_limits.max_record_bytes)?;
    if encoded != source.as_slice() {
        return Err(Error::InvalidFormat(
            "live picture-bullet owner is not losslessly representable".into(),
        ));
    }
    let picture_bullets =
        PictureBulletCollection::parse_from_with_limits(&record, limits.picture_bullets)?;
    Ok(AdmittedPictureBulletSource {
        document_persist_id,
        document: source,
        document_record: record,
        picture_bullets,
    })
}

impl PictureBulletSnapshot {
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    #[must_use]
    pub fn document_bytes(&self) -> &[u8] {
        &self.document
    }

    #[must_use]
    pub fn picture_bullets(&self) -> Option<&PictureBulletCollection> {
        self.picture_bullets.as_deref()
    }

    #[must_use]
    pub const fn document_persist_id(&self) -> u32 {
        self.document_persist_id
    }

    #[must_use]
    pub const fn limits(&self) -> PictureBulletPackageLimits {
        self.limits
    }

    #[must_use]
    pub fn revision(&self) -> PictureBulletRevision {
        PictureBulletRevision::from_bytes(&self.bytes)
    }

    /// Start an isolated picture-bullet transaction.
    pub fn edit(&self) -> Result<PictureBulletTransaction> {
        Ok(PictureBulletTransaction::new(self.clone()))
    }
}

/// Reversible, source-bound picture-bullet patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PictureBulletPatch {
    base: PictureBulletRevision,
    target: PictureBulletRevision,
    before: std::sync::Arc<[u8]>,
    after: std::sync::Arc<[u8]>,
    changes: Vec<PictureBulletChange>,
    limits: PictureBulletPackageLimits,
}

impl PictureBulletPatch {
    #[must_use]
    pub const fn base(&self) -> PictureBulletRevision {
        self.base
    }

    #[must_use]
    pub const fn target(&self) -> PictureBulletRevision {
        self.target
    }

    #[must_use]
    pub fn before_bytes(&self) -> &[u8] {
        &self.before
    }

    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        &self.after
    }

    #[must_use]
    pub fn changes(&self) -> &[PictureBulletChange] {
        &self.changes
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before == self.after
    }

    /// Apply this patch to its exact source, accepting an already-applied target.
    pub fn apply(&self, current: &PictureBulletSnapshot) -> Result<PictureBulletSnapshot> {
        if current.revision() == self.target && current.bytes() == self.after.as_ref() {
            return Ok(current.clone());
        }
        if current.revision() != self.base || current.bytes() != self.before.as_ref() {
            return Err(Error::InvalidFormat(
                "cannot apply picture-bullet patch to a different CFB source".into(),
            ));
        }
        PictureBulletSnapshot::from_arc(self.after.clone(), self.limits)
    }

    /// Re-apply this patch.
    pub fn redo(&self, current: &PictureBulletSnapshot) -> Result<PictureBulletSnapshot> {
        self.apply(current)
    }

    /// Undo this patch against its exact target.
    pub fn undo(&self, current: &PictureBulletSnapshot) -> Result<PictureBulletSnapshot> {
        if current.revision() == self.base && current.bytes() == self.before.as_ref() {
            return Ok(current.clone());
        }
        if current.revision() != self.target || current.bytes() != self.after.as_ref() {
            return Err(Error::InvalidFormat(
                "cannot undo picture-bullet patch from a different CFB source".into(),
            ));
        }
        PictureBulletSnapshot::from_arc(self.before.clone(), self.limits)
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        let mut changes = self.changes.clone();
        changes.reverse();
        Self {
            base: self.target,
            target: self.base,
            before: self.after.clone(),
            after: self.before.clone(),
            changes,
            limits: self.limits,
        }
    }
}

/// Published picture-bullet transaction result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PictureBulletCommit {
    snapshot: PictureBulletSnapshot,
    patch: PictureBulletPatch,
}

impl PictureBulletCommit {
    #[must_use]
    pub const fn snapshot(&self) -> &PictureBulletSnapshot {
        &self.snapshot
    }

    #[must_use]
    pub const fn patch(&self) -> &PictureBulletPatch {
        &self.patch
    }

    #[must_use]
    pub fn picture_bullets(&self) -> Option<&PictureBulletCollection> {
        self.snapshot.picture_bullets()
    }

    pub fn undo(&self, current: &PictureBulletSnapshot) -> Result<PictureBulletSnapshot> {
        self.patch.undo(current)
    }

    pub fn redo(&self, current: &PictureBulletSnapshot) -> Result<PictureBulletSnapshot> {
        self.patch.redo(current)
    }
}

/// Isolated transaction for source-bound picture-bullet lifecycle edits.
#[derive(Debug, Clone)]
pub struct PictureBulletTransaction {
    source: PictureBulletSnapshot,
    document: std::sync::Arc<Record>,
    picture_bullets: Option<std::sync::Arc<PictureBulletCollection>>,
    changes: Vec<PictureBulletChange>,
}

/// Picture-bullet package limits under the conventional owner-module name.
pub type PackageLimits = PictureBulletPackageLimits;
/// Picture-bullet snapshot under the conventional owner-module name.
pub type Snapshot = PictureBulletSnapshot;
/// Picture-bullet patch under the conventional owner-module name.
pub type Patch = PictureBulletPatch;
/// Picture-bullet commit under the conventional owner-module name.
pub type Commit = PictureBulletCommit;
/// Picture-bullet transaction under the conventional owner-module name.
pub type Transaction = PictureBulletTransaction;
/// Picture-bullet change under the conventional owner-module name.
pub type Change = PictureBulletChange;
/// Picture-bullet change kind under the conventional owner-module name.
pub type ChangeKind = PictureBulletChangeKind;
/// Picture-bullet revision under the conventional owner-module name.
pub type Revision = PictureBulletRevision;

impl PictureBulletTransaction {
    fn new(source: PictureBulletSnapshot) -> Self {
        Self {
            picture_bullets: source.picture_bullets.clone(),
            document: source.document_record.clone(),
            source,
            changes: Vec::new(),
        }
    }

    #[must_use]
    pub const fn source(&self) -> &PictureBulletSnapshot {
        &self.source
    }

    #[must_use]
    pub fn picture_bullets(&self) -> Option<&PictureBulletCollection> {
        self.picture_bullets.as_deref()
    }

    #[must_use]
    pub fn changes(&self) -> &[PictureBulletChange] {
        &self.changes
    }

    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.changes.is_empty()
    }

    /// Replace an existing bullet without changing its index.
    pub fn replace_picture_bullet(&mut self, index: u16, bullet: PictureBullet) -> Result<()> {
        if bullet.index != index {
            return Err(Error::InvalidFormat(
                "picture-bullet replacement cannot change its referenced index".into(),
            ));
        }
        let Some(current) = self.picture_bullets.as_ref() else {
            return Err(Error::InvalidFormat(
                "picture-bullet collection is absent; owner creation is not losslessly provable"
                    .into(),
            ));
        };
        let Some(position) = current
            .bullets
            .iter()
            .position(|value| value.index == index)
        else {
            return Err(Error::InvalidFormat(format!(
                "picture-bullet index {index} is absent"
            )));
        };
        if current.bullets[position] == bullet {
            return Ok(());
        }
        let mut candidate = current.as_ref().clone();
        candidate.bullets[position] = bullet;
        candidate.to_record_with_limits(self.source.limits.picture_bullets)?;
        self.picture_bullets = Some(std::sync::Arc::new(candidate));
        self.changes.push(PictureBulletChange {
            kind: PictureBulletChangeKind::Replace,
            index,
        });
        Ok(())
    }

    /// Append a bullet at its explicit, currently unused index.
    pub fn append_picture_bullet(&mut self, bullet: PictureBullet) -> Result<u16> {
        let Some(current) = self.picture_bullets.as_ref() else {
            return Err(Error::InvalidFormat(
                "picture-bullet collection is absent; owner creation is not losslessly provable"
                    .into(),
            ));
        };
        if current
            .bullets
            .iter()
            .any(|value| value.index == bullet.index)
        {
            return Err(Error::InvalidFormat(format!(
                "picture-bullet index {} is already used",
                bullet.index
            )));
        }
        let index = bullet.index;
        let mut candidate = current.as_ref().clone();
        candidate.bullets.push(bullet);
        candidate.to_record_with_limits(self.source.limits.picture_bullets)?;
        self.picture_bullets = Some(std::sync::Arc::new(candidate));
        self.changes.push(PictureBulletChange {
            kind: PictureBulletChangeKind::Append,
            index,
        });
        Ok(index)
    }

    /// Remove a bullet only when every discoverable `TextPFException9` user
    /// proves it is unreferenced.
    pub fn remove_picture_bullet(&mut self, index: u16) -> Result<PictureBullet> {
        let Some(current) = self.picture_bullets.as_ref() else {
            return Err(Error::InvalidFormat(
                "picture-bullet collection is absent; owner creation is not losslessly provable"
                    .into(),
            ));
        };
        let Some(position) = current
            .bullets
            .iter()
            .position(|value| value.index == index)
        else {
            return Err(Error::InvalidFormat(format!(
                "picture-bullet index {index} is absent"
            )));
        };
        let references = collect_source_picture_bullet_references(
            &self.source,
            &self.document,
            current.as_ref(),
            self.source.limits.picture_bullets,
        )?;
        if references.contains(&index) {
            return Err(Error::InvalidFormat(format!(
                "picture-bullet index {index} is referenced by TextPFException9"
            )));
        }
        let mut candidate = current.as_ref().clone();
        let removed = candidate.bullets.remove(position);
        candidate.to_record_with_limits(self.source.limits.picture_bullets)?;
        self.picture_bullets = Some(std::sync::Arc::new(candidate));
        self.changes.push(PictureBulletChange {
            kind: PictureBulletChangeKind::Remove,
            index,
        });
        Ok(removed)
    }

    /// Reordering is refused because it would require an exhaustive reference
    /// rewrite across all shape and master owners. Re-submitting the existing
    /// order is an exact no-op.
    pub fn reorder_picture_bullets(&mut self, order: &[u16]) -> Result<()> {
        let Some(current) = self.picture_bullets.as_ref() else {
            if order.is_empty() {
                return Ok(());
            }
            return Err(Error::InvalidFormat(
                "picture-bullet collection is absent; owner creation is not losslessly provable"
                    .into(),
            ));
        };
        if current
            .bullets
            .iter()
            .map(|bullet| bullet.index)
            .eq(order.iter().copied())
        {
            return Ok(());
        }
        Err(Error::InvalidFormat(
            "picture-bullet reordering requires a complete TextPFException9 reference remap".into(),
        ))
    }

    /// Publish staged edits as an append-only PPT transaction.
    pub fn commit(mut self) -> Result<PictureBulletCommit> {
        if self.changes.is_empty()
            || self.picture_bullets.as_deref() == self.source.picture_bullets.as_deref()
        {
            let revision = self.source.revision();
            return Ok(PictureBulletCommit {
                snapshot: self.source.clone(),
                patch: PictureBulletPatch {
                    base: revision,
                    target: revision,
                    before: self.source.bytes.clone(),
                    after: self.source.bytes.clone(),
                    changes: Vec::new(),
                    limits: self.source.limits,
                },
            });
        }
        crate::font::require_stream_only_cfb(self.source.bytes())?;
        let collection = self.picture_bullets.as_deref().ok_or_else(|| {
            Error::InvalidFormat(
                "picture-bullet collection is absent; owner creation is not losslessly provable"
                    .into(),
            )
        })?;
        let replacement = collection.to_record_with_limits(self.source.limits.picture_bullets)?;
        let mut editor = crate::embedded::object::Editor::open_records_arc_with_limit(
            self.source.bytes.clone(),
            self.source.limits.max_source_bytes,
        )?;
        let live = editor.persisted_record(self.source.document_persist_id)?;
        if live.as_slice() != self.source.document.as_ref() {
            return Err(Error::InvalidFormat(
                "live picture-bullet owner changed during staging".into(),
            ));
        }
        let root = std::sync::Arc::make_mut(&mut self.document);
        rewrite_picture_bullet_collection(root, &replacement, self.source.limits.picture_bullets)?;
        let target_document = encode_record(
            root,
            self.source.limits.picture_bullets.records.max_record_bytes,
        )?;
        if target_document == self.source.document.as_ref() {
            let revision = self.source.revision();
            return Ok(PictureBulletCommit {
                snapshot: self.source.clone(),
                patch: PictureBulletPatch {
                    base: revision,
                    target: revision,
                    before: self.source.bytes.clone(),
                    after: self.source.bytes.clone(),
                    changes: Vec::new(),
                    limits: self.source.limits,
                },
            });
        }
        preflight_picture_bullet_publication(
            self.source.bytes.len(),
            target_document.len(),
            self.source.limits.max_source_bytes,
        )?;
        editor
            .replace_persisted_record(self.source.document_persist_id, target_document.clone())?;
        let bytes = editor.finish()?;
        crate::font::validate_unrelated_streams(self.source.bytes(), &bytes)?;
        let snapshot =
            PictureBulletSnapshot::from_arc(std::sync::Arc::from(bytes), self.source.limits)?;
        if snapshot.document_persist_id != self.source.document_persist_id
            || snapshot.document.as_ref() != target_document
            || snapshot.picture_bullets.as_deref() != Some(collection)
        {
            return Err(Error::Corrupted(
                "published picture-bullet candidate failed semantic reopen".into(),
            ));
        }
        let patch = PictureBulletPatch {
            base: self.source.revision(),
            target: snapshot.revision(),
            before: self.source.bytes,
            after: snapshot.bytes.clone(),
            changes: self.changes,
            limits: snapshot.limits,
        };
        Ok(PictureBulletCommit { snapshot, patch })
    }

    /// Consume the transaction and publish its staged changes.
    pub fn finish(self) -> Result<PictureBulletCommit> {
        self.commit()
    }

    #[must_use]
    pub fn rollback(self) -> PictureBulletSnapshot {
        self.source
    }
}

fn validate_picture_bullet_source_len(
    len: usize,
    limits: PictureBulletPackageLimits,
) -> Result<()> {
    if limits.max_source_bytes == 0 || len > limits.max_source_bytes {
        return Err(Error::ResourceLimit(format!(
            "PowerPoint picture-bullet snapshot source size {len} exceeds limit {}",
            limits.max_source_bytes
        )));
    }
    Ok(())
}

fn preflight_picture_bullet_publication(
    source_bytes: usize,
    document_bytes: usize,
    maximum: usize,
) -> Result<()> {
    let projected = source_bytes
        .checked_add(document_bytes)
        .and_then(|value| value.checked_add(64))
        .ok_or_else(|| Error::ResourceLimit("PowerPoint publication size overflows".into()))?;
    if projected > maximum {
        return Err(Error::ResourceLimit(format!(
            "PowerPoint picture-bullet publication requires at least {projected} bytes, exceeding the {maximum}-byte source limit"
        )));
    }
    Ok(())
}

/// Read the live picture-bullet collection from a cursor-backed package.
impl<R: std::io::Read + std::io::Seek> crate::Package<R> {
    /// Read the live picture-bullet collection under explicit limits.
    pub fn picture_bullets_with_limits(
        &mut self,
        limits: PictureBulletLimits,
    ) -> Result<Option<PictureBulletCollection>> {
        self.presentation()?.picture_bullets_with_limits(limits)
    }

    /// Read the live picture-bullet collection with default limits.
    pub fn picture_bullets(&mut self) -> Result<Option<PictureBulletCollection>> {
        self.picture_bullets_with_limits(PictureBulletLimits::default())
    }
}

/// Read the live picture-bullet collection from a positional package.
impl crate::SourceBackedPackage {
    /// Read the live picture-bullet collection under explicit limits.
    pub fn picture_bullets_with_limits(
        &self,
        limits: PictureBulletLimits,
    ) -> Result<Option<PictureBulletCollection>> {
        self.presentation()?.picture_bullets_with_limits(limits)
    }

    /// Read the live picture-bullet collection with default limits.
    pub fn picture_bullets(&self) -> Result<Option<PictureBulletCollection>> {
        self.picture_bullets_with_limits(PictureBulletLimits::default())
    }
}

/// Read the live picture-bullet collection from the current presentation.
impl crate::Presentation {
    /// Read the live picture-bullet collection with default limits.
    pub fn picture_bullets(&self) -> Result<Option<PictureBulletCollection>> {
        self.picture_bullets_with_limits(PictureBulletLimits::default())
    }

    /// Read the live picture-bullet collection under explicit limits.
    pub fn picture_bullets_with_limits(
        &self,
        mut limits: PictureBulletLimits,
    ) -> Result<Option<PictureBulletCollection>> {
        limits.records = limits.records.constrained_by(self.record_limits);
        PictureBulletCollection::parse_from_with_limits(&self.live_document_record()?, limits)
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
        data.extend_from_slice(payload);
        data
    }

    fn png_bullet(index: u16) -> Vec<u8> {
        let mut blip = Vec::new();
        blip.extend_from_slice(&(0x06e0u16 << 4).to_le_bytes());
        blip.extend_from_slice(&0xf01eu16.to_le_bytes());
        blip.extend_from_slice(&17u32.to_le_bytes());
        blip.extend_from_slice(&[0; 17]);
        let mut payload = vec![0x06, 0x7f];
        payload.extend_from_slice(&blip);
        record_bytes(0, index, 2041, &payload)
    }

    fn fbse_png_bullet(index: u16) -> Vec<u8> {
        let direct = png_bullet(index);
        let blip = &direct[10..];
        let mut fbse = vec![0x06, 0x06];
        fbse.extend_from_slice(&[0; 16]);
        fbse.extend_from_slice(&0xffu16.to_le_bytes());
        fbse.extend_from_slice(&u32::try_from(blip.len()).unwrap().to_le_bytes());
        fbse.extend_from_slice(&1u32.to_le_bytes());
        fbse.extend_from_slice(&u32::MAX.to_le_bytes());
        fbse.extend_from_slice(&[0, 0, 0, 0]);
        fbse.extend_from_slice(blip);
        let mut payload = vec![0x06, 0];
        payload.extend_from_slice(&record_bytes(2, 0x06, 0xf007, &fbse));
        record_bytes(0, index, 2041, &payload)
    }

    fn delayed_png_bullet(index: u16) -> (Vec<u8>, Vec<u8>) {
        let direct = png_bullet(index);
        let blip = direct[10..].to_vec();
        let mut fbse = vec![0x06, 0x06];
        fbse.extend_from_slice(&[0; 16]);
        fbse.extend_from_slice(&0xffu16.to_le_bytes());
        fbse.extend_from_slice(&u32::try_from(blip.len()).unwrap().to_le_bytes());
        fbse.extend_from_slice(&1u32.to_le_bytes());
        fbse.extend_from_slice(&0u32.to_le_bytes());
        fbse.extend_from_slice(&[0, 0, 0, 0]);
        let mut payload = vec![0x06, 0];
        payload.extend_from_slice(&record_bytes(2, 0x06, 0xf007, &fbse));
        (record_bytes(0, index, 2041, &payload), blip)
    }

    fn collection(payload: Vec<u8>) -> Record {
        Record {
            record_type: RecordType::BlipCollection9,
            record_type_raw: 2040,
            version: 0x0f,
            instance: 0,
            data_length: u32::try_from(payload.len()).unwrap(),
            data: payload.into(),
            children: Vec::new(),
        }
    }

    fn collection_bytes(payload: Vec<u8>) -> Vec<u8> {
        let record = collection(payload);
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&((record.instance << 4) | record.version).to_le_bytes());
        bytes.extend_from_slice(&record.record_type_raw.to_le_bytes());
        bytes.extend_from_slice(&record.data_length.to_le_bytes());
        bytes.extend_from_slice(&record.data);
        bytes
    }

    fn prog_tags_record(blob_payload: &[u8]) -> Record {
        let tag_name: Vec<u8> = "___PPT9"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let name = record_bytes(0, 0, 4026, &tag_name);
        let blob = record_bytes(0, 0, 0x138b, blob_payload);
        let mut tag_payload = name;
        tag_payload.extend_from_slice(&blob);
        let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
        Record {
            record_type: RecordType::ProgTags,
            record_type_raw: 0x1388,
            version: 0x0f,
            instance: 0,
            data_length: u32::try_from(tag.len()).unwrap(),
            data: tag.into(),
            children: Vec::new(),
        }
    }

    fn source_with_bullets(referenced: bool) -> Vec<u8> {
        source_with_bullet_reference(referenced.then_some(4))
    }

    fn empty_source() -> Vec<u8> {
        use std::io::Cursor;

        let mut writer = crate::Writer::new();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn source_with_bullet_reference(reference: Option<i16>) -> Vec<u8> {
        let source = empty_source();
        let (persist_id, document_bytes) =
            crate::embedded::object::Editor::inspect_live_document(&source).unwrap();
        let (mut document, consumed) = Record::parse_strict(&document_bytes, 0).unwrap();
        assert_eq!(consumed, document_bytes.len());

        let mut blob = collection_bytes({
            let mut bullets = png_bullet(4);
            bullets.extend_from_slice(&png_bullet(5));
            bullets
        });
        if let Some(reference) = reference {
            let mut style = Vec::new();
            style.extend_from_slice(&0x0080_0000u32.to_le_bytes());
            style.extend_from_slice(&reference.to_le_bytes());
            style.extend_from_slice(&0u32.to_le_bytes());
            style.extend_from_slice(&0u32.to_le_bytes());
            blob.extend_from_slice(&record_bytes(
                0,
                0,
                RecordType::StyleTextProp9Atom.as_u16(),
                &style,
            ));
        }
        document.children.push(prog_tags_record(&blob));
        let replacement = encode_record(
            &document,
            PictureBulletLimits::default().records.max_record_bytes,
        )
        .unwrap();
        let mut editor = crate::embedded::object::Editor::open_records(source).unwrap();
        editor
            .replace_persisted_record(persist_id, replacement)
            .unwrap();
        editor.finish().unwrap()
    }

    fn shape_drawing_with_bullet_reference(reference: i16) -> Vec<u8> {
        use crate::officeart_wire::{ShapeBuilder, record_type, write_atom, write_container};

        let mut style = Vec::new();
        style.extend_from_slice(&0x0080_0000u32.to_le_bytes());
        style.extend_from_slice(&reference.to_le_bytes());
        style.extend_from_slice(&0u32.to_le_bytes());
        style.extend_from_slice(&0u32.to_le_bytes());
        let style_record = record_bytes(0, 0, RecordType::StyleTextProp9Atom.as_u16(), &style);
        let tags = encode_record(
            &prog_tags_record(&style_record),
            PictureBulletLimits::default().records.max_record_bytes,
        )
        .unwrap();

        let mut shape_children = Vec::new();
        ShapeBuilder::new(1, 1)
            .with_flags(0x0A00u32)
            .write(&mut shape_children)
            .unwrap();
        crate::Anchor::full(0, 0, 1, 1)
            .unwrap()
            .write_to(&mut shape_children)
            .unwrap();
        write_container(&mut shape_children, 0, record_type::CLIENT_DATA, &tags).unwrap();

        let mut shape = Vec::new();
        write_container(&mut shape, 0, record_type::SP_CONTAINER, &shape_children).unwrap();

        let mut drawing_body = Vec::new();
        write_atom(&mut drawing_body, 0, 0, record_type::DG, &[0; 8]).unwrap();
        drawing_body.extend_from_slice(&shape);
        let mut drawing = Vec::new();
        write_container(&mut drawing, 0, record_type::DG_CONTAINER, &drawing_body).unwrap();
        drawing
    }

    fn source_with_shape_bullet_reference(reference: i16) -> Vec<u8> {
        let source = source_with_bullets(false);
        let (persist_id, document_bytes) =
            crate::embedded::object::Editor::inspect_live_document(&source).unwrap();
        let (mut document, consumed) = Record::parse_strict(&document_bytes, 0).unwrap();
        assert_eq!(consumed, document_bytes.len());
        let ppdrawing_bytes = record_bytes(
            0x0f,
            0,
            RecordType::PPDrawing.as_u16(),
            &shape_drawing_with_bullet_reference(reference),
        );
        document
            .children
            .push(Record::parse_strict(&ppdrawing_bytes, 0).unwrap().0);
        let replacement = encode_record(
            &document,
            PictureBulletLimits::default().records.max_record_bytes,
        )
        .unwrap();
        let mut editor = crate::embedded::object::Editor::open_records(source).unwrap();
        editor
            .replace_persisted_record(persist_id, replacement)
            .unwrap();
        editor.finish().unwrap()
    }

    fn with_storage(source: Vec<u8>, path: &[&str]) -> Vec<u8> {
        use litchi_cfb::{OleFile, OleWriter};
        use std::io::Cursor;

        let mut source_file = OleFile::open(Cursor::new(source)).unwrap();
        let streams = source_file
            .list_streams()
            .into_iter()
            .map(|stream_path| {
                let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
                let data = source_file.open_stream(&refs).unwrap();
                (stream_path, data)
            })
            .collect::<Vec<_>>();
        let mut writer = OleWriter::new();
        for (stream_path, data) in streams {
            let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
            writer.create_stream(&refs, &data).unwrap();
        }
        writer.create_storage(path).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn with_stream(source: Vec<u8>, path: &[&str], data: &[u8]) -> Vec<u8> {
        use litchi_cfb::{OleFile, OleWriter};
        use std::io::Cursor;

        let mut source_file = OleFile::open(Cursor::new(source)).unwrap();
        let streams = source_file
            .list_streams()
            .into_iter()
            .map(|stream_path| {
                let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
                let stream_data = source_file.open_stream(&refs).unwrap();
                (stream_path, stream_data)
            })
            .collect::<Vec<_>>();
        let mut writer = OleWriter::new();
        for (stream_path, stream_data) in streams {
            let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
            writer.create_stream(&refs, &stream_data).unwrap();
        }
        writer.create_stream(path, data).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn parses_and_resolves_picture_bullets() {
        let mut records = png_bullet(4);
        records.extend_from_slice(&fbse_png_bullet(5));
        let bullets = PictureBulletCollection::parse(&collection(records)).unwrap();
        let bullet = bullets.get(4).unwrap();
        assert_eq!(bullet.index, 4);
        assert_eq!(bullet.picture_type, PictureBulletType::Png);
        assert_eq!(bullet.unused, 0x7f);
        assert_eq!(bullet.officeart_record.len(), 25);
        assert_eq!(bullets.get(5).unwrap().officeart_record[0] & 0x0f, 2);
        assert!(bullets.get(-1).is_none());
    }

    #[test]
    fn discovers_picture_bullets_in_powerpoint_9_tags() {
        let collection = record_bytes(0x0f, 0, 2040, &png_bullet(7));
        let root = Record {
            record_type: RecordType::Document,
            record_type_raw: 1000,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new().into(),
            children: vec![prog_tags_record(&collection)],
        };
        let bullets = PictureBulletCollection::parse_from(&root).unwrap().unwrap();
        assert_eq!(bullets.get(7).unwrap().picture_type, PictureBulletType::Png);
    }

    #[test]
    fn decodes_direct_and_fbse_picture_bullets() {
        let direct = PictureBulletCollection::parse(&collection(png_bullet(1))).unwrap();
        assert_eq!(direct.get(1).unwrap().blip().unwrap().data(), b"");

        let embedded = PictureBulletCollection::parse(&collection(fbse_png_bullet(2))).unwrap();
        assert_eq!(embedded.get(2).unwrap().blip().unwrap().data(), b"");
    }

    #[test]
    fn rejects_malformed_picture_bullet_collections() {
        let mut duplicate = png_bullet(1);
        duplicate.extend_from_slice(&png_bullet(1));
        assert!(PictureBulletCollection::parse(&collection(duplicate)).is_err());

        let mut mismatched = png_bullet(2);
        mismatched[8] = 0x05;
        assert!(PictureBulletCollection::parse(&collection(mismatched)).is_err());

        let mut truncated = png_bullet(3);
        truncated.pop();
        assert!(PictureBulletCollection::parse(&collection(truncated)).is_err());

        let mut invalid_instance = png_bullet(4);
        invalid_instance[10] = 0;
        invalid_instance[11] = 0;
        assert!(PictureBulletCollection::parse(&collection(invalid_instance)).is_err());

        let mut invalid_fbse = fbse_png_bullet(5);
        invalid_fbse[10] = 0;
        assert!(PictureBulletCollection::parse(&collection(invalid_fbse)).is_err());
    }

    #[test]
    fn collection_serialization_is_source_preserving_and_bounded() {
        let source = collection_bytes(png_bullet(0x80));
        let record = Record::parse(&source, 0).unwrap().0;
        let parsed = PictureBulletCollection::parse(&record).unwrap();
        assert_eq!(parsed.to_bytes().unwrap(), source);
        assert_eq!(parsed.to_record().unwrap().data, record.data);
        let mut opaque = parsed.clone();
        let last = opaque.bullets[0].officeart_record.len() - 1;
        opaque.bullets[0].officeart_record[last] = 0xa5;
        assert_eq!(
            PictureBulletCollection::parse(
                &Record::parse(&opaque.to_bytes().unwrap(), 0).unwrap().0
            )
            .unwrap()
            .bullets[0]
                .officeart_record,
            opaque.bullets[0].officeart_record
        );

        let mut limits = PictureBulletLimits {
            max_collection_bytes: source.len() - 1,
            ..PictureBulletLimits::default()
        };
        assert!(matches!(
            PictureBulletCollection::parse_with_limits(&record, limits),
            Err(Error::ResourceLimit(_))
        ));
        limits.max_collection_bytes = PictureBulletLimits::default().max_collection_bytes;
        limits.max_bullets = 1;
        assert!(PictureBulletCollection::parse_with_limits(&record, limits).is_ok());
    }

    #[test]
    fn authored_records_revalidate_preferred_type_and_fbse_shape() {
        let parsed = PictureBulletCollection::parse(&collection(png_bullet(3))).unwrap();
        let mut authored = parsed.clone();
        authored.bullets[0].unused = 0xa5;
        assert_eq!(
            PictureBulletCollection::parse(
                &Record::parse(&authored.to_bytes().unwrap(), 0).unwrap().0
            )
            .unwrap(),
            authored
        );

        authored.bullets[0].picture_type = PictureBulletType::Jpeg;
        assert!(authored.to_record().is_err());
        let (delayed_bullet, delay) = delayed_png_bullet(4);
        let delayed = PictureBulletCollection::parse(&collection(delayed_bullet)).unwrap();
        assert!(delayed.get(4).unwrap().to_record().is_ok());
        assert!(delayed.get(4).unwrap().to_record_with_delay(None).is_err());
        assert!(
            delayed
                .get(4)
                .unwrap()
                .to_record_with_delay(Some(&[]))
                .is_err()
        );
        assert!(
            delayed
                .get(4)
                .unwrap()
                .to_record_with_delay(Some(&delay))
                .is_ok()
        );
    }

    #[test]
    fn null_bullet_references_are_not_treated_as_invalid_indices() {
        let mut style = Vec::new();
        style.extend_from_slice(&0x0080_0000u32.to_le_bytes());
        style.extend_from_slice(&(-1i16).to_le_bytes());
        style.extend_from_slice(&0u32.to_le_bytes());
        style.extend_from_slice(&0u32.to_le_bytes());
        let root = Record {
            record_type: RecordType::Document,
            record_type_raw: RecordType::Document.as_u16(),
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new().into(),
            children: vec![prog_tags_record(&record_bytes(
                0,
                0,
                RecordType::StyleTextProp9Atom.as_u16(),
                &style,
            ))],
        };
        assert!(
            collect_picture_bullet_references(&root, PictureBulletLimits::default())
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn nonnull_bullet_references_are_bounded_and_must_resolve() {
        let source =
            PictureBulletSnapshot::from_bytes(source_with_bullet_reference(Some(6))).unwrap();
        let mut transaction = source.edit().unwrap();
        assert!(matches!(
            transaction.remove_picture_bullet(4),
            Err(Error::InvalidFormat(message)) if message.contains("missing picture-bullet index 6")
        ));
        assert!(transaction.changes().is_empty());

        let source =
            PictureBulletSnapshot::from_bytes(source_with_bullet_reference(Some(0x81))).unwrap();
        let mut transaction = source.edit().unwrap();
        assert!(matches!(
            transaction.remove_picture_bullet(4),
            Err(Error::InvalidFormat(message)) if message.contains("unsupported picture-bullet index 129")
        ));
        assert!(transaction.changes().is_empty());
    }

    #[test]
    fn removal_sees_shape_owned_pp9_references_in_ppdrawing() {
        let source =
            PictureBulletSnapshot::from_bytes(source_with_shape_bullet_reference(4)).unwrap();
        let mut transaction = source.edit().unwrap();
        let result = transaction.remove_picture_bullet(4);
        assert!(matches!(
            result,
            Err(Error::InvalidFormat(message))
                if message.contains("picture-bullet index 4 is referenced")
        ));
        assert!(transaction.changes().is_empty());
    }

    #[test]
    fn empty_reorder_is_a_noop_when_source_has_no_collection() {
        let source = PictureBulletSnapshot::from_bytes(empty_source()).unwrap();
        assert!(source.picture_bullets().is_none());

        let mut transaction = source.edit().unwrap();
        assert!(transaction.reorder_picture_bullets(&[]).is_ok());
        assert!(transaction.changes().is_empty());
        assert!(transaction.reorder_picture_bullets(&[4]).is_err());
        let commit = transaction.commit().unwrap();
        assert!(commit.patch().is_empty());
        assert!(std::sync::Arc::ptr_eq(
            &source.bytes,
            &commit.snapshot.bytes
        ));
    }

    #[test]
    fn source_transaction_replaces_appends_removes_and_reopens_exactly() {
        let bytes = source_with_bullets(false);
        let mut package = crate::Package::from_reader(std::io::Cursor::new(bytes.clone())).unwrap();
        assert_eq!(
            package
                .picture_bullets()
                .unwrap()
                .unwrap()
                .get(4)
                .unwrap()
                .index,
            4
        );
        let source = PictureBulletSnapshot::from_bytes(bytes).unwrap();
        assert_eq!(source.picture_bullets().unwrap().bullets.len(), 2);

        let mut replacement = source.picture_bullets().unwrap().get(4).unwrap().clone();
        replacement.unused = 0x23;
        let mut transaction = source.edit().unwrap();
        transaction.replace_picture_bullet(4, replacement).unwrap();
        let mut appended = source.picture_bullets().unwrap().get(5).unwrap().clone();
        appended.index = 6;
        assert_eq!(transaction.append_picture_bullet(appended).unwrap(), 6);
        assert!(transaction.remove_picture_bullet(4).is_ok());
        let commit = transaction.commit().unwrap();
        assert!(commit.picture_bullets().unwrap().get(4).is_none());
        assert_eq!(commit.picture_bullets().unwrap().get(5).unwrap().index, 5);
        assert_eq!(commit.picture_bullets().unwrap().get(6).unwrap().index, 6);
        assert_eq!(commit.patch().changes().len(), 3);
        assert_eq!(
            commit.patch().apply(&source).unwrap().bytes(),
            commit.snapshot().bytes()
        );
        assert_eq!(
            commit.patch().apply(commit.snapshot()).unwrap().bytes(),
            commit.snapshot().bytes()
        );
        assert_eq!(
            commit.undo(commit.snapshot()).unwrap().bytes(),
            source.bytes()
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.snapshot())
                .unwrap()
                .bytes(),
            source.bytes()
        );
    }

    #[test]
    fn source_transaction_is_atomic_for_references_stale_sources_and_noop() {
        let source = PictureBulletSnapshot::from_bytes(source_with_bullets(true)).unwrap();
        let mut blocked = source.edit().unwrap();
        let before = blocked.picture_bullets().unwrap().clone();
        assert!(blocked.remove_picture_bullet(4).is_err());
        assert_eq!(blocked.picture_bullets().unwrap(), &before);
        assert!(blocked.changes().is_empty());
        assert!(blocked.reorder_picture_bullets(&[5, 4]).is_err());

        let mut identity = source.edit().unwrap();
        assert!(identity.reorder_picture_bullets(&[4, 5]).is_ok());
        assert!(identity.changes().is_empty());
        assert!(identity.commit().unwrap().patch().is_empty());

        let noop = source.edit().unwrap().commit().unwrap();
        assert!(noop.patch().is_empty());
        assert_eq!(noop.snapshot().bytes(), source.bytes());

        let mut left = source.edit().unwrap();
        let mut left_bullet = left.picture_bullets().unwrap().get(5).unwrap().clone();
        left_bullet.unused = 0x31;
        left.replace_picture_bullet(5, left_bullet).unwrap();
        let left = left.commit().unwrap();

        let mut right = source.edit().unwrap();
        let mut right_bullet = right.picture_bullets().unwrap().get(5).unwrap().clone();
        right_bullet.unused = 0x32;
        right.replace_picture_bullet(5, right_bullet).unwrap();
        let right = right.commit().unwrap();
        assert!(left.patch().apply(right.snapshot()).is_err());
        assert!(left.patch().undo(right.snapshot()).is_err());
    }

    #[test]
    fn protected_and_nested_sources_allow_noop_but_refuse_changed_publication() {
        for source in [
            with_stream(source_with_bullets(false), &["_SIGNATURES"], b"signed"),
            with_storage(source_with_bullets(false), &["ObjectPool"]),
        ] {
            let snapshot = PictureBulletSnapshot::from_bytes(source.clone()).unwrap();
            let noop = snapshot.edit().unwrap().commit().unwrap();
            assert!(noop.patch().is_empty());
            assert!(std::sync::Arc::ptr_eq(
                &snapshot.bytes,
                &noop.snapshot.bytes
            ));

            let mut transaction = snapshot.edit().unwrap();
            let mut replacement = transaction
                .picture_bullets()
                .unwrap()
                .get(4)
                .unwrap()
                .clone();
            replacement.unused ^= 1;
            transaction.replace_picture_bullet(4, replacement).unwrap();
            assert!(transaction.commit().is_err());
            assert_eq!(snapshot.bytes(), source.as_slice());
        }
    }

    #[test]
    fn snapshots_and_transactions_share_admitted_picture_bullets_until_write() {
        let source = PictureBulletSnapshot::from_bytes(source_with_bullets(false)).unwrap();
        let clone = source.clone();
        let source_collection = source.picture_bullets.as_ref().unwrap();
        let clone_collection = clone.picture_bullets.as_ref().unwrap();
        assert!(std::sync::Arc::ptr_eq(source_collection, clone_collection));
        assert!(std::ptr::eq(
            source.picture_bullets().unwrap(),
            clone.picture_bullets().unwrap()
        ));

        let mut transaction = source.edit().unwrap();
        assert!(std::sync::Arc::ptr_eq(
            source_collection,
            transaction.picture_bullets.as_ref().unwrap()
        ));
        let mut replacement = transaction
            .picture_bullets()
            .unwrap()
            .get(4)
            .unwrap()
            .clone();
        replacement.unused ^= 1;
        transaction.replace_picture_bullet(4, replacement).unwrap();
        assert!(!std::sync::Arc::ptr_eq(
            source_collection,
            transaction.picture_bullets.as_ref().unwrap()
        ));
        assert!(std::ptr::eq(
            source.picture_bullets().unwrap(),
            transaction.source().picture_bullets().unwrap()
        ));
    }

    #[test]
    fn semantic_reverts_return_the_original_source_and_bytes() {
        let source = PictureBulletSnapshot::from_bytes(source_with_bullets(false)).unwrap();
        let original = source.picture_bullets().unwrap().get(4).unwrap().clone();
        let mut replacement = original.clone();
        replacement.unused ^= 1;

        let mut transaction = source.edit().unwrap();
        transaction.replace_picture_bullet(4, replacement).unwrap();
        transaction.replace_picture_bullet(4, original).unwrap();
        assert_eq!(transaction.changes().len(), 2);
        let commit = transaction.commit().unwrap();

        assert!(commit.patch().is_empty());
        assert!(commit.patch().changes().is_empty());
        assert!(std::sync::Arc::ptr_eq(
            &source.bytes,
            &commit.snapshot.bytes
        ));
        assert!(std::sync::Arc::ptr_eq(
            source.picture_bullets.as_ref().unwrap(),
            commit.snapshot.picture_bullets.as_ref().unwrap()
        ));
    }

    #[test]
    fn remove_and_readd_preserves_order_before_semantic_noop_detection() {
        let source = PictureBulletSnapshot::from_bytes(source_with_bullets(false)).unwrap();
        let mut transaction = source.edit().unwrap();
        let removed = transaction.remove_picture_bullet(5).unwrap();
        transaction.append_picture_bullet(removed).unwrap();
        assert_eq!(
            transaction
                .picture_bullets()
                .unwrap()
                .bullets
                .iter()
                .map(|bullet| bullet.index)
                .collect::<Vec<_>>(),
            vec![4, 5]
        );
        let commit = transaction.commit().unwrap();
        assert!(commit.patch().is_empty());
        assert!(std::sync::Arc::ptr_eq(
            &source.bytes,
            &commit.snapshot.bytes
        ));

        let mut transaction = source.edit().unwrap();
        let removed = transaction.remove_picture_bullet(4).unwrap();
        transaction.append_picture_bullet(removed).unwrap();
        assert_eq!(
            transaction
                .picture_bullets()
                .unwrap()
                .bullets
                .iter()
                .map(|bullet| bullet.index)
                .collect::<Vec<_>>(),
            vec![5, 4]
        );
        assert!(!transaction.commit().unwrap().patch().is_empty());
    }

    #[test]
    fn borrowed_parse_rejects_a_malformed_live_picture_bullet_owner() {
        let source = source_with_bullets(false);
        let (_, document) =
            crate::embedded::object::Editor::inspect_live_document(&source).unwrap();
        let bullet = png_bullet(4);
        let document_offset = document
            .windows(bullet.len())
            .position(|window| window == bullet.as_slice())
            .unwrap();
        let source_offset = source
            .windows(document.len())
            .position(|window| window == document.as_slice())
            .unwrap();
        let mut malformed = source;
        // The first byte of a BlipEntityAtom payload is its preferred image
        // type. Keep the CFB and record framing intact while invalidating the
        // admitted picture-bullet owner.
        malformed[source_offset + document_offset + 8] = 0x04;
        assert!(PictureBulletSnapshot::parse(&malformed).is_err());
    }
}
