//! PowerPoint 9 picture-bullet collection parsing.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;

/// Preferred native picture format for a PowerPoint 9 bullet.
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
    type Error = PptError;

    fn try_from(value: u8) -> Result<Self> {
        match value {
            0x02 => Ok(Self::Emf),
            0x03 => Ok(Self::Wmf),
            0x05 => Ok(Self::Jpeg),
            0x06 => Ok(Self::Png),
            _ => Err(PptError::Corrupted(
                "BlipEntityAtom has an invalid winBlipType".to_string(),
            )),
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
    /// Complete embedded OfficeArt BLIP or FBSE record, including its header.
    pub officeart_record: Vec<u8>,
}

impl PictureBullet {
    /// Decode the embedded OfficeArt image using the image-conversion feature.
    #[cfg(feature = "imgconv")]
    pub fn blip(&self) -> Result<litchi_imgconv::Blip<'static>> {
        self.blip_with_delay_stream(None)
    }

    /// Decode the image, optionally resolving a delay-loaded FBSE from its stream.
    #[cfg(feature = "imgconv")]
    pub fn blip_with_delay_stream(
        &self,
        delay_stream: Option<&[u8]>,
    ) -> Result<litchi_imgconv::Blip<'static>> {
        let (record, consumed) = crate::escher::EscherRecord::parse(&self.officeart_record, 0)
            .map_err(|error| {
                PptError::Corrupted(format!("Invalid picture-bullet BLIP: {error}"))
            })?;
        if consumed != self.officeart_record.len() {
            return Err(PptError::Corrupted(
                "Picture-bullet BLIP was only partially parsed".to_string(),
            ));
        }
        crate::extractor::ImageExtractor::extract_from_escher_record_with_stream(
            &record,
            delay_stream,
        )
        .map(|image| image.blip)
        .map_err(|error| {
            PptError::Corrupted(format!("Could not decode picture-bullet BLIP: {error}"))
        })
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
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.record_type != PptRecordType::BlipCollection9
            || record.version != 0x0f
            || record.instance != 0
        {
            return Err(PptError::Corrupted(
                "BlipCollection9Container has an invalid record header".to_string(),
            ));
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "picture-bullet collection")?;
        let mut bullets = Vec::with_capacity(children.len());
        for child in children {
            if child.record_type != PptRecordType::BlipEntity9Atom
                || child.version != 0
                || child.instance > 0x80
            {
                return Err(PptError::Corrupted(
                    "Picture-bullet collection has an invalid child record".to_string(),
                ));
            }
            if bullets
                .iter()
                .any(|bullet: &PictureBullet| bullet.index == child.instance)
            {
                return Err(PptError::Corrupted(
                    "Picture-bullet collection has a duplicate index".to_string(),
                ));
            }
            bullets.push(parse_picture_bullet(&child)?);
        }
        Ok(Self { bullets })
    }

    /// Discover the single PowerPoint 9 picture-bullet collection below `root`.
    pub fn parse_from(root: &PptRecord) -> Result<Option<Self>> {
        let mut result = None;
        for record in root.versioned_binary_tag_records(9)? {
            if record.record_type != PptRecordType::BlipCollection9 {
                continue;
            }
            if result.replace(Self::parse(&record)?).is_some() {
                return Err(PptError::Corrupted(
                    "Record tree contains multiple picture-bullet collections".to_string(),
                ));
            }
        }
        Ok(result)
    }

    /// Resolve a `bulletBlipRef`; `-1` is the null reference.
    pub fn get(&self, reference: i16) -> Option<&PictureBullet> {
        let index = u16::try_from(reference).ok()?;
        self.bullets.iter().find(|bullet| bullet.index == index)
    }
}

fn parse_picture_bullet(record: &PptRecord) -> Result<PictureBullet> {
    if record.data.len() < 10 {
        return Err(PptError::Corrupted(
            "BlipEntityAtom is truncated".to_string(),
        ));
    }
    let picture_type = PictureBulletType::try_from(record.data[0])?;
    let unused = record.data[1];
    let officeart_record = &record.data[2..];
    let instance_version = u16::from_le_bytes([officeart_record[0], officeart_record[1]]);
    let version = instance_version & 0x000f;
    let instance = instance_version >> 4;
    let record_type = u16::from_le_bytes([officeart_record[2], officeart_record[3]]);
    let length = u32::from_le_bytes([
        officeart_record[4],
        officeart_record[5],
        officeart_record[6],
        officeart_record[7],
    ]);
    let length = usize::try_from(length)
        .map_err(|_| PptError::Corrupted("Picture-bullet BLIP size overflow".to_string()))?;
    if 8usize.checked_add(length) != Some(officeart_record.len()) {
        return Err(PptError::Corrupted(
            "Picture-bullet OfficeArt record has an invalid size".to_string(),
        ));
    }
    if record_type == 0xf007 {
        validate_fbse(officeart_record, version, instance, picture_type)?;
    } else {
        validate_direct_blip(
            officeart_record,
            version,
            instance,
            record_type,
            picture_type,
        )?;
    }

    Ok(PictureBullet {
        index: record.instance,
        picture_type,
        unused,
        officeart_record: officeart_record.to_vec(),
    })
}

fn validate_fbse(
    record: &[u8],
    version: u16,
    instance: u16,
    picture_type: PictureBulletType,
) -> Result<()> {
    let payload = &record[8..];
    if version != 2 || payload.len() < 36 {
        return Err(PptError::Corrupted(
            "Picture-bullet FBSE has an invalid version or size".to_string(),
        ));
    }
    let win_type = payload[0];
    let mac_type = payload[1];
    if instance != u16::from(win_type) && instance != u16::from(mac_type) {
        return Err(PptError::Corrupted(
            "Picture-bullet FBSE instance does not match its BLIP types".to_string(),
        ));
    }
    if !picture_type_matches(picture_type, win_type) {
        return Err(PptError::Corrupted(
            "Picture-bullet preferred and stored BLIP types disagree".to_string(),
        ));
    }

    let name_length = usize::from(payload[33]);
    if name_length & 1 != 0 {
        return Err(PptError::Corrupted(
            "Picture-bullet FBSE name has an odd byte length".to_string(),
        ));
    }
    let embedded_offset = 36usize
        .checked_add(name_length)
        .ok_or_else(|| PptError::Corrupted("Picture-bullet FBSE name size overflow".to_string()))?;
    if embedded_offset > payload.len() {
        return Err(PptError::Corrupted(
            "Picture-bullet FBSE name extends beyond the record".to_string(),
        ));
    }
    if name_length > 0 {
        let name = &payload[36..embedded_offset];
        let valid_utf16 = name.ends_with(&[0, 0])
            && char::decode_utf16(
                name[..name.len() - 2]
                    .chunks_exact(2)
                    .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]])),
            )
            .all(|character| character.is_ok());
        if !valid_utf16 {
            return Err(PptError::Corrupted(
                "Picture-bullet FBSE name is not valid null-terminated UTF-16".to_string(),
            ));
        }
    }
    if embedded_offset == payload.len() {
        let delay_offset = u32::from_le_bytes([payload[28], payload[29], payload[30], payload[31]]);
        if delay_offset == u32::MAX {
            return Err(PptError::Corrupted(
                "Picture-bullet FBSE does not identify any BLIP data".to_string(),
            ));
        }
        return Ok(());
    }

    let embedded_size = usize::try_from(u32::from_le_bytes([
        payload[20],
        payload[21],
        payload[22],
        payload[23],
    ]))
    .map_err(|_| PptError::Corrupted("Picture-bullet FBSE size overflow".to_string()))?;
    if embedded_offset.checked_add(embedded_size) != Some(payload.len()) {
        return Err(PptError::Corrupted(
            "Picture-bullet FBSE embedded BLIP has an invalid size".to_string(),
        ));
    }
    let embedded = &payload[embedded_offset..];
    if embedded.len() < 8 {
        return Err(PptError::Corrupted(
            "Picture-bullet FBSE embedded BLIP is truncated".to_string(),
        ));
    }
    let options = u16::from_le_bytes([embedded[0], embedded[1]]);
    let embedded_type = u16::from_le_bytes([embedded[2], embedded[3]]);
    validate_direct_blip(
        embedded,
        options & 0x000f,
        options >> 4,
        embedded_type,
        picture_type,
    )
}

fn validate_direct_blip(
    record: &[u8],
    version: u16,
    instance: u16,
    record_type: u16,
    picture_type: PictureBulletType,
) -> Result<()> {
    if version != 0 || record.len() < 8 {
        return Err(PptError::Corrupted(
            "Picture-bullet OfficeArt BLIP has an invalid version or size".to_string(),
        ));
    }
    let length = usize::try_from(u32::from_le_bytes([
        record[4], record[5], record[6], record[7],
    ]))
    .map_err(|_| PptError::Corrupted("Picture-bullet BLIP size overflow".to_string()))?;
    if 8usize.checked_add(length) != Some(record.len()) {
        return Err(PptError::Corrupted(
            "Picture-bullet OfficeArt BLIP has an invalid size".to_string(),
        ));
    }

    let (stored_type, one_uid_instance, two_uid_instance, one_uid_size, two_uid_size) =
        match record_type {
            0xf01a => (0x02, 0x3d4, 0x3d5, 50, 66),
            0xf01b => (0x03, 0x216, 0x217, 50, 66),
            0xf01d | 0xf02a if matches!(instance, 0x46a | 0x46b) => (0x05, 0x46a, 0x46b, 17, 33),
            0xf01d | 0xf02a => (0x05, 0x6e2, 0x6e3, 17, 33),
            0xf01e => (0x06, 0x6e0, 0x6e1, 17, 33),
            _ => {
                return Err(PptError::Corrupted(
                    "Picture bullet contains an unsupported OfficeArt BLIP type".to_string(),
                ));
            },
        };
    let minimum_size = if instance == one_uid_instance {
        one_uid_size
    } else if instance == two_uid_instance {
        two_uid_size
    } else {
        return Err(PptError::Corrupted(
            "Picture-bullet OfficeArt BLIP has an invalid record instance".to_string(),
        ));
    };
    if length < minimum_size || stored_type != picture_type as u8 {
        return Err(PptError::Corrupted(
            "Picture-bullet preferred and stored BLIP types disagree".to_string(),
        ));
    }
    Ok(())
}

fn picture_type_matches(picture_type: PictureBulletType, stored_type: u8) -> bool {
    stored_type == picture_type as u8
        || (picture_type == PictureBulletType::Jpeg && stored_type == 0x12)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
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
        fbse.extend_from_slice(&(blip.len() as u32).to_le_bytes());
        fbse.extend_from_slice(&1u32.to_le_bytes());
        fbse.extend_from_slice(&u32::MAX.to_le_bytes());
        fbse.extend_from_slice(&[0, 0, 0, 0]);
        fbse.extend_from_slice(blip);
        let mut payload = vec![0x06, 0];
        payload.extend_from_slice(&record_bytes(2, 0x06, 0xf007, &fbse));
        record_bytes(0, index, 2041, &payload)
    }

    fn collection(payload: Vec<u8>) -> PptRecord {
        PptRecord {
            record_type: PptRecordType::BlipCollection9,
            record_type_raw: 2040,
            version: 0x0f,
            instance: 0,
            data_length: payload.len() as u32,
            data: payload,
            children: Vec::new(),
        }
    }

    fn prog_tags_record(blob_payload: &[u8]) -> PptRecord {
        let tag_name: Vec<u8> = "___PPT9"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let name = record_bytes(0, 0, 4026, &tag_name);
        let blob = record_bytes(0, 0, 0x138b, blob_payload);
        let mut tag_payload = name;
        tag_payload.extend_from_slice(&blob);
        let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
        PptRecord {
            record_type: PptRecordType::ProgTags,
            record_type_raw: 0x1388,
            version: 0x0f,
            instance: 0,
            data_length: tag.len() as u32,
            data: tag,
            children: Vec::new(),
        }
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
        let root = PptRecord {
            record_type: PptRecordType::Document,
            record_type_raw: 1000,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children: vec![prog_tags_record(&collection)],
        };
        let bullets = PictureBulletCollection::parse_from(&root).unwrap().unwrap();
        assert_eq!(bullets.get(7).unwrap().picture_type, PictureBulletType::Png);
    }

    #[cfg(feature = "imgconv")]
    #[test]
    fn decodes_direct_and_fbse_picture_bullets() {
        let direct = PictureBulletCollection::parse(&collection(png_bullet(1))).unwrap();
        assert_eq!(direct.get(1).unwrap().blip().unwrap().picture_data(), []);

        let embedded = PictureBulletCollection::parse(&collection(fbse_png_bullet(2))).unwrap();
        assert_eq!(embedded.get(2).unwrap().blip().unwrap().picture_data(), []);
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
}
