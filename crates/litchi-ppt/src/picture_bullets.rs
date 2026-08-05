//! PowerPoint 9 picture-bullet collection parsing.

use super::package::{Error, Result};
use super::records::Record;
use crate::consts::RecordType;

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
    /// Complete embedded OfficeArt BLIP or FBSE record, including its header.
    pub officeart_record: Vec<u8>,
}

impl PictureBullet {
    /// Parses the embedded OfficeArt image without copying its file data.
    pub fn blip(&self) -> Result<litchi_odraw::image::Blip<'_>> {
        self.blip_with_delay(None)
    }

    /// Parses the image, optionally resolving a delay-loaded FBSE.
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
}

/// Parsed `BlipCollection9Container`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PictureBulletCollection {
    /// Picture bullets in record order.
    pub bullets: Vec<PictureBullet>,
}

impl PictureBulletCollection {
    /// Parse a `BlipCollection9Container` record.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.record_type != RecordType::BlipCollection9
            || record.version != 0x0f
            || record.instance != 0
        {
            return Err(Error::Corrupted(
                "BlipCollection9Container has an invalid record header".to_string(),
            ));
        }
        let children = Record::parse_sequence_strict(&record.data, "picture-bullet collection")?;
        let mut bullets = Vec::with_capacity(children.len());
        for child in children {
            if child.record_type != RecordType::BlipEntity9Atom
                || child.version != 0
                || child.instance > 0x80
            {
                return Err(Error::Corrupted(
                    "Picture-bullet collection has an invalid child record".to_string(),
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
            bullets.push(parse_picture_bullet(&child)?);
        }
        Ok(Self { bullets })
    }

    /// Discover the single PowerPoint 9 picture-bullet collection below `root`.
    pub fn parse_from(root: &Record) -> Result<Option<Self>> {
        let mut result = None;
        for record in root.versioned_binary_tag_records(9)? {
            if record.record_type != RecordType::BlipCollection9 {
                continue;
            }
            if result.replace(Self::parse(&record)?).is_some() {
                return Err(Error::Corrupted(
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

fn parse_picture_bullet(record: &Record) -> Result<PictureBullet> {
    if record.data.len() < 10 {
        return Err(Error::Corrupted("BlipEntityAtom is truncated".to_string()));
    }
    let picture_type = PictureBulletType::try_from(record.data[0])?;
    let unused = record.data[1];
    let officeart_record = &record.data[2..];
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

    Ok(PictureBullet {
        index: record.instance,
        picture_type,
        unused,
        officeart_record: officeart_record.to_vec(),
    })
}

fn picture_type_matches(
    picture_type: PictureBulletType,
    stored_type: litchi_odraw::image::Kind,
) -> bool {
    stored_type == picture_type.kind()
        || (picture_type == PictureBulletType::Jpeg
            && stored_type == litchi_odraw::image::Kind::CmykJpeg)
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

    fn collection(payload: Vec<u8>) -> Record {
        Record {
            record_type: RecordType::BlipCollection9,
            record_type_raw: 2040,
            version: 0x0f,
            instance: 0,
            data_length: payload.len() as u32,
            data: payload,
            children: Vec::new(),
        }
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
        let root = Record {
            record_type: RecordType::Document,
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

    #[test]
    fn decodes_direct_and_fbse_picture_bullets() {
        let direct = PictureBulletCollection::parse(&collection(png_bullet(1))).unwrap();
        assert_eq!(direct.get(1).unwrap().blip().unwrap().data(), []);

        let embedded = PictureBulletCollection::parse(&collection(fbse_png_bullet(2))).unwrap();
        assert_eq!(embedded.get(2).unwrap().blip().unwrap().data(), []);
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
