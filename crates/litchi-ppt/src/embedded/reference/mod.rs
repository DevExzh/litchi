//! Strict `PowerPoint` shape references to inert external objects.

use crate::consts::RecordType;
use crate::embedded::object::{Collection as OleCollection, ExternalObject};
use crate::external_media::{Collection as MediaCollection, Object};
use crate::package::{Error, Result};
use crate::records::Record;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Reference {
    pub id: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Target<'a> {
    Media(&'a Object),
    Ole(&'a ExternalObject),
}

impl Reference {
    /// Create a reference to the external object with the given `id`.
    ///
    /// # Errors
    ///
    /// Returns an error if `id` is zero.
    pub fn new(id: u32) -> Result<Self> {
        if id == 0 {
            return corrupted("ExObjRefAtom external-object ID must be positive");
        }
        Ok(Self { id })
    }

    /// Parse an `ExObjRefAtom` record into an external-object reference.
    ///
    /// # Errors
    ///
    /// Returns an error if the record header or payload size is invalid or if
    /// the referenced external-object ID is zero.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.version != 0
            || record.instance != 0
            || record.record_type_raw != RecordType::ExternalObjectRefAtom.as_u16()
            || record.data_length != 4
        {
            return corrupted("ExObjRefAtom has an invalid header or size");
        }
        Self::parse_payload(&record.data)
    }

    /// Parse the four-byte external-object ID carried by an `ExObjRefAtom`.
    pub(crate) fn parse_payload(payload: &[u8]) -> Result<Self> {
        if payload.len() != 4 {
            return corrupted("ExObjRefAtom has an invalid payload size");
        }
        Self::new(u32::from_le_bytes([
            payload[0], payload[1], payload[2], payload[3],
        ]))
    }

    /// Serialize as a fully parsed `ExObjRefAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the referenced ID is zero or if the encoded
    /// record cannot be re-parsed.
    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
    }

    /// Serialize to the raw bytes of a complete `ExObjRefAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the referenced external-object ID is zero.
    pub fn to_record_bytes(&self) -> Result<[u8; 12]> {
        let payload = self.to_payload_bytes()?;
        let mut bytes = [0; 12];
        bytes[2..4].copy_from_slice(&RecordType::ExternalObjectRefAtom.as_u16().to_le_bytes());
        bytes[4..8].copy_from_slice(&4u32.to_le_bytes());
        bytes[8..12].copy_from_slice(&payload);
        Ok(bytes)
    }

    /// Encode the four-byte external-object ID payload.
    pub(crate) fn to_payload_bytes(self) -> Result<[u8; 4]> {
        if self.id == 0 {
            return corrupted("ExObjRefAtom external-object ID must be positive");
        }
        Ok(self.id.to_le_bytes())
    }

    /// Resolve the reference against the media and OLE object collections.
    ///
    /// # Errors
    ///
    /// Returns an error if the referenced ID is missing from both collections
    /// or if it is defined by both a media and an OLE object.
    pub fn resolve<'a>(
        &self,
        media: Option<&'a MediaCollection>,
        ole: Option<&'a OleCollection>,
    ) -> Result<Target<'a>> {
        let media_target = media.and_then(|values| values.get(self.id));
        let ole_target = ole.and_then(|values| values.get(self.id));
        match (media_target, ole_target) {
            (Some(value), None) => Ok(Target::Media(value)),
            (None, Some(value)) => Ok(Target::Ole(value)),
            (None, None) => corrupted(format!(
                "ExObjRefAtom references missing external-object ID {}",
                self.id
            )),
            (Some(_), Some(_)) => corrupted(format!(
                "external-object ID {} is defined by both media and OLE objects",
                self.id
            )),
        }
    }
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use crate::{LinkedAudio, LinkedAudioKind, Media};

    #[test]
    fn external_object_reference_roundtrips_exactly() {
        let expected = Reference::new(77).unwrap();
        let parsed = Reference::parse(&expected.to_record().unwrap()).unwrap();
        assert_eq!(parsed, expected);
        assert_eq!(
            parsed.to_record_bytes().unwrap(),
            expected.to_record_bytes().unwrap()
        );
    }

    #[test]
    fn external_object_reference_rejects_null_and_wrong_headers() {
        assert!(Reference::new(0).is_err());
        let mut bytes = Reference::new(1).unwrap().to_record_bytes().unwrap();
        bytes[0] = 1;
        assert!(Reference::parse(&Record::parse(&bytes, 0).unwrap().0).is_err());
    }

    #[test]
    fn resolver_returns_inert_media_and_rejects_missing_ids() {
        let media = MediaCollection {
            id_seed: 5,
            objects: vec![Object::LinkedAudio(LinkedAudio {
                kind: LinkedAudioKind::Wav,
                media: Media {
                    id: 5,
                    loop_playback: false,
                    rewind_after_playing: false,
                    narration: false,
                    unused: [0, 0],
                },
                path: Some("sound.wav".into()),
            })],
            unknown_records: Vec::new(),
        };
        assert!(matches!(
            Reference::new(5)
                .unwrap()
                .resolve(Some(&media), None)
                .unwrap(),
            Target::Media(_)
        ));
        assert!(
            Reference::new(6)
                .unwrap()
                .resolve(Some(&media), None)
                .is_err()
        );
    }
}
