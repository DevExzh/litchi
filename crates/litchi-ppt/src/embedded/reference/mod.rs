//! Strict PowerPoint shape references to inert external objects.

use crate::consts::PptRecordType;
use crate::embedded::object::{Collection, ExternalObject};
use crate::external_media::{PowerPointExternalMediaCollection, PowerPointExternalMediaObject};
use crate::package::{PptError, Result};
use crate::records::PptRecord;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Reference {
    pub id: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Target<'a> {
    Media(&'a PowerPointExternalMediaObject),
    Ole(&'a ExternalObject),
}

impl Reference {
    pub fn new(id: u32) -> Result<Self> {
        if id == 0 {
            return corrupted("ExObjRefAtom external-object ID must be positive");
        }
        Ok(Self { id })
    }

    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.version != 0
            || record.instance != 0
            || record.record_type_raw != PptRecordType::ExternalObjectRefAtom.as_u16()
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
        Self::new(u32::from_le_bytes(
            payload.try_into().expect("validated payload size"),
        ))
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<[u8; 12]> {
        let payload = self.to_payload_bytes()?;
        let mut bytes = [0; 12];
        bytes[2..4].copy_from_slice(&PptRecordType::ExternalObjectRefAtom.as_u16().to_le_bytes());
        bytes[4..8].copy_from_slice(&4u32.to_le_bytes());
        bytes[8..12].copy_from_slice(&payload);
        Ok(bytes)
    }

    /// Encode the four-byte external-object ID payload.
    pub(crate) fn to_payload_bytes(&self) -> Result<[u8; 4]> {
        if self.id == 0 {
            return corrupted("ExObjRefAtom external-object ID must be positive");
        }
        Ok(self.id.to_le_bytes())
    }

    pub fn resolve<'a>(
        &self,
        media: Option<&'a PowerPointExternalMediaCollection>,
        ole: Option<&'a Collection>,
    ) -> Result<Target<'a>> {
        let media = media.and_then(|values| values.get(self.id));
        let ole = ole.and_then(|values| values.get(self.id));
        match (media, ole) {
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
    Err(PptError::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{PowerPointExternalMedia, PowerPointLinkedAudio, PowerPointLinkedAudioKind};

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
        assert!(Reference::parse(&PptRecord::parse(&bytes, 0).unwrap().0).is_err());
    }

    #[test]
    fn resolver_returns_inert_media_and_rejects_missing_ids() {
        let media = PowerPointExternalMediaCollection {
            id_seed: 5,
            objects: vec![PowerPointExternalMediaObject::LinkedAudio(
                PowerPointLinkedAudio {
                    kind: PowerPointLinkedAudioKind::Wav,
                    media: PowerPointExternalMedia {
                        id: 5,
                        loop_playback: false,
                        rewind_after_playing: false,
                        narration: false,
                        unused: [0, 0],
                    },
                    path: Some("sound.wav".into()),
                },
            )],
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
