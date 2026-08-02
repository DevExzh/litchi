//! Strict PowerPoint shape references to inert external objects.

use super::external_media::{PowerPointExternalMediaCollection, PowerPointExternalMediaObject};
use super::ole_object::{PowerPointOleExternalObject, PowerPointOleObjectCollection};
use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PowerPointExternalObjectReference {
    pub id: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PowerPointExternalObjectTarget<'a> {
    Media(&'a PowerPointExternalMediaObject),
    Ole(&'a PowerPointOleExternalObject),
}

impl PowerPointExternalObjectReference {
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
            || record.data.len() != 4
            || record.data_length != 4
        {
            return corrupted("ExObjRefAtom has an invalid header or size");
        }
        Self::new(u32::from_le_bytes(
            record.data[..4].try_into().expect("fixed slice"),
        ))
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<[u8; 12]> {
        if self.id == 0 {
            return corrupted("ExObjRefAtom external-object ID must be positive");
        }
        let mut bytes = [0; 12];
        bytes[2..4].copy_from_slice(&PptRecordType::ExternalObjectRefAtom.as_u16().to_le_bytes());
        bytes[4..8].copy_from_slice(&4u32.to_le_bytes());
        bytes[8..12].copy_from_slice(&self.id.to_le_bytes());
        Ok(bytes)
    }

    pub fn resolve<'a>(
        &self,
        media: Option<&'a PowerPointExternalMediaCollection>,
        ole: Option<&'a PowerPointOleObjectCollection>,
    ) -> Result<PowerPointExternalObjectTarget<'a>> {
        let media = media.and_then(|values| values.get(self.id));
        let ole = ole.and_then(|values| values.get(self.id));
        match (media, ole) {
            (Some(value), None) => Ok(PowerPointExternalObjectTarget::Media(value)),
            (None, Some(value)) => Ok(PowerPointExternalObjectTarget::Ole(value)),
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
        let expected = PowerPointExternalObjectReference::new(77).unwrap();
        let parsed =
            PowerPointExternalObjectReference::parse(&expected.to_record().unwrap()).unwrap();
        assert_eq!(parsed, expected);
        assert_eq!(
            parsed.to_record_bytes().unwrap(),
            expected.to_record_bytes().unwrap()
        );
    }

    #[test]
    fn external_object_reference_rejects_null_and_wrong_headers() {
        assert!(PowerPointExternalObjectReference::new(0).is_err());
        let mut bytes = PowerPointExternalObjectReference::new(1)
            .unwrap()
            .to_record_bytes()
            .unwrap();
        bytes[0] = 1;
        assert!(
            PowerPointExternalObjectReference::parse(&PptRecord::parse(&bytes, 0).unwrap().0)
                .is_err()
        );
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
            PowerPointExternalObjectReference::new(5)
                .unwrap()
                .resolve(Some(&media), None)
                .unwrap(),
            PowerPointExternalObjectTarget::Media(_)
        ));
        assert!(
            PowerPointExternalObjectReference::new(6)
                .unwrap()
                .resolve(Some(&media), None)
                .is_err()
        );
    }
}
