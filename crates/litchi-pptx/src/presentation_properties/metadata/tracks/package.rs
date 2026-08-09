//! `WebVTT` parts and contextual media-track OPC ownership.

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI, Part};

pub use super::codec::{load, store};
use super::model::{Caption, CaptionTarget, MediaKey, MediaMetadata, TracksInfo};
use super::tracks_info::{self, Found};
use super::transaction::{Commit, Patch, Snapshot};
use super::validation::{parse_boolean, validate_relationship_id};
use crate::{Error, Result};

/// Discover one media picture's typed TracksInfo/narration metadata.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_media(package: &OpcPackage, key: &MediaKey) -> Result<Option<Snapshot>> {
    let source_name = PackURI::new(&key.slide_part_name).map_err(Error::Uri)?;
    let source = package.get_part(&source_name)?;
    if source.content_type() != ct::PML_SLIDE {
        return Err(Error::ContentType {
            expected: ct::PML_SLIDE.to_owned(),
            actual: source.content_type().to_owned(),
        });
    }
    let Some(found) = tracks_info::discover(source.blob(), key)? else {
        return Ok(None);
    };
    let metadata = resolve_metadata(package, source, &found)?;
    Snapshot::from_wire(
        key.slide_part_name.clone(),
        source.blob().to_vec(),
        found,
        metadata,
    )
    .map(Some)
}

/// Publish a source-checked patch atomically to the owning slide part.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn apply_media_patch(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    let source_name = PackURI::new(&patch.source_part_name).map_err(Error::Uri)?;
    let source = package.get_part(&source_name)?;
    if fingerprint(source.blob()) != patch.expected_revision
        || source.blob() != patch.expected_xml.as_slice()
    {
        return Err(invalid("media tracks source is stale"));
    }
    if !patch.is_changed() {
        return load_media(package, &patch.key)?.ok_or_else(|| {
            invalid("media tracks no-op source no longer contains the selected shape")
        });
    }

    let mut staged = package.clone();
    staged
        .get_part_mut(&source_name)?
        .set_blob(patch.updated_xml.as_ref().clone());
    let snapshot = load_media(&staged, &patch.key)?
        .ok_or_else(|| invalid("media tracks patch removed the selected media shape"))?;
    *package = staged;
    Ok(snapshot)
}

/// Publish a commit and return the validated post-publication snapshot.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn apply_media_commit(package: &mut OpcPackage, commit: Commit) -> Result<Snapshot> {
    apply_media_patch(package, &commit.into_patch())
}

fn resolve_metadata(
    package: &OpcPackage,
    source: &dyn Part,
    found: &Found,
) -> Result<MediaMetadata> {
    let media_relationship_id = effective(&found.media.embed, &found.media.link)
        .ok_or_else(|| invalid("p14:media requires r:embed or r:link"))?;
    validate_relationship_id(media_relationship_id, "media relationship ID")?;
    validate_media_relationship(package, source, media_relationship_id)?;

    let tracks_info = found
        .media
        .tracks_info
        .as_ref()
        .map(|tracks| resolve_tracks(package, source, tracks))
        .transpose()?;
    let narration = match found.narration.as_ref() {
        None => None,
        Some(value) => value
            .value
            .as_ref()
            .map(|value| parse_boolean(&value.value, "isNarration/@val"))
            .transpose()?,
    };
    Ok(MediaMetadata {
        key: found.key.clone(),
        media_relationship_id: Some(media_relationship_id.to_owned()),
        tracks_info,
        narration,
    })
}

fn resolve_tracks(
    package: &OpcPackage,
    source: &dyn Part,
    value: &tracks_info::TracksInfo,
) -> Result<TracksInfo> {
    let display_location =
        super::model::DisplayLocation::from_token(&value.display_location.value)?;
    let mut captions = Vec::with_capacity(value.tracks.len());
    for track in &value.tracks {
        let relationship_id = effective(&track.embed, &track.link)
            .ok_or_else(|| invalid("track requires r:embed or r:link"))?;
        validate_relationship_id(relationship_id, "caption track relationship ID")?;
        let target = resolve_caption_relationship(package, source, relationship_id)?;
        captions.push(Caption {
            id: track.id.value.clone(),
            label: track.label.value.clone(),
            language: track.language.as_ref().map(|value| value.value.clone()),
            target,
        });
    }
    Ok(TracksInfo {
        display_location,
        captions,
    })
}

fn validate_media_relationship(package: &OpcPackage, source: &dyn Part, id: &str) -> Result<()> {
    let relationship = source
        .rels()
        .get(id)
        .ok_or_else(|| Error::Relationship(format!("missing media relationship '{id}'")))?;
    if relationship.is_external() {
        if relationship.target_ref().is_empty() {
            return Err(Error::Relationship(format!(
                "external media relationship '{id}' has an empty target"
            )));
        }
    } else {
        let target = relationship.target_partname()?;
        package.get_part(&target)?;
    }
    Ok(())
}

fn resolve_caption_relationship(
    package: &OpcPackage,
    source: &dyn Part,
    id: &str,
) -> Result<CaptionTarget> {
    let relationship = source
        .rels()
        .get(id)
        .ok_or_else(|| Error::Relationship(format!("missing caption relationship '{id}'")))?;
    if relationship.reltype() != super::codec::RELATIONSHIP_TYPE {
        return Err(Error::Relationship(format!(
            "caption relationship '{id}' has unexpected type '{}'",
            relationship.reltype()
        )));
    }
    if relationship.is_external() {
        if relationship.target_ref().is_empty() {
            return Err(Error::Relationship(format!(
                "external caption relationship '{id}' has an empty target"
            )));
        }
        return Ok(CaptionTarget::External {
            target: relationship.target_ref().to_owned(),
        });
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    if part.content_type() != super::codec::CONTENT_TYPE {
        return Err(Error::ContentType {
            expected: super::codec::CONTENT_TYPE.to_owned(),
            actual: part.content_type().to_owned(),
        });
    }
    if !part.rels().is_empty() {
        return Err(invalid(format!(
            "caption track part '{target}' has outbound relationships"
        )));
    }
    Ok(CaptionTarget::Internal {
        part_name: target.to_string(),
        content_type: part.content_type().to_owned(),
    })
}

fn effective<'a>(
    embed: &'a Option<tracks_info::Attr>,
    link: &'a Option<tracks_info::Attr>,
) -> Option<&'a str> {
    link.as_ref()
        .or(embed.as_ref())
        .map(|value| value.value.as_str())
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100_0000_01b3);
    }
    hash
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
