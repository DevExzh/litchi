//! Ordered media-item CRUD for a Keynote presentation soundtrack.

use litchi_iwa_common::media::Type as MediaType;

use super::soundtrack_wire::{
    encoded_media_reference, read_soundtrack, replace_soundtrack_message, rewrite_soundtrack_media,
    soundtrack_media_identifiers, soundtrack_media_payloads,
};
use super::*;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::media::MediaAssetId;

const DOCUMENT_ARCHIVE: &str = "Index/Document.iwa";

/// One ordered audio item assigned to a Keynote presentation soundtrack.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct KeynoteSoundtrackItemInfo {
    /// Zero-based playback position in the soundtrack.
    pub index: usize,
    /// The embedded package asset referenced at this position.
    pub asset: EmbeddedMediaAsset,
}

struct SoundtrackMutation {
    object_id: u64,
    archive_name: String,
    component_id: u64,
    original_data: Vec<u8>,
    payloads: Vec<Vec<u8>>,
}

impl KeynoteEditor {
    /// List soundtrack audio items in playback order.
    pub fn soundtrack_items(&self) -> Result<Vec<KeynoteSoundtrackItemInfo>> {
        let graph = ObjectGraph::read(self.package())?;
        let Some(record) = read_soundtrack(&graph)? else {
            return Ok(Vec::new());
        };
        let identifiers = soundtrack_media_identifiers(record.data)?
            .into_iter()
            .map(MediaAssetId::try_from)
            .collect::<Result<Vec<_>>>()?;
        let media = IWorkMediaEditor::from_package(self.package().clone())?;
        identifiers
            .into_iter()
            .enumerate()
            .map(|(index, identifier)| {
                let asset = media.asset(identifier).cloned().ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote soundtrack references missing data identifier {identifier}"
                    ))
                })?;
                validate_soundtrack_asset(&asset)?;
                Ok(KeynoteSoundtrackItemInfo { index, asset })
            })
            .collect()
    }

    /// Append a new embedded audio item and return its assigned soundtrack entry.
    pub fn add_soundtrack_item(
        &mut self,
        preferred_filename: &str,
        data: &[u8],
    ) -> Result<KeynoteSoundtrackItemInfo> {
        let index = self.soundtrack_context()?.payloads.len();
        self.insert_soundtrack_item(index, preferred_filename, data)
    }

    /// Insert a new embedded audio item at `index`.
    pub fn insert_soundtrack_item(
        &mut self,
        index: usize,
        preferred_filename: &str,
        data: &[u8],
    ) -> Result<KeynoteSoundtrackItemInfo> {
        validate_new_soundtrack_audio(preferred_filename, data)?;
        let mut context = self.soundtrack_context()?;
        if index > context.payloads.len() {
            return Err(index_error(index, context.payloads.len(), true));
        }

        let mut media = IWorkMediaEditor::from_package(self.package().clone())?;
        let inserted = media.insert_unreferenced(preferred_filename, data)?;
        context.payloads.insert(
            index,
            encoded_media_reference(inserted.data_identifier.get())?,
        );
        let mut staged = media.into_package();
        apply_soundtrack_payloads(&mut staged, &context)?;
        add_component_data_reference(
            &mut staged,
            context.component_id,
            inserted.data_identifier.get(),
            context.object_id,
        )?;
        self.commit_soundtrack_items(staged, &context.payloads)?;
        self.soundtrack_items()?
            .get(index)
            .cloned()
            .ok_or_else(|| Error::InvalidFormat("Inserted soundtrack item vanished".to_owned()))
    }

    /// Replace one soundtrack entry with a newly embedded audio asset.
    ///
    /// A fresh data identifier is used, so other objects or duplicate soundtrack
    /// entries that reference the previous asset are not modified.
    pub fn replace_soundtrack_item(
        &mut self,
        index: usize,
        preferred_filename: &str,
        data: &[u8],
    ) -> Result<KeynoteSoundtrackItemInfo> {
        validate_new_soundtrack_audio(preferred_filename, data)?;
        let mut context = self.soundtrack_context()?;
        let old_identifier = identifier_at(&context.payloads, index)?;

        let mut media = IWorkMediaEditor::from_package(self.package().clone())?;
        let inserted = media.insert_unreferenced(preferred_filename, data)?;
        context.payloads[index] = encoded_media_reference(inserted.data_identifier.get())?;
        let mut staged = media.into_package();
        apply_soundtrack_payloads(&mut staged, &context)?;
        add_component_data_reference(
            &mut staged,
            context.component_id,
            inserted.data_identifier.get(),
            context.object_id,
        )?;
        remove_component_data_reference(
            &mut staged,
            context.component_id,
            old_identifier.get(),
            context.object_id,
        )?;
        remove_asset_if_unreferenced(&mut staged, old_identifier)?;
        self.commit_soundtrack_items(staged, &context.payloads)?;
        self.soundtrack_items()?
            .get(index)
            .cloned()
            .ok_or_else(|| Error::InvalidFormat("Replaced soundtrack item vanished".to_owned()))
    }

    /// Move one soundtrack item to a new final index without rewriting its reference payload.
    pub fn move_soundtrack_item(&mut self, from_index: usize, to_index: usize) -> Result<()> {
        let mut context = self.soundtrack_context()?;
        let count = context.payloads.len();
        if from_index >= count {
            return Err(index_error(from_index, count, false));
        }
        if to_index >= count {
            return Err(index_error(to_index, count, false));
        }
        if from_index == to_index {
            return Ok(());
        }
        let payload = context.payloads.remove(from_index);
        context.payloads.insert(to_index, payload);
        let mut staged = self.package().clone();
        apply_soundtrack_payloads(&mut staged, &context)?;
        self.commit_soundtrack_items(staged, &context.payloads)
    }

    /// Remove one soundtrack entry and cull its media asset when no reference remains.
    pub fn remove_soundtrack_item(&mut self, index: usize) -> Result<KeynoteSoundtrackItemInfo> {
        let items = self.soundtrack_items()?;
        let removed = items
            .get(index)
            .cloned()
            .ok_or_else(|| index_error(index, items.len(), false))?;
        let mut context = self.soundtrack_context()?;
        let identifier = identifier_at(&context.payloads, index)?;
        context.payloads.remove(index);

        let mut staged = self.package().clone();
        apply_soundtrack_payloads(&mut staged, &context)?;
        remove_component_data_reference(
            &mut staged,
            context.component_id,
            identifier.get(),
            context.object_id,
        )?;
        remove_asset_if_unreferenced(&mut staged, identifier)?;
        self.commit_soundtrack_items(staged, &context.payloads)?;
        Ok(removed)
    }

    fn soundtrack_context(&self) -> Result<SoundtrackMutation> {
        let graph = ObjectGraph::read(self.package())?;
        let record = read_soundtrack(&graph)?.ok_or_else(|| {
            Error::InvalidFormat("Keynote show has no soundtrack object".to_owned())
        })?;
        let component_id = component_identifier_for_entry(self.package(), DOCUMENT_ARCHIVE)?
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Keynote document has no PackageMetadata component registration".to_owned(),
                )
            })?;
        Ok(SoundtrackMutation {
            object_id: record.id,
            archive_name: graph.archive_name(record.id)?.to_owned(),
            component_id,
            original_data: record.data.to_vec(),
            payloads: soundtrack_media_payloads(record.data)?,
        })
    }

    fn commit_soundtrack_items(
        &mut self,
        staged: IWorkPackage,
        expected_payloads: &[Vec<u8>],
    ) -> Result<()> {
        let serialized = staged.to_bytes()?;
        let verified = Self::from_bytes(&serialized)?;
        let graph = ObjectGraph::read(verified.package())?;
        let record = read_soundtrack(&graph)?.ok_or_else(|| {
            Error::InvalidFormat("Keynote soundtrack vanished during verification".to_owned())
        })?;
        if soundtrack_media_payloads(record.data)? != expected_payloads {
            return Err(Error::InvalidFormat(
                "Keynote soundtrack item mutation failed round-trip validation".to_owned(),
            ));
        }
        verified.soundtrack_items()?;
        self.text = IWorkTextEditor::from_package(staged);
        Ok(())
    }
}

fn apply_soundtrack_payloads(
    package: &mut IWorkPackage,
    context: &SoundtrackMutation,
) -> Result<()> {
    let data = rewrite_soundtrack_media(&context.original_data, &context.payloads)?;
    package.update_archive(&context.archive_name, |archive| {
        replace_soundtrack_message(archive, context.object_id, data)
    })
}

fn remove_asset_if_unreferenced(
    package: &mut IWorkPackage,
    identifier: MediaAssetId,
) -> Result<()> {
    let mut media = IWorkMediaEditor::from_package(package.clone())?;
    let asset = media.asset(identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Removed soundtrack data identifier {identifier} is missing"
        ))
    })?;
    if !asset.is_referenced() {
        media.remove_unreferenced(identifier)?;
        *package = media.into_package();
    }
    Ok(())
}

fn identifier_at(payloads: &[Vec<u8>], index: usize) -> Result<MediaAssetId> {
    let payload = payloads
        .get(index)
        .ok_or_else(|| index_error(index, payloads.len(), false))?;
    MediaAssetId::try_from(tsp::DataReference::decode(payload.as_slice())?.identifier)
}

fn validate_soundtrack_asset(asset: &EmbeddedMediaAsset) -> Result<()> {
    if asset.media_type != MediaType::Audio || !asset.is_materialized() {
        return Err(Error::InvalidFormat(format!(
            "Keynote soundtrack data identifier {} is not materialized audio",
            asset.data_identifier
        )));
    }
    Ok(())
}

fn validate_new_soundtrack_audio(preferred_filename: &str, data: &[u8]) -> Result<()> {
    let extension_type = Path::new(preferred_filename)
        .extension()
        .and_then(|extension| extension.to_str())
        .map(MediaType::from_extension)
        .unwrap_or(MediaType::Unknown);
    let signature_type = MediaType::from_bytes(data);
    if extension_type != MediaType::Audio || signature_type != MediaType::Audio {
        return Err(Error::ParseError(format!(
            "Keynote soundtrack items require an audio filename and signature; got {} and {}",
            extension_type.name(),
            signature_type.name()
        )));
    }
    Ok(())
}

fn index_error(index: usize, count: usize, insertion: bool) -> Error {
    let valid = if insertion {
        format!("0..={count}")
    } else if count == 0 {
        "no valid indexes".to_owned()
    } else {
        format!("0..{}", count - 1)
    };
    Error::ParseError(format!(
        "Keynote soundtrack item index {index} is out of range ({valid})"
    ))
}
