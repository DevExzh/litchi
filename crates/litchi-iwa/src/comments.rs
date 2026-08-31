//! Cross-application comments attached directly to drawable objects.

use std::collections::{HashMap, HashSet};
use std::path::Path;
use std::str;
use std::time::{SystemTime, UNIX_EPOCH};

use litchi_iwa_common::comment::{
    AuthorId, Comment, DrawableComment, DrawableId, DrawableInfo, DrawableReply, StorageId, Uuid,
};
use litchi_iwa_protos::comment_storage_codec;
use litchi_iwa_protos::package_metadata_codec::{
    DataReferenceOwnerDescriptor, ExternalReferenceDescriptor, ObjectUuidDescriptor,
    PackageMetadataVisitor, RewriteError,
};
use prost::Message;

use crate::application::Application;
use crate::application_detection::detect;
use crate::archive::{
    ArchiveObject, ArchiveReferenceOccurrence, ArchiveReferencePolicy, ArchiveReferenceVisitor,
    CoreResult, FieldInfo, FieldPath, RawMessage, UnknownFieldRule,
};
use crate::package_metadata::{
    PACKAGE_METADATA_ENTRY, PACKAGE_METADATA_MESSAGE_TYPE, add_component_external_reference,
    advance_package_save_token_for_components, component_identifier_for_entry,
    inspect_package_metadata_source, next_object_identifier, package_metadata_read_options,
    release_package_identifier_suffix, remove_component_external_references_to_object,
    set_package_last_object_identifier,
};
#[cfg(test)]
use crate::protobuf::{kn, tn, tp, tsch, tst, tswp};
use crate::protobuf::{tsd, tsk, tsp};
use crate::wire::{
    append_repeated_length_delimited_field, parse_wire_fields, patch_length_delimited_field,
    patch_nested_length_delimited_field, patch_nested_varint_field, patch_varint_field,
    remove_repeated_length_delimited_field_where, transform_length_delimited_fields_at_path,
};
use crate::{Error, IWorkPackage, Result};

const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3056;
const ANNOTATION_AUTHOR_MESSAGE_TYPE: u32 = 212;
const ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE: u32 = 213;
const APPLE_EPOCH_UNIX_OFFSET_SECONDS: f64 = 978_307_200.0;
const GENERATED_AUTHOR_NAME: &str = "litchi-iwa";
const GENERATED_AUTHOR_PUBLIC_ID: &str = "4C495443-4849-4957-8100-000000000001:058e44481db1c6fdeeac88af010136d7f8949f54bde61ef9af3e078562b968b6";
const COMMENT_STORAGE_CODEC_RECURSION_LIMIT: u32 = 64;
const COMMENT_REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const ANNOTATION_AUTHOR_NAME_FIELD: u32 = 1;
const ANNOTATION_AUTHOR_COLOR_FIELD: u32 = 2;
const ANNOTATION_AUTHOR_PUBLIC_ID_FIELD: u32 = 3;
const ANNOTATION_AUTHOR_IS_PUBLIC_FIELD: u32 = 4;
const ANNOTATION_AUTHOR_PUBLIC_IDS_FIELD: u32 = 5;
const COLOR_MODEL_FIELD: u32 = 1;
const COLOR_RED_FIELD: u32 = 3;
const COLOR_GREEN_FIELD: u32 = 4;
const COLOR_BLUE_FIELD: u32 = 5;
const COLOR_ALPHA_FIELD: u32 = 6;
const COLOR_CYAN_FIELD: u32 = 7;
const COLOR_MAGENTA_FIELD: u32 = 8;
const COLOR_YELLOW_FIELD: u32 = 9;
const COLOR_BLACK_FIELD: u32 = 10;
const COLOR_WHITE_FIELD: u32 = 11;
const COLOR_RGBSPACE_FIELD: u32 = 12;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

fn comment_storage_allocation_error(resource: &'static str, amount: usize) -> Error {
    Error::IwaCommon(litchi_iwa_common::Error::Allocation { resource, amount })
}

fn comment_storage_decode_options(source: &[u8]) -> comment_storage_codec::DecodeOptions {
    comment_storage_codec::DecodeOptions::new(
        source.len().max(1),
        source.len().max(1),
        source.len().saturating_mul(32).max(1),
        COMMENT_STORAGE_CODEC_RECURSION_LIMIT,
        source.len().max(1),
        source.len().max(1),
    )
}

#[derive(Debug, Default)]
struct CommentStorageReplyIds {
    identifiers: Vec<u64>,
    allocation_failed: Option<usize>,
}

impl CommentStorageReplyIds {
    fn into_identifiers(self) -> Result<Vec<u64>> {
        match self.allocation_failed {
            Some(amount) => Err(comment_storage_allocation_error(
                "iWork comment reply identifiers",
                amount,
            )),
            None => Ok(self.identifiers),
        }
    }
}

impl comment_storage_codec::CommentStorageVisitor for CommentStorageReplyIds {
    fn visit_reply(
        &mut self,
        reply: comment_storage_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), comment_storage_codec::DecodeError> {
        if self.allocation_failed.is_some() {
            return Ok(());
        }
        if self.identifiers.try_reserve(1).is_err() {
            // Continue strict traversal so a later malformed record still
            // wins over this candidate-local allocation failure.
            self.allocation_failed = Some(self.identifiers.len().saturating_add(1));
            return Ok(());
        }
        self.identifiers.push(reply.identifier());
        Ok(())
    }
}

fn strict_comment_storage_error(
    storage_id: u64,
    error: comment_storage_codec::DecodeError,
) -> Error {
    Error::InvalidFormat(format!(
        "comment storage object {storage_id} failed strict validation: {error}"
    ))
}

fn decode_comment_storage_payload<'source>(
    storage_id: u64,
    source: &'source [u8],
) -> Result<(
    comment_storage_codec::CommentStorageSnapshot<'source>,
    Vec<u64>,
)> {
    let mut replies = CommentStorageReplyIds::default();
    let (comment, report) = comment_storage_codec::decode_comment_storage_archive_with_visitor(
        source,
        comment_storage_decode_options(source),
        &mut replies,
    )
    .map_err(|error| strict_comment_storage_error(storage_id, error))?;
    let reply_ids = replies.into_identifiers()?;
    if reply_ids.len() != report.reply_references() {
        return Err(Error::InvalidFormat(format!(
            "comment storage object {storage_id} streamed {} replies but reported {}",
            reply_ids.len(),
            report.reply_references(),
        )));
    }
    Ok((comment, reply_ids))
}

fn validate_comment_storage_payload(storage_id: u64, source: &[u8]) -> Result<()> {
    comment_storage_codec::decode_comment_storage_archive(
        source,
        comment_storage_decode_options(source),
    )
    .map(|_| ())
    .map_err(|error| strict_comment_storage_error(storage_id, error))
}

fn drawable_id_from_raw(raw: u64) -> Result<DrawableId> {
    DrawableId::from_raw(raw).map_err(|error| Error::ParseError(error.to_string()))
}

fn storage_id_from_raw(raw: u64) -> Result<StorageId> {
    StorageId::from_raw(raw).map_err(|error| Error::ParseError(error.to_string()))
}

fn author_id_from_raw(raw: u64) -> Result<AuthorId> {
    AuthorId::from_raw(raw).map_err(|error| Error::ParseError(error.to_string()))
}

fn comment_uuid(lower: u64, upper: u64) -> Result<Uuid> {
    Uuid::from_parts(lower, upper).map_err(|error| Error::ParseError(error.to_string()))
}

impl From<litchi_iwa_common::comment::Error> for Error {
    fn from(error: litchi_iwa_common::comment::Error) -> Self {
        Self::ParseError(error.to_string())
    }
}

/// Transactional direct-comment editor shared by Pages, Numbers, and Keynote.
///
/// It edits the `TSD.DrawableArchive.comment` reference nested in each known
/// drawable payload and stores comment bodies in `TSD.CommentStorageArchive`
/// objects. Table-cell comments use a separate table-list indirection and are
/// available through each application's semantic editor.
#[deprecated(
    since = "0.0.1",
    note = "legacy migration-host drawable-comment editor; use focused format-semantic comment owners where available; direct-drawable comment migration remains pending"
)]
#[derive(Debug, Clone)]
pub struct IWorkDrawableCommentEditor {
    package: IWorkPackage,
    application: Application,
}

#[allow(deprecated)]
impl IWorkDrawableCommentEditor {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_package(IWorkPackage::open(path)?)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_package(IWorkPackage::from_bytes(bytes)?)
    }

    pub fn from_package(package: IWorkPackage) -> Result<Self> {
        let application = package_application(&package)?;
        // Decode the full drawable surface now so malformed known payloads fail
        // before a caller receives an editor.
        drawable_locations(&package, application)?;
        Ok(Self {
            package,
            application,
        })
    }

    pub fn application(&self) -> Application {
        self.application
    }

    pub fn drawables(&self) -> Result<Vec<DrawableInfo>> {
        let mut drawables = drawable_locations(&self.package, self.application)?
            .into_values()
            .map(|location| {
                Ok(DrawableInfo {
                    id: drawable_id_from_raw(location.object_id)?,
                    message_type: location.message_type,
                    comment_id: location
                        .comment_storage_object_id
                        .map(storage_id_from_raw)
                        .transpose()?,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        drawables.sort_by_key(|drawable| drawable.id.get());
        Ok(drawables)
    }

    pub fn comment(&self, drawable_object_id: DrawableId) -> Result<Option<DrawableComment>> {
        drawable_comment_in_package(&self.package, self.application, drawable_object_id.get())
    }

    /// Resolves the direct replies to a drawable comment in stored order.
    pub fn replies(&self, drawable_object_id: DrawableId) -> Result<Vec<DrawableReply>> {
        drawable_comment_replies_in_package(
            &self.package,
            self.application,
            drawable_object_id.get(),
        )
    }

    pub fn set_comment(
        &mut self,
        drawable_object_id: DrawableId,
        text: impl Into<String>,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        set_drawable_comment_in_package(
            &mut staged,
            self.application,
            drawable_object_id.get(),
            text.into(),
        )?;
        validate_package_round_trip(&staged)?;
        self.package = staged;
        Ok(())
    }

    pub fn clear_comment(&mut self, drawable_object_id: DrawableId) -> Result<()> {
        let mut staged = self.package.clone();
        clear_drawable_comment_in_package(&mut staged, self.application, drawable_object_id.get())?;
        validate_package_round_trip(&staged)?;
        self.package = staged;
        Ok(())
    }

    /// Adds a reply and returns its new comment-storage object identifier.
    ///
    /// The root storage is copy-on-written, matching native iWork saves and
    /// isolating a drawable when multiple drawables share one thread.
    pub fn add_reply(
        &mut self,
        drawable_object_id: DrawableId,
        text: impl Into<String>,
    ) -> Result<StorageId> {
        let mut staged = self.package.clone();
        let reply_id = add_drawable_comment_reply_in_package(
            &mut staged,
            self.application,
            drawable_object_id.get(),
            text.into(),
        )?;
        validate_package_round_trip(&staged)?;
        let storage_id = storage_id_from_raw(reply_id)?;
        self.package = staged;
        Ok(storage_id)
    }

    /// Updates one direct reply and returns its current storage identifier.
    ///
    /// A changed reply and its root are copy-on-written. The returned value can
    /// therefore differ from `reply_storage_object_id`.
    pub fn set_reply(
        &mut self,
        drawable_object_id: DrawableId,
        reply_storage_object_id: StorageId,
        text: impl Into<String>,
    ) -> Result<StorageId> {
        let mut staged = self.package.clone();
        let reply_id = set_drawable_comment_reply_in_package(
            &mut staged,
            self.application,
            drawable_object_id.get(),
            reply_storage_object_id.get(),
            text.into(),
        )?;
        validate_package_round_trip(&staged)?;
        let reply_id = storage_id_from_raw(reply_id)?;
        self.package = staged;
        Ok(reply_id)
    }

    /// Removes one direct reply from a drawable's comment thread.
    pub fn remove_reply(
        &mut self,
        drawable_object_id: DrawableId,
        reply_storage_object_id: StorageId,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        remove_drawable_comment_reply_in_package(
            &mut staged,
            self.application,
            drawable_object_id.get(),
            reply_storage_object_id.get(),
        )?;
        validate_package_round_trip(&staged)?;
        self.package = staged;
        Ok(())
    }

    pub fn package(&self) -> &IWorkPackage {
        &self.package
    }

    pub fn into_package(self) -> IWorkPackage {
        self.package
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.package.to_bytes()
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.package.save(path)
    }
}

#[derive(Debug, Clone)]
struct DrawableLocation {
    object_id: u64,
    archive_name: String,
    message_index: usize,
    message_type: u32,
    comment_storage_object_id: Option<u64>,
}

fn validate_package_round_trip(package: &IWorkPackage) -> Result<()> {
    let bytes = package.to_bytes()?;
    IWorkPackage::from_bytes(&bytes)?;
    Ok(())
}

fn package_application(package: &IWorkPackage) -> Result<Application> {
    let mut detected = None;
    for name in package.iwa_entry_names() {
        package.with_parsed_archive(name, |archive| {
            let Some(document) = archive.object(1) else {
                return Ok(());
            };
            for message in &document.messages {
                let Some(application) = detect(&message.data) else {
                    continue;
                };
                if detected
                    .replace(application)
                    .is_some_and(|old| old != application)
                {
                    return Err(Error::InvalidFormat(
                        "package contains conflicting iWork document roots".to_owned(),
                    ));
                }
            }
            Ok(())
        })?;
    }
    detected.ok_or_else(|| {
        Error::InvalidFormat("package has no recognizable iWork document root".to_owned())
    })
}

fn object_locations(package: &IWorkPackage) -> Result<HashMap<u64, String>> {
    let mut locations = HashMap::new();
    for name in package.iwa_entry_names() {
        package.with_parsed_archive(name, |archive| {
            for object in &archive.objects {
                let identifier = object.archive_info.identifier.ok_or_else(|| {
                    Error::Archive(format!("object in {name} has no archive identifier"))
                })?;
                if let Some(previous) = locations.insert(identifier, name.to_owned()) {
                    return Err(Error::Archive(format!(
                        "object {identifier} appears in both {previous} and {name}"
                    )));
                }
            }
            Ok(())
        })?;
    }
    Ok(locations)
}

fn drawable_locations(
    package: &IWorkPackage,
    application: Application,
) -> Result<HashMap<u64, DrawableLocation>> {
    let mut result = HashMap::new();
    for name in package.iwa_entry_names() {
        package.with_parsed_archive(name, |archive| {
            for object in &archive.objects {
                let object_id = object.archive_info.identifier.ok_or_else(|| {
                    Error::Archive(format!("object in {name} has no archive identifier"))
                })?;
                let mut location = None;
                for (message_index, message) in object.messages.iter().enumerate() {
                    let Some(payload) = DrawableCommentProjection::decode(
                        application,
                        message.type_,
                        message.data.as_slice(),
                    )?
                    else {
                        continue;
                    };
                    if location.is_some() {
                        return Err(Error::InvalidFormat(format!(
                            "object {object_id} contains multiple direct drawable payloads"
                        )));
                    }
                    location = Some(DrawableLocation {
                        object_id,
                        archive_name: name.to_owned(),
                        message_index,
                        message_type: message.type_,
                        comment_storage_object_id: payload.comment_identifier,
                    });
                }
                if let Some(location) = location {
                    result.insert(object_id, location);
                }
            }
            Ok(())
        })?;
    }
    Ok(result)
}

fn drawable_comment_in_package(
    package: &IWorkPackage,
    application: Application,
    drawable_object_id: u64,
) -> Result<Option<DrawableComment>> {
    let location = drawable_locations(package, application)?
        .remove(&drawable_object_id)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "drawable object {drawable_object_id} was not found"
            ))
        })?;
    let Some(storage_object_id) = location.comment_storage_object_id else {
        return Ok(None);
    };
    let locations = object_locations(package)?;
    let comment = read_comment_storage(package, &locations, storage_object_id)?;
    let drawable_id = drawable_id_from_raw(drawable_object_id)?;
    let storage_id = storage_id_from_raw(storage_object_id)?;
    Ok(Some(DrawableComment {
        drawable_id,
        storage_id,
        comment,
    }))
}

fn drawable_comment_replies_in_package(
    package: &IWorkPackage,
    application: Application,
    drawable_object_id: u64,
) -> Result<Vec<DrawableReply>> {
    let Some(root) = drawable_comment_in_package(package, application, drawable_object_id)? else {
        return Ok(Vec::new());
    };
    let drawable_id = drawable_id_from_raw(drawable_object_id)?;
    let root_storage_id = root.storage_id.get();
    let locations = object_locations(package)?;
    let mut seen = HashSet::new();
    seen.try_reserve(root.comment.reply_ids.len())
        .map_err(|_| {
            comment_storage_allocation_error(
                "iWork drawable comment reply identifiers",
                root.comment.reply_ids.len(),
            )
        })?;
    let mut replies = Vec::new();
    replies
        .try_reserve_exact(root.comment.reply_ids.len())
        .map_err(|_| {
            comment_storage_allocation_error(
                "iWork drawable comment replies",
                root.comment.reply_ids.len(),
            )
        })?;
    for reply_id in root.comment.reply_ids {
        let reply_id = reply_id.get();
        if reply_id == root_storage_id || !seen.insert(reply_id) {
            return Err(Error::InvalidFormat(format!(
                "comment storage {} contains a duplicate or cyclic reply reference to {reply_id}",
                root_storage_id
            )));
        }
        let storage_id = storage_id_from_raw(reply_id)?;
        replies.push(DrawableReply {
            drawable_id,
            root_storage_id: root.storage_id,
            storage_id,
            comment: read_comment_storage(package, &locations, reply_id)?,
        });
    }
    Ok(replies)
}

fn set_drawable_comment_in_package(
    package: &mut IWorkPackage,
    application: Application,
    drawable_object_id: u64,
    text: String,
) -> Result<()> {
    let drawables = drawable_locations(package, application)?;
    let location = drawables.get(&drawable_object_id).cloned().ok_or_else(|| {
        Error::ParseError(format!(
            "drawable object {drawable_object_id} was not found"
        ))
    })?;
    let locations = object_locations(package)?;

    if let Some(storage_id) = location.comment_storage_object_id {
        let old = read_comment_storage(package, &locations, storage_id)?;
        if old.text == text {
            return Ok(());
        }
        validate_direct_reply_graph(package, &locations, storage_id, &old)?;
        let storage_entry = locations.get(&storage_id).cloned().ok_or_else(|| {
            Error::InvalidFormat(format!("comment storage object {storage_id} is missing"))
        })?;
        let direct_users = drawables
            .values()
            .filter(|candidate| candidate.comment_storage_object_id == Some(storage_id))
            .count();
        let globally_owned = if direct_users == 1 {
            prove_global_comment_ownership(package, application, &location, storage_id)?
        } else {
            false
        };
        if globally_owned {
            update_comment_storage_text(package, &locations, storage_id, text)?;
            return advance_save_tokens_for_entries(package, &[storage_entry]);
        }

        let new_storage_id = next_object_identifier(package)?;
        let storage_uuid = fresh_comment_storage_uuid(package)?;
        clone_comment_storage(
            package,
            &locations,
            storage_id,
            new_storage_id,
            storage_uuid,
            text,
        )?;
        replace_drawable_comment_reference(
            package,
            application,
            &location,
            Some(storage_id),
            Some(new_storage_id),
        )?;
        set_package_last_object_identifier(package, new_storage_id)?;
        return advance_save_tokens_for_entries(package, &[storage_entry, location.archive_name]);
    }

    let (author_id, author_component_entry, created_author) = ensure_annotation_author(package)?;
    let storage_id = next_object_identifier(package)?;
    let storage_uuid = fresh_comment_storage_uuid(package)?;
    package.update_archive(&location.archive_name, |archive| {
        let mut object = ArchiveObject::new(
            storage_id,
            vec![RawMessage {
                type_: COMMENT_STORAGE_MESSAGE_TYPE,
                data: tsd::CommentStorageArchive {
                    text: Some(text),
                    creation_date: Some(current_apple_reference_date()?),
                    author: author_id.map(object_reference),
                    storage_uuid: Some(storage_uuid),
                    ..Default::default()
                }
                .encode_to_vec(),
            }],
        )?;
        if let Some(author_id) = author_id {
            object.archive_info.message_infos[0]
                .object_references
                .push(author_id);
        }
        Ok(archive.insert_object(object)?)
    })?;
    replace_drawable_comment_reference(package, application, &location, None, Some(storage_id))?;
    set_package_last_object_identifier(package, storage_id)?;
    if let (Some(author_id), Some(author_component_entry)) =
        (author_id, author_component_entry.as_deref())
    {
        let source_component = component_identifier_for_entry(package, &location.archive_name)?;
        let author_component = component_identifier_for_entry(package, author_component_entry)?;
        if let (Some(source_component), Some(author_component)) =
            (source_component, author_component)
            && source_component != author_component
        {
            add_component_external_reference(
                package,
                source_component,
                author_component,
                author_id,
            )?;
        }
    }
    let mut modified_entries = vec![location.archive_name];
    if created_author && let Some(entry) = author_component_entry {
        modified_entries.push(entry);
    }
    advance_save_tokens_for_entries(package, &modified_entries)
}

fn clear_drawable_comment_in_package(
    package: &mut IWorkPackage,
    application: Application,
    drawable_object_id: u64,
) -> Result<()> {
    let location = drawable_locations(package, application)?
        .remove(&drawable_object_id)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "drawable object {drawable_object_id} was not found"
            ))
        })?;
    let Some(storage_id) = location.comment_storage_object_id else {
        return Ok(());
    };
    // Validate the target before changing the attachment. This keeps malformed
    // comment graphs transactional instead of silently detaching corrupt data.
    let locations = object_locations(package)?;
    let root = read_comment_storage(package, &locations, storage_id)?;
    validate_direct_reply_graph(package, &locations, storage_id, &root)?;
    replace_drawable_comment_reference(package, application, &location, Some(storage_id), None)?;
    let mut removed = remove_unreferenced_comment_graph(package, application, storage_id)?;
    let mut modified_entries = vec![location.archive_name];
    for identifier in &removed.object_ids {
        if let Some(entry) = locations.get(identifier)
            && !modified_entries.contains(entry)
        {
            modified_entries.push(entry.clone());
        }
    }
    for author_id in removed.author_ids {
        if remove_generated_annotation_author_if_unused(package, author_id)? {
            if let Some(entry) = locations.get(&author_id)
                && !modified_entries.contains(entry)
            {
                modified_entries.push(entry.clone());
            }
            removed.object_ids.push(author_id);
        }
    }
    release_package_identifier_suffix(package, &removed.object_ids)?;
    advance_save_tokens_for_entries(package, &modified_entries)
}

fn add_drawable_comment_reply_in_package(
    package: &mut IWorkPackage,
    application: Application,
    drawable_object_id: u64,
    text: String,
) -> Result<u64> {
    let location = drawable_locations(package, application)?
        .remove(&drawable_object_id)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "drawable object {drawable_object_id} was not found"
            ))
        })?;
    let old_root_id = location.comment_storage_object_id.ok_or_else(|| {
        Error::ParseError(format!(
            "drawable object {drawable_object_id} has no direct comment"
        ))
    })?;
    let locations = object_locations(package)?;
    let root = read_comment_storage(package, &locations, old_root_id)?;
    validate_direct_reply_graph(package, &locations, old_root_id, &root)?;

    let (author_id, author_component_entry, created_author) = ensure_annotation_author(package)?;
    let new_root_id = next_object_identifier(package)?;
    let root_entry = clone_comment_storage_exact(package, &locations, old_root_id, new_root_id)?;
    replace_drawable_comment_reference(
        package,
        application,
        &location,
        Some(old_root_id),
        Some(new_root_id),
    )?;

    let reply_id = next_object_identifier(package)?;
    insert_comment_storage(
        package,
        &root_entry,
        reply_id,
        text,
        author_id,
        fresh_comment_storage_uuid(package)?,
    )?;
    update_comment_reply_reference(package, new_root_id, None, Some(reply_id))?;
    set_package_last_object_identifier(package, reply_id)?;

    if let (Some(author_id), Some(author_entry)) = (author_id, author_component_entry.as_deref()) {
        let source_component = component_identifier_for_entry(package, &root_entry)?;
        let author_component = component_identifier_for_entry(package, author_entry)?;
        if let (Some(source_component), Some(author_component)) =
            (source_component, author_component)
            && source_component != author_component
        {
            add_component_external_reference(
                package,
                source_component,
                author_component,
                author_id,
            )?;
        }
    }

    let mut removed = remove_unreferenced_comment_graph(package, application, old_root_id)?;
    let mut modified_entries = vec![location.archive_name, root_entry];
    if created_author && let Some(entry) = author_component_entry {
        modified_entries.push(entry);
    }
    cleanup_removed_comment_graph(package, &locations, &mut removed, &mut modified_entries)?;
    release_package_identifier_suffix(package, &removed.object_ids)?;
    advance_save_tokens_for_entries(package, &modified_entries)?;
    Ok(reply_id)
}

fn set_drawable_comment_reply_in_package(
    package: &mut IWorkPackage,
    application: Application,
    drawable_object_id: u64,
    reply_storage_object_id: u64,
    text: String,
) -> Result<u64> {
    let location = drawable_locations(package, application)?
        .remove(&drawable_object_id)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "drawable object {drawable_object_id} was not found"
            ))
        })?;
    let old_root_id = location.comment_storage_object_id.ok_or_else(|| {
        Error::ParseError(format!(
            "drawable object {drawable_object_id} has no direct comment"
        ))
    })?;
    let locations = object_locations(package)?;
    let root = read_comment_storage(package, &locations, old_root_id)?;
    validate_direct_reply_graph(package, &locations, old_root_id, &root)?;
    validate_direct_reply_reference(&root, old_root_id, reply_storage_object_id)?;
    let reply = read_comment_storage(package, &locations, reply_storage_object_id)?;
    if reply.text == text {
        return Ok(reply_storage_object_id);
    }

    let new_root_id = next_object_identifier(package)?;
    let root_entry = clone_comment_storage_exact(package, &locations, old_root_id, new_root_id)?;
    replace_drawable_comment_reference(
        package,
        application,
        &location,
        Some(old_root_id),
        Some(new_root_id),
    )?;
    let new_reply_id = next_object_identifier(package)?;
    let reply_entry =
        clone_comment_storage_exact(package, &locations, reply_storage_object_id, new_reply_id)?;
    let updated_locations = object_locations(package)?;
    update_comment_storage_text(package, &updated_locations, new_reply_id, text)?;
    update_comment_reply_reference(
        package,
        new_root_id,
        Some(reply_storage_object_id),
        Some(new_reply_id),
    )?;
    set_package_last_object_identifier(package, new_reply_id)?;

    let mut removed = remove_unreferenced_comment_graph(package, application, old_root_id)?;
    let mut modified_entries = vec![location.archive_name, root_entry, reply_entry];
    cleanup_removed_comment_graph(package, &locations, &mut removed, &mut modified_entries)?;
    release_package_identifier_suffix(package, &removed.object_ids)?;
    advance_save_tokens_for_entries(package, &modified_entries)?;
    Ok(new_reply_id)
}

fn remove_drawable_comment_reply_in_package(
    package: &mut IWorkPackage,
    application: Application,
    drawable_object_id: u64,
    reply_storage_object_id: u64,
) -> Result<()> {
    let location = drawable_locations(package, application)?
        .remove(&drawable_object_id)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "drawable object {drawable_object_id} was not found"
            ))
        })?;
    let old_root_id = location.comment_storage_object_id.ok_or_else(|| {
        Error::ParseError(format!(
            "drawable object {drawable_object_id} has no direct comment"
        ))
    })?;
    let locations = object_locations(package)?;
    let root = read_comment_storage(package, &locations, old_root_id)?;
    validate_direct_reply_graph(package, &locations, old_root_id, &root)?;
    validate_direct_reply_reference(&root, old_root_id, reply_storage_object_id)?;
    read_comment_storage(package, &locations, reply_storage_object_id)?;

    let new_root_id = next_object_identifier(package)?;
    let root_entry = clone_comment_storage_exact(package, &locations, old_root_id, new_root_id)?;
    replace_drawable_comment_reference(
        package,
        application,
        &location,
        Some(old_root_id),
        Some(new_root_id),
    )?;
    update_comment_reply_reference(package, new_root_id, Some(reply_storage_object_id), None)?;
    set_package_last_object_identifier(package, new_root_id)?;

    let mut removed = remove_unreferenced_comment_graph(package, application, old_root_id)?;
    let mut modified_entries = vec![location.archive_name, root_entry];
    cleanup_removed_comment_graph(package, &locations, &mut removed, &mut modified_entries)?;
    release_package_identifier_suffix(package, &removed.object_ids)?;
    advance_save_tokens_for_entries(package, &modified_entries)
}

fn validate_direct_reply_reference(
    root: &Comment,
    root_storage_id: u64,
    reply_storage_id: u64,
) -> Result<()> {
    if root_storage_id == reply_storage_id {
        return Err(Error::InvalidFormat(format!(
            "comment storage {root_storage_id} references itself as a reply"
        )));
    }
    match root
        .reply_ids
        .iter()
        .filter(|identifier| identifier.get() == reply_storage_id)
        .count()
    {
        1 => Ok(()),
        0 => Err(Error::ParseError(format!(
            "comment storage {reply_storage_id} is not a direct reply to {root_storage_id}"
        ))),
        _ => Err(Error::InvalidFormat(format!(
            "comment storage {root_storage_id} duplicates reply {reply_storage_id}"
        ))),
    }
}

fn validate_direct_reply_graph(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    root_storage_id: u64,
    root: &Comment,
) -> Result<()> {
    let mut seen = HashSet::new();
    seen.try_reserve(root.reply_ids.len()).map_err(|_| {
        comment_storage_allocation_error(
            "iWork direct comment reply identifiers",
            root.reply_ids.len(),
        )
    })?;
    for reply_id in &root.reply_ids {
        let reply_id = reply_id.get();
        if reply_id == root_storage_id || !seen.insert(reply_id) {
            return Err(Error::InvalidFormat(format!(
                "comment storage {root_storage_id} contains a duplicate or cyclic reply reference to {reply_id}"
            )));
        }
        read_comment_storage(package, locations, reply_id)?;
    }
    Ok(())
}

fn cleanup_removed_comment_graph(
    package: &mut IWorkPackage,
    original_locations: &HashMap<u64, String>,
    removed: &mut RemovedCommentGraph,
    modified_entries: &mut Vec<String>,
) -> Result<()> {
    for identifier in &removed.object_ids {
        if let Some(entry) = original_locations.get(identifier)
            && !modified_entries.contains(entry)
        {
            modified_entries.push(entry.clone());
        }
    }
    for author_id in std::mem::take(&mut removed.author_ids) {
        if remove_generated_annotation_author_if_unused(package, author_id)? {
            if let Some(entry) = original_locations.get(&author_id)
                && !modified_entries.contains(entry)
            {
                modified_entries.push(entry.clone());
            }
            removed.object_ids.push(author_id);
        }
    }
    Ok(())
}

pub(crate) fn advance_save_tokens_for_entries(
    package: &mut IWorkPackage,
    entry_names: &[String],
) -> Result<()> {
    let mut component_identifiers = Vec::new();
    for entry_name in entry_names {
        if let Some(identifier) = component_identifier_for_entry(package, entry_name)?
            && !component_identifiers.contains(&identifier)
        {
            component_identifiers.push(identifier);
        }
    }
    advance_package_save_token_for_components(package, &component_identifiers)
}

fn replace_drawable_comment_reference(
    package: &mut IWorkPackage,
    application: Application,
    location: &DrawableLocation,
    old: Option<u64>,
    new: Option<u64>,
) -> Result<()> {
    package.update_archive(&location.archive_name, |archive| {
        let object = archive.object_mut(location.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("drawable object {} is missing", location.object_id))
        })?;
        let message = object.messages.get(location.message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "drawable object {} lost payload {}",
                location.object_id, location.message_index
            ))
        })?;
        let payload =
            DrawableCommentProjection::decode(application, message.type_, message.data.as_slice())?
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "object {} payload {} is no longer a drawable",
                        location.object_id, location.message_index
                    ))
                })?;
        if payload.comment_identifier != old {
            return Err(Error::InvalidFormat(format!(
                "drawable object {} comment changed during mutation",
                location.object_id
            )));
        }
        let message_type = message.type_;
        let data = match (old, new) {
            // Keep the existing TSP.Reference as the preservation
            // representation.  Re-encoding it here would silently discard
            // producer extensions and unknown fields nested inside the leaf.
            (Some(_), Some(identifier)) => {
                let (path, path_len) = drawable_comment_identifier_path(payload.comment_wire_path)?;
                patch_nested_varint_field(
                    message.data.as_slice(),
                    &path[..path_len],
                    true,
                    Some(identifier),
                )?
            },
            // A newly attached reference has no source leaf to preserve.
            (None, Some(identifier)) => {
                let replacement = tsp::Reference {
                    identifier,
                    ..Default::default()
                }
                .encode_to_vec();
                patch_nested_length_delimited_field(
                    message.data.as_slice(),
                    payload.comment_wire_path,
                    false,
                    Some(replacement.as_slice()),
                )?
            },
            (Some(_), None) | (None, None) => patch_nested_length_delimited_field(
                message.data.as_slice(),
                payload.comment_wire_path,
                old.is_some(),
                None,
            )?,
        };
        let verified =
            DrawableCommentProjection::decode(application, message_type, data.as_slice())?
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "object {} stopped decoding as a drawable after comment patch",
                        location.object_id
                    ))
                })?;
        if verified.comment_identifier != new {
            return Err(Error::InvalidFormat(format!(
                "drawable object {} comment patch failed validation",
                location.object_id
            )));
        }
        object.replace_message(
            location.message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        update_reference_metadata(
            &mut object.archive_info.message_infos[location.message_index],
            old,
            new,
        );
        Ok(())
    })
}

/// Append the identifier leaf to a supported drawable comment route without
/// allocating.  The route depth is format-owned and currently at most four;
/// refusing a future/unknown route keeps this host-side patch bounded.
fn drawable_comment_identifier_path(path: &[u32]) -> Result<([u32; 5], usize)> {
    let path_len = path.len().checked_add(1).ok_or_else(|| {
        Error::InvalidFormat("drawable comment identifier path length overflow".to_owned())
    })?;
    if path_len > 5 {
        return Err(Error::InvalidFormat(format!(
            "drawable comment identifier path is too deep: {path_len}"
        )));
    }
    let mut identifier_path = [0; 5];
    identifier_path[..path.len()].copy_from_slice(path);
    identifier_path[path.len()] = COMMENT_REFERENCE_IDENTIFIER_FIELD;
    Ok((identifier_path, path_len))
}

fn update_reference_metadata(
    info: &mut crate::archive::MessageInfo,
    old: Option<u64>,
    new: Option<u64>,
) {
    update_reference_list(&mut info.object_references, old, new);
    for field in &mut info.field_infos {
        // Existing field-level references must not become stale, but a new
        // reference cannot be assigned to an arbitrary field path safely.
        update_reference_list(&mut field.object_references, old, None);
    }
}

fn update_reference_list(references: &mut Vec<u64>, old: Option<u64>, new: Option<u64>) {
    if let Some(old) = old {
        references.retain(|reference| *reference != old);
    }
    if let Some(new) = new
        && !references.contains(&new)
    {
        references.push(new);
    }
}

fn comment_storage_message_index(object: &ArchiveObject, storage_id: u64) -> Result<usize> {
    let mut index = None;
    for (candidate, message) in object.messages.iter().enumerate() {
        if message.type_ != COMMENT_STORAGE_MESSAGE_TYPE {
            continue;
        }
        if index.replace(candidate).is_some() {
            return Err(Error::InvalidFormat(format!(
                "object {storage_id} must contain exactly one TSD comment-storage payload"
            )));
        }
    }
    index.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "object {storage_id} must contain exactly one TSD comment-storage payload"
        ))
    })
}

fn read_comment_storage(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    storage_id: u64,
) -> Result<Comment> {
    let archive_name = locations.get(&storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!("comment storage object {storage_id} is missing"))
    })?;
    package.with_parsed_archive(archive_name, |archive| {
        let object = archive.object(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("comment storage object {storage_id} is missing"))
        })?;
        let index = comment_storage_message_index(object, storage_id)?;
        let (comment, reply_ids) =
            decode_comment_storage_payload(storage_id, object.messages[index].data.as_slice())?;
        let author_id = comment
            .author()
            .map(|author| author_id_from_raw(author.identifier()))
            .transpose()?;
        let mut typed_reply_ids = Vec::new();
        for reply_id in &reply_ids {
            // Preserve malformed-ID precedence before requesting the full
            // typed collection's capacity.
            storage_id_from_raw(*reply_id)?;
        }
        typed_reply_ids
            .try_reserve_exact(reply_ids.len())
            .map_err(|_| {
                comment_storage_allocation_error("iWork comment reply identifiers", reply_ids.len())
            })?;
        for reply_id in reply_ids {
            typed_reply_ids.push(storage_id_from_raw(reply_id)?);
        }
        let reply_ids = typed_reply_ids.into_boxed_slice();
        let text = materialize_comment_text(storage_id, comment.text().unwrap_or_default())?;
        let storage_uuid = comment
            .storage_uuid()
            .map(|uuid| comment_uuid(uuid.lower(), uuid.upper()))
            .transpose()?;
        Ok(Comment {
            text,
            creation_date_seconds: comment.creation_date().map(|date| date.seconds()),
            author_id,
            reply_ids,
            storage_uuid,
        })
    })
}

/// Materialize the borrowed strict-codec text only after validation succeeds.
///
/// The codec's text budget is bounded by the source payload, while the common
/// comment model owns its text.  Keep that ownership transition fallible so a
/// hostile payload cannot turn the final host-side copy into an infallible
/// allocation.  Unknown wire bytes remain in the source payload and are never
/// reconstructed by this projection.
fn materialize_comment_text(storage_id: u64, text: &str) -> Result<String> {
    let mut materialized = String::new();
    materialized
        .try_reserve_exact(text.len())
        .map_err(|_| comment_storage_allocation_error("iWork comment text", text.len()))?;
    materialized.push_str(text);
    if materialized.len() != text.len() {
        return Err(Error::InvalidFormat(format!(
            "comment storage object {storage_id} text materialization changed length"
        )));
    }
    Ok(materialized)
}

fn update_comment_storage_text(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    storage_id: u64,
    text: String,
) -> Result<()> {
    let archive_name = locations.get(&storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!("comment storage object {storage_id} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("comment storage object {storage_id} is missing"))
        })?;
        let index = comment_storage_message_index(object, storage_id)?;
        let original = object.messages[index].data.as_slice();
        let (comment, _) = decode_comment_storage_payload(storage_id, original)?;
        let data = patch_length_delimited_field(
            original,
            1,
            comment.text().is_some(),
            Some(text.as_bytes()),
        )?;
        let (verified, _) = decode_comment_storage_payload(storage_id, data.as_slice())?;
        if verified.text() != Some(text.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "comment storage object {storage_id} text patch failed validation"
            )));
        }
        object.replace_message(
            index,
            RawMessage {
                type_: COMMENT_STORAGE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn clone_comment_storage(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    old_storage_id: u64,
    new_storage_id: u64,
    storage_uuid: tsp::Uuid,
    text: String,
) -> Result<()> {
    let archive_name = locations.get(&old_storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "comment storage object {old_storage_id} is missing"
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let source = archive.object(old_storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "comment storage object {old_storage_id} is missing"
        ))
    })?;
    if source.messages.len() != 1 || source.messages[0].type_ != COMMENT_STORAGE_MESSAGE_TYPE {
        return Err(Error::InvalidFormat(format!(
            "cannot safely clone multi-payload comment object {old_storage_id}"
        )));
    }
    let (comment, _) =
        decode_comment_storage_payload(old_storage_id, source.messages[0].data.as_slice())?;
    let data = patch_length_delimited_field(
        source.messages[0].data.as_slice(),
        1,
        comment.text().is_some(),
        Some(text.as_bytes()),
    )?;
    let uuid = storage_uuid.encode_to_vec();
    let data = patch_length_delimited_field(
        data.as_slice(),
        5,
        comment.storage_uuid().is_some(),
        Some(uuid.as_slice()),
    )?;
    let (verified, _) = decode_comment_storage_payload(new_storage_id, data.as_slice())?;
    let verified_uuid = verified
        .storage_uuid()
        .map(|uuid| (uuid.lower(), uuid.upper()));
    if verified.text() != Some(text.as_str())
        || verified_uuid != Some((storage_uuid.lower, storage_uuid.upper))
    {
        return Err(Error::InvalidFormat(format!(
            "comment storage clone {new_storage_id} failed validation"
        )));
    }
    let mut clone = ArchiveObject::new(
        new_storage_id,
        vec![RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data,
        }],
    )?;
    clone.archive_info.should_merge = source.archive_info.should_merge;
    clone.archive_info.message_infos[0] = source.archive_info.message_infos[0].clone();
    clone.archive_info.message_infos[0].length = u32::try_from(clone.messages[0].data.len())
        .map_err(|_| Error::Archive("comment payload exceeds the u32 format limit".to_owned()))?;
    package.update_archive(archive_name, |archive| Ok(archive.insert_object(clone)?))
}

pub(crate) fn clone_comment_storage_exact(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    old_storage_id: u64,
    new_storage_id: u64,
) -> Result<String> {
    let archive_name = locations.get(&old_storage_id).cloned().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "comment storage object {old_storage_id} is missing"
        ))
    })?;
    let archive = package.archive(&archive_name)?;
    let source = archive.object(old_storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "comment storage object {old_storage_id} is missing"
        ))
    })?;
    if source.messages.len() != 1 || source.messages[0].type_ != COMMENT_STORAGE_MESSAGE_TYPE {
        return Err(Error::InvalidFormat(format!(
            "cannot safely clone multi-payload comment object {old_storage_id}"
        )));
    }
    validate_comment_storage_payload(old_storage_id, source.messages[0].data.as_slice())?;
    let mut clone = ArchiveObject::new(new_storage_id, source.messages.clone())?;
    clone.archive_info.should_merge = source.archive_info.should_merge;
    clone.archive_info.message_infos = source.archive_info.message_infos.clone();
    package.update_archive(&archive_name, |archive| Ok(archive.insert_object(clone)?))?;
    Ok(archive_name)
}

pub(crate) fn insert_comment_storage(
    package: &mut IWorkPackage,
    archive_name: &str,
    storage_id: u64,
    text: String,
    author_id: Option<u64>,
    storage_uuid: tsp::Uuid,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let mut object = ArchiveObject::new(
            storage_id,
            vec![RawMessage {
                type_: COMMENT_STORAGE_MESSAGE_TYPE,
                data: tsd::CommentStorageArchive {
                    text: Some(text),
                    creation_date: Some(current_apple_reference_date()?),
                    author: author_id.map(object_reference),
                    storage_uuid: Some(storage_uuid),
                    ..Default::default()
                }
                .encode_to_vec(),
            }],
        )?;
        if let Some(author_id) = author_id {
            object.archive_info.message_infos[0]
                .object_references
                .push(author_id);
        }
        Ok(archive.insert_object(object)?)
    })
}

pub(crate) fn update_comment_reply_reference(
    package: &mut IWorkPackage,
    root_storage_id: u64,
    old_reply_id: Option<u64>,
    new_reply_id: Option<u64>,
) -> Result<()> {
    let locations = object_locations(package)?;
    let archive_name = locations.get(&root_storage_id).cloned().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "comment storage object {root_storage_id} is missing"
        ))
    })?;
    package.update_archive(&archive_name, |archive| {
        let object = archive.object_mut(root_storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "comment storage object {root_storage_id} is missing"
            ))
        })?;
        let index = comment_storage_message_index(object, root_storage_id)?;
        let original = object.messages[index].data.as_slice();
        let (_, before_replies) = decode_comment_storage_payload(root_storage_id, original)?;
        let old_count = old_reply_id.map_or(0, |identifier| {
            before_replies
                .iter()
                .filter(|reply| **reply == identifier)
                .count()
        });
        if old_reply_id.is_some() && old_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "comment storage {root_storage_id} must reference reply {} exactly once",
                old_reply_id.unwrap_or_default()
            )));
        }
        if let Some(identifier) = new_reply_id
            && Some(identifier) != old_reply_id
            && before_replies.contains(&identifier)
        {
            return Err(Error::InvalidFormat(format!(
                "comment storage {root_storage_id} already references reply {identifier}"
            )));
        }
        let data = match (old_reply_id, new_reply_id) {
            (None, Some(identifier)) => append_repeated_length_delimited_field(
                original,
                4,
                &object_reference(identifier).encode_to_vec(),
            )?,
            (Some(old), Some(new)) => {
                let mut reply_index = 0;
                let data = transform_length_delimited_fields_at_path(original, &[4], |payload| {
                    let identifier = *before_replies.get(reply_index).ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "comment storage {root_storage_id} reply count changed during patch"
                        ))
                    })?;
                    reply_index += 1;
                    if identifier == old {
                        patch_varint_field(payload, 1, true, Some(new))
                    } else {
                        Ok(payload.to_vec())
                    }
                })?;
                if reply_index != before_replies.len() {
                    return Err(Error::InvalidFormat(format!(
                        "comment storage {root_storage_id} reply count changed during patch"
                    )));
                }
                data
            },
            (Some(identifier), None) => {
                let mut reply_index = 0;
                let data = remove_repeated_length_delimited_field_where(original, 4, |_payload| {
                    let current = *before_replies.get(reply_index).ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "comment storage {root_storage_id} reply count changed during patch"
                        ))
                    })?;
                    reply_index += 1;
                    Ok(current == identifier)
                })?;
                if reply_index != before_replies.len() {
                    return Err(Error::InvalidFormat(format!(
                        "comment storage {root_storage_id} reply count changed during patch"
                    )));
                }
                data
            },
            (None, None) => return Ok(()),
        };
        let (_, verified_replies) =
            decode_comment_storage_payload(root_storage_id, data.as_slice())?;
        let expected = before_replies
            .iter()
            .filter_map(|identifier| {
                if old_reply_id == Some(*identifier) {
                    new_reply_id
                } else {
                    Some(*identifier)
                }
            })
            .chain((old_reply_id.is_none()).then_some(new_reply_id).flatten());
        if !verified_replies.iter().copied().eq(expected) {
            return Err(Error::InvalidFormat(format!(
                "comment storage {root_storage_id} reply patch failed validation"
            )));
        }
        object.replace_message(
            index,
            RawMessage {
                type_: COMMENT_STORAGE_MESSAGE_TYPE,
                data,
            },
        )?;
        update_reference_metadata(
            &mut object.archive_info.message_infos[index],
            old_reply_id,
            new_reply_id,
        );
        Ok(())
    })
}

/// Borrowed, generated-free facts from one annotation-author payload.
///
/// The source message remains the preservation representation.  In
/// particular, this projection never reconstructs an author or color, so
/// unknown fields survive every storage-list patch byte-for-byte.
#[derive(Debug, PartialEq)]
struct AnnotationAuthorProjection<'source> {
    name: Option<&'source str>,
    color: Option<AnnotationColorProjection>,
    public_id: Option<&'source str>,
    is_public_author: Option<bool>,
    public_ids: Vec<&'source str>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct AnnotationColorProjection {
    model: i32,
    red: Option<u32>,
    green: Option<u32>,
    blue: Option<u32>,
    alpha: Option<u32>,
    cyan: Option<u32>,
    magenta: Option<u32>,
    yellow: Option<u32>,
    black: Option<u32>,
    white: Option<u32>,
    rgbspace: Option<i32>,
}

impl AnnotationColorProjection {
    fn generated() -> Self {
        Self {
            model: tsp::color::ColorModel::Rgb as i32,
            red: Some(0.368_627_46_f32.to_bits()),
            green: Some(0.568_627_5_f32.to_bits()),
            blue: Some(0.937_254_9_f32.to_bits()),
            alpha: Some(1.0_f32.to_bits()),
            cyan: None,
            magenta: None,
            yellow: None,
            black: None,
            white: None,
            rgbspace: Some(tsp::color::RgbColorSpace::Srgb as i32),
        }
    }
}

impl AnnotationAuthorProjection<'_> {
    fn is_generated(&self, public_id: bool) -> bool {
        self.name == Some(GENERATED_AUTHOR_NAME)
            && self.color == Some(AnnotationColorProjection::generated())
            && self.public_id == public_id.then_some(GENERATED_AUTHOR_PUBLIC_ID)
            && self.is_public_author == Some(false)
            && if public_id {
                self.public_ids == [GENERATED_AUTHOR_PUBLIC_ID]
            } else {
                self.public_ids.is_empty()
            }
    }
}

fn strict_annotation_author_error(object_id: u64, error: impl std::fmt::Display) -> Error {
    Error::InvalidFormat(format!(
        "annotation author object {object_id} failed strict validation: {error}"
    ))
}

/// Parse one annotation payload's fields under the shared finite wire budget.
///
/// The generated author messages are not used as a preservation
/// representation.  Keep the source bytes authoritative and retain the
/// common parser's typed allocation error instead of turning an exhausted
/// candidate-local vector into a generic malformed-message error.
fn annotation_wire_fields(object_id: u64, source: &[u8]) -> Result<Vec<crate::wire::WireField>> {
    match litchi_iwa_common::wire::parse_wire_fields_with_limits(
        source,
        litchi_iwa_common::WireLimits::default(),
    ) {
        Ok(fields) => Ok(fields),
        Err(litchi_iwa_common::Error::Allocation { resource, amount }) => {
            Err(Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource,
                amount,
            }))
        },
        Err(error) => Err(strict_annotation_author_error(object_id, error)),
    }
}

fn canonical_annotation_field(
    object_id: u64,
    source: &[u8],
    field: crate::wire::WireField,
) -> Result<()> {
    field
        .validate_canonical_key(source)
        .map_err(|error| strict_annotation_author_error(object_id, error))?;
    match field.wire_type() {
        0 => {
            let payload = field
                .payload(source)
                .map_err(|error| strict_annotation_author_error(object_id, error))?;
            let (value, width) = litchi_iwa_common::varint::decode_varint_from_bytes(payload)
                .map_err(|error| strict_annotation_author_error(object_id, error))?;
            if width != payload.len() || width != litchi_iwa_common::varint::encoded_len(value) {
                return Err(strict_annotation_author_error(
                    object_id,
                    format!(
                        "protobuf field {} has a noncanonical varint value",
                        field.number()
                    ),
                ));
            }
        },
        1 => {
            let payload = field
                .payload(source)
                .map_err(|error| strict_annotation_author_error(object_id, error))?;
            if payload.len() != 8 {
                return Err(strict_annotation_author_error(
                    object_id,
                    format!(
                        "protobuf field {} has an invalid fixed64 width",
                        field.number()
                    ),
                ));
            }
        },
        2 => field
            .validate_canonical_length(source)
            .map_err(|error| strict_annotation_author_error(object_id, error))?,
        5 => {
            let payload = field
                .payload(source)
                .map_err(|error| strict_annotation_author_error(object_id, error))?;
            if payload.len() != 4 {
                return Err(strict_annotation_author_error(
                    object_id,
                    format!(
                        "protobuf field {} has an invalid fixed32 width",
                        field.number()
                    ),
                ));
            }
        },
        wire_type => {
            return Err(strict_annotation_author_error(
                object_id,
                format!(
                    "protobuf field {} has unsupported wire type {wire_type}",
                    field.number()
                ),
            ));
        },
    }
    Ok(())
}

fn annotation_length(
    object_id: u64,
    source: &[u8],
    field: crate::wire::WireField,
) -> Result<&[u8]> {
    if field.wire_type() != 2 {
        return Err(strict_annotation_author_error(
            object_id,
            format!("protobuf field {} is not length-delimited", field.number()),
        ));
    }
    field
        .payload(source)
        .map_err(|error| strict_annotation_author_error(object_id, error))
}

fn annotation_varint(object_id: u64, source: &[u8], field: crate::wire::WireField) -> Result<u64> {
    if field.wire_type() != 0 {
        return Err(strict_annotation_author_error(
            object_id,
            format!("protobuf field {} is not a varint", field.number()),
        ));
    }
    let payload = field
        .payload(source)
        .map_err(|error| strict_annotation_author_error(object_id, error))?;
    litchi_iwa_common::varint::decode_varint_from_bytes(payload)
        .map_err(|error| strict_annotation_author_error(object_id, error))
        .and_then(|(value, width)| {
            if width != payload.len() || width != litchi_iwa_common::varint::encoded_len(value) {
                Err(strict_annotation_author_error(
                    object_id,
                    format!(
                        "protobuf field {} has a noncanonical varint value",
                        field.number()
                    ),
                ))
            } else {
                Ok(value)
            }
        })
}

fn annotation_fixed32(object_id: u64, source: &[u8], field: crate::wire::WireField) -> Result<u32> {
    if field.wire_type() != 5 {
        return Err(strict_annotation_author_error(
            object_id,
            format!("protobuf field {} is not a fixed32", field.number()),
        ));
    }
    let payload = field
        .payload(source)
        .map_err(|error| strict_annotation_author_error(object_id, error))?;
    let bytes: [u8; 4] = payload.try_into().map_err(|_error| {
        strict_annotation_author_error(
            object_id,
            format!(
                "protobuf field {} has an invalid fixed32 width",
                field.number()
            ),
        )
    })?;
    Ok(u32::from_le_bytes(bytes))
}

fn annotation_int32(object_id: u64, source: &[u8], field: crate::wire::WireField) -> Result<i32> {
    let value = annotation_varint(object_id, source, field)?;
    if let Ok(value) = i32::try_from(value) {
        return Ok(value);
    }
    if value < MIN_SIGN_EXTENDED_INT32 {
        return Err(strict_annotation_author_error(
            object_id,
            format!(
                "protobuf field {} has an out-of-range int32",
                field.number()
            ),
        ));
    }
    i32::try_from(i64::from_ne_bytes(value.to_ne_bytes())).map_err(|_error| {
        strict_annotation_author_error(
            object_id,
            format!(
                "protobuf field {} has an out-of-range int32",
                field.number()
            ),
        )
    })
}

fn annotation_bool(object_id: u64, source: &[u8], field: crate::wire::WireField) -> Result<bool> {
    match annotation_varint(object_id, source, field)? {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(strict_annotation_author_error(
            object_id,
            format!(
                "protobuf field {} has a noncanonical bool value",
                field.number()
            ),
        )),
    }
}

fn annotation_string(object_id: u64, source: &[u8], field: crate::wire::WireField) -> Result<&str> {
    let payload = annotation_length(object_id, source, field)?;
    str::from_utf8(payload).map_err(|error| strict_annotation_author_error(object_id, error))
}

fn decode_annotation_color(object_id: u64, source: &[u8]) -> Result<AnnotationColorProjection> {
    let fields = annotation_wire_fields(object_id, source)?;
    let mut model = None;
    let mut red = None;
    let mut green = None;
    let mut blue = None;
    let mut alpha = None;
    let mut cyan = None;
    let mut magenta = None;
    let mut yellow = None;
    let mut black = None;
    let mut white = None;
    let mut rgbspace = None;
    for field in fields {
        canonical_annotation_field(object_id, source, field)?;
        match field.number() {
            COLOR_MODEL_FIELD
                if model
                    .replace(annotation_int32(object_id, source, field)?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSP.Color.model",
                ));
            },
            COLOR_MODEL_FIELD => {},
            COLOR_RED_FIELD
                if red
                    .replace(annotation_fixed32(object_id, source, field)?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSP.Color.r",
                ));
            },
            COLOR_RED_FIELD => {},
            COLOR_GREEN_FIELD
                if green
                    .replace(annotation_fixed32(object_id, source, field)?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSP.Color.g",
                ));
            },
            COLOR_GREEN_FIELD => {},
            COLOR_BLUE_FIELD
                if blue
                    .replace(annotation_fixed32(object_id, source, field)?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSP.Color.b",
                ));
            },
            COLOR_BLUE_FIELD => {},
            COLOR_ALPHA_FIELD
                if alpha
                    .replace(annotation_fixed32(object_id, source, field)?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSP.Color.a",
                ));
            },
            COLOR_ALPHA_FIELD => {},
            COLOR_CYAN_FIELD
                if cyan
                    .replace(annotation_fixed32(object_id, source, field)?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSP.Color.c",
                ));
            },
            COLOR_CYAN_FIELD => {},
            COLOR_MAGENTA_FIELD
                if magenta
                    .replace(annotation_fixed32(object_id, source, field)?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSP.Color.m",
                ));
            },
            COLOR_MAGENTA_FIELD => {},
            COLOR_YELLOW_FIELD
                if yellow
                    .replace(annotation_fixed32(object_id, source, field)?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSP.Color.y",
                ));
            },
            COLOR_YELLOW_FIELD => {},
            COLOR_BLACK_FIELD
                if black
                    .replace(annotation_fixed32(object_id, source, field)?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSP.Color.k",
                ));
            },
            COLOR_BLACK_FIELD => {},
            COLOR_WHITE_FIELD
                if white
                    .replace(annotation_fixed32(object_id, source, field)?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSP.Color.w",
                ));
            },
            COLOR_WHITE_FIELD => {},
            COLOR_RGBSPACE_FIELD
                if rgbspace
                    .replace(annotation_int32(object_id, source, field)?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSP.Color.rgbspace",
                ));
            },
            COLOR_RGBSPACE_FIELD => {},
            _ => {},
        }
    }
    Ok(AnnotationColorProjection {
        model: model.ok_or_else(|| {
            strict_annotation_author_error(object_id, "missing required TSP.Color.model")
        })?,
        red,
        green,
        blue,
        alpha,
        cyan,
        magenta,
        yellow,
        black,
        white,
        rgbspace,
    })
}

fn decode_annotation_author<'source>(
    object_id: u64,
    source: &'source [u8],
) -> Result<AnnotationAuthorProjection<'source>> {
    let fields = annotation_wire_fields(object_id, source)?;
    let mut name = None;
    let mut color = None;
    let mut public_id = None;
    let mut is_public_author = None;
    let mut public_ids = Vec::new();
    let mut allocation_failed = None;
    for field in fields {
        canonical_annotation_field(object_id, source, field)?;
        match field.number() {
            ANNOTATION_AUTHOR_NAME_FIELD
                if name
                    .replace(annotation_string(object_id, source, field)?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSK.AnnotationAuthorArchive.name",
                ));
            },
            ANNOTATION_AUTHOR_NAME_FIELD => {},
            ANNOTATION_AUTHOR_COLOR_FIELD
                if color
                    .replace(decode_annotation_color(
                        object_id,
                        annotation_length(object_id, source, field)?,
                    )?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSK.AnnotationAuthorArchive.color",
                ));
            },
            ANNOTATION_AUTHOR_COLOR_FIELD => {},
            ANNOTATION_AUTHOR_PUBLIC_ID_FIELD
                if public_id
                    .replace(annotation_string(object_id, source, field)?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSK.AnnotationAuthorArchive.public_id",
                ));
            },
            ANNOTATION_AUTHOR_PUBLIC_ID_FIELD => {},
            ANNOTATION_AUTHOR_IS_PUBLIC_FIELD
                if is_public_author
                    .replace(annotation_bool(object_id, source, field)?)
                    .is_some() =>
            {
                return Err(strict_annotation_author_error(
                    object_id,
                    "duplicate TSK.AnnotationAuthorArchive.is_public_author",
                ));
            },
            ANNOTATION_AUTHOR_IS_PUBLIC_FIELD => {},
            ANNOTATION_AUTHOR_PUBLIC_IDS_FIELD => {
                let value = annotation_string(object_id, source, field)?;
                if allocation_failed.is_none() {
                    if public_ids.try_reserve(1).is_err() {
                        allocation_failed = Some(public_ids.len().saturating_add(1));
                    } else {
                        public_ids.push(value);
                    }
                }
            },
            _ => {},
        }
    }
    if let Some(amount) = allocation_failed {
        return Err(comment_storage_allocation_error(
            "iWork annotation-author public IDs",
            amount,
        ));
    }
    Ok(AnnotationAuthorProjection {
        name,
        color,
        public_id,
        is_public_author,
        public_ids,
    })
}

fn annotation_author_storage_ids(storage_id: u64, source: &[u8]) -> Result<Vec<u64>> {
    let fields = annotation_wire_fields(storage_id, source)?;
    let mut identifiers = Vec::new();
    let mut allocation_failed = None;
    for field in fields {
        canonical_annotation_field(storage_id, source, field)?;
        if field.number() != 1 {
            continue;
        }
        let payload = annotation_length(storage_id, source, field)?;
        let reference = comment_storage_codec::decode_reference(
            payload,
            comment_storage_decode_options(payload),
        )
        .map_err(|error| strict_annotation_author_error(storage_id, error))?;
        let identifier = reference.identifier();
        // Validate the typed identity before a candidate-local collection
        // allocation so malformed IDs retain precedence over exhaustion.
        author_id_from_raw(identifier)?;
        if allocation_failed.is_none() {
            if identifiers.try_reserve(1).is_err() {
                allocation_failed = Some(identifiers.len().saturating_add(1));
            } else {
                identifiers.push(identifier);
            }
        }
    }
    if let Some(amount) = allocation_failed {
        return Err(comment_storage_allocation_error(
            "iWork annotation-author identifiers",
            amount,
        ));
    }
    Ok(identifiers)
}

fn annotation_author_is_generated(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    author_id: u64,
    public_id: bool,
) -> Result<bool> {
    let archive_name = locations.get(&author_id).ok_or_else(|| {
        Error::InvalidFormat(format!("annotation author object {author_id} is missing"))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(author_id).ok_or_else(|| {
        Error::InvalidFormat(format!("annotation author object {author_id} is missing"))
    })?;
    let matching_message_count = object
        .messages
        .iter()
        .filter(|message| message.type_ == ANNOTATION_AUTHOR_MESSAGE_TYPE)
        .count();
    let mut messages = Vec::new();
    messages
        .try_reserve_exact(matching_message_count)
        .map_err(|_| {
            comment_storage_allocation_error(
                "iWork annotation-author payloads",
                matching_message_count,
            )
        })?;
    messages.extend(
        object
            .messages
            .iter()
            .filter(|message| message.type_ == ANNOTATION_AUTHOR_MESSAGE_TYPE),
    );
    if messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "object {author_id} must contain exactly one annotation-author payload"
        )));
    }
    let projection = decode_annotation_author(author_id, messages[0].data.as_slice())?;
    Ok(projection.is_generated(public_id))
}

fn validate_annotation_author(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    author_id: u64,
) -> Result<()> {
    let archive_name = locations.get(&author_id).ok_or_else(|| {
        Error::InvalidFormat(format!("annotation author object {author_id} is missing"))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(author_id).ok_or_else(|| {
        Error::InvalidFormat(format!("annotation author object {author_id} is missing"))
    })?;
    let matching_message_count = object
        .messages
        .iter()
        .filter(|message| message.type_ == ANNOTATION_AUTHOR_MESSAGE_TYPE)
        .count();
    let mut messages = Vec::new();
    messages
        .try_reserve_exact(matching_message_count)
        .map_err(|_| {
            comment_storage_allocation_error(
                "iWork annotation-author payloads",
                matching_message_count,
            )
        })?;
    messages.extend(
        object
            .messages
            .iter()
            .filter(|message| message.type_ == ANNOTATION_AUTHOR_MESSAGE_TYPE),
    );
    if messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "object {author_id} must contain exactly one annotation-author payload"
        )));
    }
    decode_annotation_author(author_id, messages[0].data.as_slice()).map(|_| ())
}

struct AnnotationAuthorStorageLocation {
    archive_name: String,
    object_id: u64,
    message_index: usize,
}

fn object_reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

pub(crate) fn fresh_comment_storage_uuid(package: &IWorkPackage) -> Result<tsp::Uuid> {
    let mut existing = HashSet::new();
    for name in package.iwa_entry_names() {
        for object in package.archive(name)?.objects {
            let object_id = object.archive_info.identifier.ok_or_else(|| {
                Error::Archive(format!("object in {name} has no archive identifier"))
            })?;
            for message in object
                .messages
                .iter()
                .filter(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            {
                let (comment, _) =
                    decode_comment_storage_payload(object_id, message.data.as_slice())?;
                if let Some(uuid) = comment.storage_uuid() {
                    existing.insert((uuid.lower(), uuid.upper()));
                }
            }
        }
    }
    loop {
        let bytes = litchi_core::id::generate_guid_bytes();
        let mut lower = [0u8; 8];
        lower.copy_from_slice(&bytes[..8]);
        let mut upper = [0u8; 8];
        upper.copy_from_slice(&bytes[8..]);
        let uuid = tsp::Uuid {
            lower: u64::from_le_bytes(lower),
            upper: u64::from_le_bytes(upper),
        };
        if existing.insert((uuid.lower, uuid.upper)) {
            return Ok(uuid);
        }
    }
}

fn annotation_author_storage_location(
    package: &IWorkPackage,
) -> Result<Option<AnnotationAuthorStorageLocation>> {
    let mut result = None;
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        for object in &archive.objects {
            let object_id = object.archive_info.identifier.ok_or_else(|| {
                Error::Archive(format!("object in {name} has no archive identifier"))
            })?;
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ != ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE {
                    continue;
                }
                annotation_author_storage_ids(object_id, message.data.as_slice())?;
                if result
                    .replace(AnnotationAuthorStorageLocation {
                        archive_name: name.to_owned(),
                        object_id,
                        message_index,
                    })
                    .is_some()
                {
                    return Err(Error::InvalidFormat(
                        "package contains multiple annotation-author storages".to_owned(),
                    ));
                }
            }
        }
    }
    Ok(result)
}

fn generated_author_public_id() -> String {
    GENERATED_AUTHOR_PUBLIC_ID.to_owned()
}

fn generated_annotation_author() -> tsk::AnnotationAuthorArchive {
    let public_id = generated_author_public_id();
    tsk::AnnotationAuthorArchive {
        name: Some(GENERATED_AUTHOR_NAME.to_owned()),
        color: Some(tsp::Color {
            model: tsp::color::ColorModel::Rgb as i32,
            r: Some(0.368_627_46),
            g: Some(0.568_627_5),
            b: Some(0.937_254_9),
            rgbspace: Some(tsp::color::RgbColorSpace::Srgb as i32),
            a: Some(1.0),
            ..Default::default()
        }),
        public_id: Some(public_id.clone()),
        is_public_author: Some(false),
        public_ids: vec![public_id],
    }
}

fn generated_local_annotation_author() -> tsk::AnnotationAuthorArchive {
    tsk::AnnotationAuthorArchive {
        name: Some(GENERATED_AUTHOR_NAME.to_owned()),
        color: generated_annotation_author().color,
        public_id: None,
        is_public_author: Some(false),
        public_ids: Vec::new(),
    }
}

fn generated_annotation_author_object(
    author_id: u64,
    author: &tsk::AnnotationAuthorArchive,
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        author_id,
        vec![RawMessage {
            type_: ANNOTATION_AUTHOR_MESSAGE_TYPE,
            data: author.encode_to_vec(),
        }],
    )?;
    object.archive_info.message_infos[0].field_infos = [4, 3]
        .into_iter()
        .map(|field_number| FieldInfo {
            path: FieldPath {
                path: vec![field_number],
            },
            unknown_field_rule: Some(UnknownFieldRule::IgnoreAndPreserve),
            ..Default::default()
        })
        .collect();
    Ok(object)
}

pub(crate) fn ensure_annotation_author(
    package: &mut IWorkPackage,
) -> Result<(Option<u64>, Option<String>, bool)> {
    ensure_generated_annotation_author(package, generated_annotation_author())
}

pub(crate) fn ensure_table_annotation_author(
    package: &mut IWorkPackage,
) -> Result<(Option<u64>, Option<String>, bool)> {
    ensure_generated_annotation_author(package, generated_local_annotation_author())
}

pub(crate) fn preferred_or_ensure_table_annotation_author(
    package: &mut IWorkPackage,
) -> Result<(Option<u64>, Option<String>, bool)> {
    let Some(location) = annotation_author_storage_location(package)? else {
        return Ok((None, None, false));
    };
    let author_ids = {
        let archive = package.archive(&location.archive_name)?;
        let object = archive.object(location.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "annotation-author storage object {} is missing",
                location.object_id
            ))
        })?;
        annotation_author_storage_ids(
            location.object_id,
            object.messages[location.message_index].data.as_slice(),
        )?
    };
    if author_ids.is_empty() {
        ensure_table_annotation_author(package)
    } else {
        preferred_annotation_author(package)
    }
}

fn ensure_generated_annotation_author(
    package: &mut IWorkPackage,
    generated: tsk::AnnotationAuthorArchive,
) -> Result<(Option<u64>, Option<String>, bool)> {
    let Some(location) = annotation_author_storage_location(package)? else {
        return Ok((None, None, false));
    };
    let author_ids = {
        let archive = package.archive(&location.archive_name)?;
        let object = archive.object(location.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "annotation-author storage object {} is missing",
                location.object_id
            ))
        })?;
        annotation_author_storage_ids(
            location.object_id,
            object.messages[location.message_index].data.as_slice(),
        )?
    };
    let mut seen = HashSet::new();
    let locations = object_locations(package)?;
    for author_id in &author_ids {
        if !seen.insert(*author_id) {
            return Err(Error::InvalidFormat(format!(
                "annotation-author storage duplicates object {}",
                author_id
            )));
        }
        if annotation_author_is_generated(
            package,
            &locations,
            *author_id,
            generated.public_id.is_some(),
        )? {
            return Ok((Some(*author_id), Some(location.archive_name), false));
        }
    }

    let author_id = next_object_identifier(package)?;
    let mut expected_authors = author_ids;
    expected_authors.try_reserve(1).map_err(|_| {
        comment_storage_allocation_error(
            "iWork annotation-author identifiers",
            expected_authors.len().saturating_add(1),
        )
    })?;
    expected_authors.push(author_id);
    package.update_archive(&location.archive_name, |archive| {
        {
            let storage_object = archive.object_mut(location.object_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "annotation-author storage object {} is missing",
                    location.object_id
                ))
            })?;
            let original = storage_object.messages[location.message_index]
                .data
                .as_slice();
            let data = append_repeated_length_delimited_field(
                original,
                1,
                &object_reference(author_id).encode_to_vec(),
            )?;
            let verified = annotation_author_storage_ids(location.object_id, data.as_slice())?;
            if verified != expected_authors {
                return Err(Error::InvalidFormat(
                    "annotation-author storage update failed validation".to_owned(),
                ));
            }
            storage_object.replace_message(
                location.message_index,
                RawMessage {
                    type_: ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
                    data,
                },
            )?;
        }
        Ok(archive.insert_object(generated_annotation_author_object(author_id, &generated)?)?)
    })?;
    Ok((Some(author_id), Some(location.archive_name), true))
}

/// Return the first author already registered by iWork.
fn preferred_annotation_author(
    package: &mut IWorkPackage,
) -> Result<(Option<u64>, Option<String>, bool)> {
    let Some(location) = annotation_author_storage_location(package)? else {
        return Ok((None, None, false));
    };
    let author_ids = {
        let archive = package.archive(&location.archive_name)?;
        let object = archive.object(location.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "annotation-author storage object {} is missing",
                location.object_id
            ))
        })?;
        annotation_author_storage_ids(
            location.object_id,
            object.messages[location.message_index].data.as_slice(),
        )?
    };
    if author_ids.is_empty() {
        return Err(Error::InvalidFormat(
            "annotation-author storage unexpectedly has no registered authors".to_owned(),
        ));
    }
    let locations = object_locations(package)?;
    let mut seen = HashSet::new();
    for author_id in &author_ids {
        if !seen.insert(*author_id) {
            return Err(Error::InvalidFormat(format!(
                "annotation-author storage duplicates object {}",
                author_id
            )));
        }
        validate_annotation_author(package, &locations, *author_id)?;
    }
    Ok((Some(author_ids[0]), Some(location.archive_name), false))
}

pub(crate) fn remove_generated_annotation_author_if_unused(
    package: &mut IWorkPackage,
    author_id: u64,
) -> Result<bool> {
    let locations = object_locations(package)?;
    if !annotation_author_is_generated(package, &locations, author_id, true)?
        && !annotation_author_is_generated(package, &locations, author_id, false)?
    {
        return Ok(false);
    }
    for name in package.iwa_entry_names() {
        for object in package.archive(name)?.objects {
            let object_id = object.archive_info.identifier.ok_or_else(|| {
                Error::Archive(format!("object in {name} has no archive identifier"))
            })?;
            for message in object
                .messages
                .iter()
                .filter(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            {
                let (comment, _) =
                    decode_comment_storage_payload(object_id, message.data.as_slice())?;
                if comment
                    .author()
                    .is_some_and(|reference| reference.identifier() == author_id)
                {
                    return Ok(false);
                }
            }
        }
    }

    let location = annotation_author_storage_location(package)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "generated annotation author {author_id} has no author storage"
        ))
    })?;
    if let Some(component_identifier) =
        component_identifier_for_entry(package, &location.archive_name)?
    {
        remove_component_external_references_to_object(package, component_identifier, author_id)?;
    }
    package.update_archive(&location.archive_name, |archive| {
        let storage_object = archive.object_mut(location.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "annotation-author storage object {} is missing",
                location.object_id
            ))
        })?;
        let original = storage_object.messages[location.message_index]
            .data
            .as_slice();
        let author_ids = annotation_author_storage_ids(location.object_id, original)?;
        if author_ids
            .iter()
            .filter(|identifier| **identifier == author_id)
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "generated annotation author {author_id} is not registered exactly once"
            )));
        }
        let data = remove_repeated_length_delimited_field_where(original, 1, |payload| {
            let reference = comment_storage_codec::decode_reference(
                payload,
                comment_storage_decode_options(payload),
            )
            .map_err(|error| strict_annotation_author_error(author_id, error))?;
            Ok(reference.identifier() == author_id)
        })?;
        if annotation_author_storage_ids(location.object_id, data.as_slice())?.contains(&author_id)
        {
            return Err(Error::InvalidFormat(
                "annotation-author storage removal failed validation".to_owned(),
            ));
        }
        storage_object.replace_message(
            location.message_index,
            RawMessage {
                type_: ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
                data,
            },
        )?;
        archive.remove_object(author_id).ok_or_else(|| {
            Error::InvalidFormat(format!("annotation author object {author_id} is missing"))
        })?;
        Ok(())
    })?;
    Ok(true)
}

pub(crate) fn current_apple_reference_date() -> Result<tsp::Date> {
    let unix_seconds = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(|error| {
            Error::ParseError(format!("system clock predates the Unix epoch: {error}"))
        })?
        .as_secs_f64();
    Ok(tsp::Date {
        seconds: unix_seconds - APPLE_EPOCH_UNIX_OFFSET_SECONDS,
    })
}

#[derive(Debug, Default)]
struct RemovedCommentGraph {
    object_ids: Vec<u64>,
    author_ids: HashSet<u64>,
}

fn remove_unreferenced_comment_graph(
    package: &mut IWorkPackage,
    application: Application,
    root: u64,
) -> Result<RemovedCommentGraph> {
    let mut candidate = package.clone();
    let removed = remove_unreferenced_comment_graph_in_place(&mut candidate, application, root)?;
    *package = candidate;
    Ok(removed)
}

fn remove_unreferenced_comment_graph_in_place(
    package: &mut IWorkPackage,
    application: Application,
    root: u64,
) -> Result<RemovedCommentGraph> {
    let mut pending = vec![root];
    let mut visited = HashSet::new();
    let mut removed = RemovedCommentGraph::default();
    while let Some(identifier) = pending.pop() {
        if !visited.insert(identifier)
            || comment_object_is_referenced(package, application, identifier)?
        {
            continue;
        }
        let locations = object_locations(package)?;
        let Some(archive_name) = locations.get(&identifier).cloned() else {
            continue;
        };
        let archive = package.archive(&archive_name)?;
        let Some(object) = archive.object(identifier) else {
            continue;
        };
        let mut replies = Vec::new();
        for message in object
            .messages
            .iter()
            .filter(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
        {
            let (comment, reply_ids) =
                decode_comment_storage_payload(identifier, message.data.as_slice())?;
            if let Some(author) = comment.author() {
                removed.author_ids.insert(author.identifier());
            }
            replies.extend(reply_ids);
        }
        let mut archive = package.archive(&archive_name)?;
        archive.remove_object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("comment storage object {identifier} is missing"))
        })?;
        if archive.objects.is_empty() {
            package.remove_entry(&archive_name).ok_or_else(|| {
                Error::InvalidFormat(format!("package entry {archive_name} is missing"))
            })?;
        } else {
            package.replace_archive(&archive_name, &archive)?;
        }
        removed.object_ids.push(identifier);
        pending.extend(replies);
    }
    Ok(removed)
}

fn comment_object_is_referenced(
    package: &IWorkPackage,
    application: Application,
    identifier: u64,
) -> Result<bool> {
    let mut referenced = drawable_locations(package, application)?
        .values()
        .any(|drawable| drawable.comment_storage_object_id == Some(identifier));
    let metadata_expected = package.contains_entry(PACKAGE_METADATA_ENTRY);
    let metadata_options = package_metadata_read_options(package);
    let archive_limits = package.limits().effective_archive_limits()?;
    let mut metadata_payloads = 0usize;
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        for object in &archive.objects {
            let mut archive_visitor = CommentArchiveReferenceVisitor {
                identifier,
                referenced: false,
            };
            object.inspect_references_with_policy_and_limits(
                &mut archive_visitor,
                ArchiveReferencePolicy::RejectUnknownMetadata,
                archive_limits,
            )?;
            referenced |= archive_visitor.referenced;
            for message in &object.messages {
                if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE {
                    if name != PACKAGE_METADATA_ENTRY {
                        return Err(Error::InvalidFormat(format!(
                            "PackageMetadata payload is stored in unexpected member {name}"
                        )));
                    }
                    metadata_payloads = metadata_payloads.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("PackageMetadata payload count overflow".to_owned())
                    })?;
                    if metadata_payloads != 1 {
                        return Err(Error::InvalidFormat(
                            "Package contains multiple PackageMetadata payloads".to_owned(),
                        ));
                    }
                    let mut visitor = CommentMetadataReferenceVisitor {
                        identifier,
                        referenced: false,
                        unknown_fields: false,
                    };
                    inspect_package_metadata_source(
                        message.data.as_slice(),
                        metadata_options,
                        &mut visitor,
                    )?;
                    if visitor.unknown_fields {
                        return Err(Error::InvalidFormat(
                            "PackageMetadata contains an unknown field that may own a comment object"
                                .to_owned(),
                        ));
                    }
                    referenced |= visitor.referenced;
                }
                if message.type_ == COMMENT_STORAGE_MESSAGE_TYPE {
                    let object_id = object.archive_info.identifier.ok_or_else(|| {
                        Error::Archive(format!("object in {name} has no archive identifier"))
                    })?;
                    let (_, reply_ids) =
                        decode_comment_storage_payload(object_id, message.data.as_slice())?;
                    referenced |= reply_ids.contains(&identifier);
                }
            }
        }
    }
    if metadata_expected && metadata_payloads != 1 {
        return Err(Error::InvalidFormat(
            "PackageMetadata payload is missing from Index/Metadata.iwa".to_owned(),
        ));
    }
    Ok(referenced)
}

struct CommentArchiveReferenceVisitor {
    identifier: u64,
    referenced: bool,
}

impl ArchiveReferenceVisitor for CommentArchiveReferenceVisitor {
    fn visit_reference(&mut self, occurrence: ArchiveReferenceOccurrence) -> CoreResult<()> {
        self.referenced |= occurrence.referenced_identifier == self.identifier;
        Ok(())
    }
}

struct CommentMetadataReferenceVisitor {
    identifier: u64,
    referenced: bool,
    unknown_fields: bool,
}

impl PackageMetadataVisitor for CommentMetadataReferenceVisitor {
    fn visit_unknown_field(&mut self) -> std::result::Result<(), RewriteError> {
        self.unknown_fields = true;
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: ObjectUuidDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        self.referenced |= binding.object_identifier() == self.identifier;
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: ExternalReferenceDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        self.referenced |= reference.object_identifier() == Some(self.identifier);
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: DataReferenceOwnerDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        self.referenced |= owner.object_identifier() == self.identifier;
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: litchi_iwa_protos::package_metadata_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> std::result::Result<(), RewriteError> {
        self.referenced |= identifier == self.identifier;
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        has_unknown_fields: bool,
    ) -> std::result::Result<(), RewriteError> {
        self.referenced |= object_identifier == self.identifier;
        self.unknown_fields |= has_unknown_fields;
        Ok(())
    }
}

/// Prove that a direct drawable edge is the only known owner of a comment
/// storage object before allowing an in-place text rewrite.
///
/// The selected edge is removed only from a private package clone, then the
/// existing package-wide reference census checks drawable payloads, comment
/// reply edges, archive metadata, and component external references.  A
/// negative result deliberately falls back to the caller's copy-on-write
/// path, preserving the legacy raw APIs while preventing a hidden owner from
/// observing the in-place text mutation.
fn prove_global_comment_ownership(
    package: &IWorkPackage,
    application: Application,
    location: &DrawableLocation,
    storage_id: u64,
) -> Result<bool> {
    let archive = package.archive(&location.archive_name)?;
    let object = archive.object(location.object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("drawable object {} is missing", location.object_id))
    })?;
    let message_info = object
        .archive_info
        .message_infos
        .get(location.message_index)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "drawable object {} lost payload metadata {}",
                location.object_id, location.message_index
            ))
        })?;
    let metadata_occurrences = message_info
        .object_references
        .iter()
        .filter(|reference| **reference == storage_id)
        .count()
        + message_info
            .field_infos
            .iter()
            .flat_map(|field| &field.object_references)
            .filter(|reference| **reference == storage_id)
            .count();
    if metadata_occurrences > 1 {
        return Err(Error::InvalidFormat(format!(
            "drawable object {} duplicates comment storage reference {storage_id}",
            location.object_id
        )));
    }

    let mut detached = package.clone();
    replace_drawable_comment_reference(
        &mut detached,
        application,
        location,
        Some(storage_id),
        None,
    )?;
    Ok(!comment_object_is_referenced(
        &detached,
        application,
        storage_id,
    )?)
}

/// Borrowed routing facts for the direct drawable comment edge.
///
/// The full generated drawable envelopes are intentionally not part of this
/// production projection.  We validate the selected nested path and the
/// `TSP.Reference` at its leaf, while the original payload remains the only
/// representation used for writes.  This keeps unrelated fields (including
/// producer extensions and unknown bytes) opaque and byte-preserving.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct DrawableCommentProjection {
    comment_identifier: Option<u64>,
    comment_wire_path: &'static [u32],
}

impl DrawableCommentProjection {
    fn decode(application: Application, message_type: u32, source: &[u8]) -> Result<Option<Self>> {
        let Some((comment_wire_path, chart_super_optional)) =
            drawable_comment_wire_path(application, message_type)
        else {
            return Ok(None);
        };
        let (comment_identifier, first_envelope_present) =
            decode_drawable_comment_identifier(message_type, source, comment_wire_path)?;
        if chart_super_optional && !first_envelope_present {
            return Ok(None);
        }
        Ok(Some(Self {
            comment_identifier,
            comment_wire_path,
        }))
    }
}

fn drawable_comment_wire_path(
    application: Application,
    message_type: u32,
) -> Option<(&'static [u32], bool)> {
    let route = match message_type {
        3002 => &[6][..],
        3004 | 3005 | 3006 | 3007 | 3008 | 5021 | 6000 => &[1, 6][..],
        3009 | 2011 | 6007 => &[1, 1, 6][..],
        2014 | 12 => &[1, 1, 1, 6][..],
        7 => match application {
            Application::Pages | Application::Keynote | Application::Numbers => &[1, 1, 1, 6],
            Application::Common => return None,
        },
        _ => return None,
    };
    if message_type == 12 && application != Application::Keynote
        || (message_type == 6000 || message_type == 6007) && application == Application::Numbers
    {
        return None;
    }
    Some((route, message_type == 5021))
}

fn drawable_projection_error(message_type: u32, error: impl std::fmt::Display) -> Error {
    Error::InvalidFormat(format!(
        "drawable payload {message_type} failed strict comment validation: {error}"
    ))
}

fn decode_drawable_comment_identifier(
    message_type: u32,
    source: &[u8],
    path: &[u32],
) -> Result<(Option<u64>, bool)> {
    let mut current = source;
    let mut first_envelope_present = false;
    for (depth, field_number) in path.iter().copied().enumerate() {
        let fields = parse_wire_fields(current)
            .map_err(|error| drawable_projection_error(message_type, error))?;
        let mut selected = None;
        for field in fields {
            field
                .validate_canonical_framing(current)
                .map_err(|error| drawable_projection_error(message_type, error))?;
            if field.number() != field_number {
                continue;
            }
            if selected.is_some() {
                return Err(drawable_projection_error(
                    message_type,
                    format!("duplicate selected field {field_number}"),
                ));
            }
            if field.wire_type() != 2 {
                return Err(drawable_projection_error(
                    message_type,
                    format!("selected field {field_number} is not length-delimited"),
                ));
            }
            selected = Some(
                field
                    .payload(current)
                    .map_err(|error| drawable_projection_error(message_type, error))?,
            );
        }
        let Some(selected) = selected else {
            return Ok((None, first_envelope_present));
        };
        if depth == 0 {
            first_envelope_present = true;
        }
        if depth + 1 == path.len() {
            let reference = comment_storage_codec::decode_reference(
                selected,
                comment_storage_decode_options(selected),
            )
            .map_err(|error| drawable_projection_error(message_type, error))?;
            return Ok((Some(reference.identifier()), first_envelope_present));
        }
        current = selected;
    }
    Err(drawable_projection_error(
        message_type,
        "empty drawable comment path",
    ))
}

#[cfg(test)]
enum DrawablePayload {
    Drawable(tsd::DrawableArchive),
    Shape(tsd::ShapeArchive),
    Image(tsd::ImageArchive),
    Mask(tsd::MaskArchive),
    Movie(tsd::MovieArchive),
    Group(tsd::GroupArchive),
    ConnectionLine(tsd::ConnectionLineArchive),
    ShapeInfo(tswp::ShapeInfoArchive),
    CommentInfo(tswp::CommentInfoArchive),
    PagesPlaceholder(tp::PlaceholderArchive),
    KeynotePlaceholder(kn::PlaceholderArchive),
    NumbersPlaceholder(tn::PlaceholderArchive),
    Chart(tsch::ChartDrawableArchive),
    Table(tst::TableInfoArchive),
    WpTable(tst::WpTableInfoArchive),
}

#[cfg(test)]
impl DrawablePayload {
    fn decode(application: Application, type_: u32, data: &[u8]) -> Result<Option<Self>> {
        let payload = match type_ {
            3002 => Self::Drawable(tsd::DrawableArchive::decode(data)?),
            3004 => Self::Shape(tsd::ShapeArchive::decode(data)?),
            3005 => Self::Image(tsd::ImageArchive::decode(data)?),
            3006 => Self::Mask(tsd::MaskArchive::decode(data)?),
            3007 => Self::Movie(tsd::MovieArchive::decode(data)?),
            3008 => Self::Group(tsd::GroupArchive::decode(data)?),
            3009 => Self::ConnectionLine(tsd::ConnectionLineArchive::decode(data)?),
            2011 => Self::ShapeInfo(tswp::ShapeInfoArchive::decode(data)?),
            2014 => Self::CommentInfo(tswp::CommentInfoArchive::decode(data)?),
            7 => match application {
                Application::Pages => Self::PagesPlaceholder(tp::PlaceholderArchive::decode(data)?),
                Application::Keynote => {
                    Self::KeynotePlaceholder(kn::PlaceholderArchive::decode(data)?)
                },
                Application::Numbers => {
                    Self::NumbersPlaceholder(tn::PlaceholderArchive::decode(data)?)
                },
                Application::Common => return Ok(None),
            },
            12 if application == Application::Keynote => {
                Self::KeynotePlaceholder(kn::PlaceholderArchive::decode(data)?)
            },
            5021 => Self::Chart(tsch::ChartDrawableArchive::decode(data)?),
            6000 if application != Application::Numbers => {
                Self::Table(tst::TableInfoArchive::decode(data)?)
            },
            6007 if application != Application::Numbers => {
                Self::WpTable(tst::WpTableInfoArchive::decode(data)?)
            },
            _ => return Ok(None),
        };
        if matches!(&payload, Self::Chart(chart) if chart.super_.is_none()) {
            return Ok(None);
        }
        Ok(Some(payload))
    }

    fn drawable(&self) -> &tsd::DrawableArchive {
        match self {
            Self::Drawable(value) => value,
            Self::Shape(value) => &value.super_,
            Self::Image(value) => &value.super_,
            Self::Mask(value) => &value.super_,
            Self::Movie(value) => &value.super_,
            Self::Group(value) => &value.super_,
            Self::ConnectionLine(value) => &value.super_.super_,
            Self::ShapeInfo(value) => &value.super_.super_,
            Self::CommentInfo(value) => &value.super_.super_.super_,
            Self::PagesPlaceholder(value) => &value.super_.super_.super_,
            Self::KeynotePlaceholder(value) => &value.super_.super_.super_,
            Self::NumbersPlaceholder(value) => &value.super_.super_.super_,
            Self::Chart(value) => value.super_.as_ref().expect("checked while decoding"),
            Self::Table(value) => &value.super_,
            Self::WpTable(value) => &value.super_.super_,
        }
    }

    #[cfg(test)]
    fn drawable_mut(&mut self) -> &mut tsd::DrawableArchive {
        match self {
            Self::Drawable(value) => value,
            Self::Shape(value) => &mut value.super_,
            Self::Image(value) => &mut value.super_,
            Self::Mask(value) => &mut value.super_,
            Self::Movie(value) => &mut value.super_,
            Self::Group(value) => &mut value.super_,
            Self::ConnectionLine(value) => &mut value.super_.super_,
            Self::ShapeInfo(value) => &mut value.super_.super_,
            Self::CommentInfo(value) => &mut value.super_.super_.super_,
            Self::PagesPlaceholder(value) => &mut value.super_.super_.super_,
            Self::KeynotePlaceholder(value) => &mut value.super_.super_.super_,
            Self::NumbersPlaceholder(value) => &mut value.super_.super_.super_,
            Self::Chart(value) => value.super_.as_mut().expect("checked while decoding"),
            Self::Table(value) => &mut value.super_,
            Self::WpTable(value) => &mut value.super_.super_,
        }
    }

    fn comment_identifier(&self) -> Option<u64> {
        self.drawable()
            .comment
            .as_ref()
            .map(|value| value.identifier)
    }

    #[cfg(test)]
    fn set_comment_identifier(&mut self, identifier: Option<u64>) {
        self.drawable_mut().comment = identifier.map(|identifier| tsp::Reference {
            identifier,
            ..Default::default()
        });
    }

    #[cfg(test)]
    fn encode_to_vec(&self) -> Vec<u8> {
        match self {
            Self::Drawable(value) => value.encode_to_vec(),
            Self::Shape(value) => value.encode_to_vec(),
            Self::Image(value) => value.encode_to_vec(),
            Self::Mask(value) => value.encode_to_vec(),
            Self::Movie(value) => value.encode_to_vec(),
            Self::Group(value) => value.encode_to_vec(),
            Self::ConnectionLine(value) => value.encode_to_vec(),
            Self::ShapeInfo(value) => value.encode_to_vec(),
            Self::CommentInfo(value) => value.encode_to_vec(),
            Self::PagesPlaceholder(value) => value.encode_to_vec(),
            Self::KeynotePlaceholder(value) => value.encode_to_vec(),
            Self::NumbersPlaceholder(value) => value.encode_to_vec(),
            Self::Chart(value) => value.encode_to_vec(),
            Self::Table(value) => value.encode_to_vec(),
            Self::WpTable(value) => value.encode_to_vec(),
        }
    }
}

#[cfg(test)]
#[allow(deprecated)]
mod tests;
