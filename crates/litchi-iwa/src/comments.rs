//! Cross-application comments attached directly to drawable objects.

use std::collections::{HashMap, HashSet};
use std::fmt;
use std::num::NonZeroU64;
use std::path::Path;
use std::time::{SystemTime, UNIX_EPOCH};

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::detect::detect_application_from_document;
use crate::package_metadata::{
    add_component_external_reference, advance_package_save_token_for_components,
    component_identifier_for_entry, next_object_identifier, release_package_identifier_suffix,
    remove_component_external_references_to_object, set_package_last_object_identifier,
};
use crate::protobuf::{kn, tn, tp, tsch, tsd, tsk, tsp, tst, tswp};
use crate::registry::Application;
#[cfg(test)]
use crate::wire::parse_wire_fields;
use crate::wire::{
    append_repeated_length_delimited_field, patch_length_delimited_field,
    patch_nested_length_delimited_field, remove_repeated_length_delimited_field_where,
    transform_repeated_length_delimited_fields,
};
use crate::{Error, IWorkPackage, Result};

const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3056;
const ANNOTATION_AUTHOR_MESSAGE_TYPE: u32 = 212;
const ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE: u32 = 213;
const APPLE_EPOCH_UNIX_OFFSET_SECONDS: f64 = 978_307_200.0;
const GENERATED_AUTHOR_NAME: &str = "litchi-iwa";

/// A validated native object identifier for a drawable.
///
/// iWork uses zero as the absence value in protobuf references. Keeping that
/// sentinel out of the semantic API prevents callers from accidentally
/// addressing an invalid drawable while leaving the wire representation as a
/// plain `u64` at the archive boundary.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct DrawableObjectId(NonZeroU64);

impl DrawableObjectId {
    /// Construct a drawable identifier, returning `None` for the protobuf
    /// null sentinel.
    pub const fn new(raw: u64) -> Option<Self> {
        match NonZeroU64::new(raw) {
            Some(value) => Some(Self(value)),
            None => None,
        }
    }

    /// Construct an identifier from a native iWork object identifier.
    pub fn from_object_id(raw: u64) -> Result<Self> {
        Self::new(raw).ok_or_else(|| {
            Error::ParseError("drawable object identifier must be non-zero".to_owned())
        })
    }

    /// Return the native identifier used at the archive boundary.
    pub const fn object_id(self) -> u64 {
        self.0.get()
    }
}

impl TryFrom<u64> for DrawableObjectId {
    type Error = Error;

    fn try_from(raw: u64) -> Result<Self> {
        Self::from_object_id(raw)
    }
}

impl From<DrawableObjectId> for u64 {
    fn from(value: DrawableObjectId) -> Self {
        value.object_id()
    }
}

impl fmt::Display for DrawableObjectId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.object_id().fmt(formatter)
    }
}

/// A validated native object identifier for a comment-storage archive.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct CommentStorageId(NonZeroU64);

impl CommentStorageId {
    /// Construct a comment-storage identifier, returning `None` for the
    /// protobuf null sentinel.
    pub const fn new(raw: u64) -> Option<Self> {
        match NonZeroU64::new(raw) {
            Some(value) => Some(Self(value)),
            None => None,
        }
    }

    /// Construct an identifier from a native iWork object identifier.
    pub fn from_object_id(raw: u64) -> Result<Self> {
        Self::new(raw).ok_or_else(|| {
            Error::ParseError("comment-storage object identifier must be non-zero".to_owned())
        })
    }

    /// Return the native identifier used at the archive boundary.
    pub const fn object_id(self) -> u64 {
        self.0.get()
    }
}

impl TryFrom<u64> for CommentStorageId {
    type Error = Error;

    fn try_from(raw: u64) -> Result<Self> {
        Self::from_object_id(raw)
    }
}

impl From<CommentStorageId> for u64 {
    fn from(value: CommentStorageId) -> Self {
        value.object_id()
    }
}

impl fmt::Display for CommentStorageId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.object_id().fmt(formatter)
    }
}

/// Stable UUID stored on an iWork comment archive.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct IWorkCommentUuid {
    pub lower: u64,
    pub upper: u64,
}

/// Application-independent representation of a TSD comment.
#[derive(Debug, Clone, PartialEq)]
pub struct IWorkComment {
    pub text: String,
    /// Seconds since Apple's 2001-01-01 reference date.
    pub creation_date_seconds: Option<f64>,
    pub author_object_id: Option<u64>,
    pub reply_object_ids: Vec<u64>,
    pub storage_uuid: Option<IWorkCommentUuid>,
}

/// A drawable object and its direct comment attachment, if present.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IWorkDrawableInfo {
    pub object_id: DrawableObjectId,
    pub message_type: u32,
    pub comment_storage_object_id: Option<CommentStorageId>,
}

/// A resolved direct drawable comment.
#[derive(Debug, Clone, PartialEq)]
pub struct DrawableCommentInfo {
    pub drawable_object_id: DrawableObjectId,
    pub storage_object_id: CommentStorageId,
    pub comment: IWorkComment,
}

/// A resolved reply in a drawable's direct comment thread.
#[derive(Debug, Clone, PartialEq)]
pub struct DrawableCommentReplyInfo {
    pub drawable_object_id: DrawableObjectId,
    pub root_storage_object_id: CommentStorageId,
    pub storage_object_id: CommentStorageId,
    pub comment: IWorkComment,
}

/// Address and storage identity of a comment attached to an iWork table cell.
#[derive(Debug, Clone, PartialEq)]
pub struct IWorkTableCellCommentInfo {
    pub table_id: u64,
    pub row: usize,
    pub column: usize,
    pub list_identifier: u32,
    pub storage_object_id: u64,
    pub comment: IWorkComment,
}

/// A resolved direct reply in an iWork table-cell comment thread.
#[derive(Debug, Clone, PartialEq)]
pub struct IWorkTableCellCommentReplyInfo {
    pub table_id: u64,
    pub row: usize,
    pub column: usize,
    pub root_storage_object_id: u64,
    pub storage_object_id: u64,
    pub comment: IWorkComment,
}

/// Transactional direct-comment editor shared by Pages, Numbers, and Keynote.
///
/// It edits the `TSD.DrawableArchive.comment` reference nested in each known
/// drawable payload and stores comment bodies in `TSD.CommentStorageArchive`
/// objects. Table-cell comments use a separate table-list indirection and are
/// available through each application's semantic editor.
#[derive(Debug, Clone)]
pub struct IWorkDrawableCommentEditor {
    package: IWorkPackage,
    application: Application,
}

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

    pub fn drawables(&self) -> Result<Vec<IWorkDrawableInfo>> {
        let mut drawables = drawable_locations(&self.package, self.application)?
            .into_values()
            .map(|location| {
                Ok(IWorkDrawableInfo {
                    object_id: DrawableObjectId::from_object_id(location.object_id)?,
                    message_type: location.message_type,
                    comment_storage_object_id: location
                        .comment_storage_object_id
                        .map(CommentStorageId::from_object_id)
                        .transpose()?,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        drawables.sort_by_key(|drawable| drawable.object_id.object_id());
        Ok(drawables)
    }

    pub fn comment<D>(&self, drawable_object_id: D) -> Result<Option<DrawableCommentInfo>>
    where
        D: TryInto<DrawableObjectId>,
        D::Error: fmt::Debug,
    {
        let drawable_object_id = normalize_drawable_object_id(drawable_object_id)?;
        drawable_comment_in_package(&self.package, self.application, drawable_object_id)
    }

    /// Resolves the direct replies to a drawable comment in stored order.
    pub fn replies<D>(&self, drawable_object_id: D) -> Result<Vec<DrawableCommentReplyInfo>>
    where
        D: TryInto<DrawableObjectId>,
        D::Error: fmt::Debug,
    {
        let drawable_object_id = normalize_drawable_object_id(drawable_object_id)?;
        drawable_comment_replies_in_package(&self.package, self.application, drawable_object_id)
    }

    pub fn set_comment<D>(&mut self, drawable_object_id: D, text: impl Into<String>) -> Result<()>
    where
        D: TryInto<DrawableObjectId>,
        D::Error: fmt::Debug,
    {
        let drawable_object_id = normalize_drawable_object_id(drawable_object_id)?;
        let mut staged = self.package.clone();
        set_drawable_comment_in_package(
            &mut staged,
            self.application,
            drawable_object_id,
            text.into(),
        )?;
        validate_package_round_trip(&staged)?;
        self.package = staged;
        Ok(())
    }

    pub fn clear_comment<D>(&mut self, drawable_object_id: D) -> Result<()>
    where
        D: TryInto<DrawableObjectId>,
        D::Error: fmt::Debug,
    {
        let drawable_object_id = normalize_drawable_object_id(drawable_object_id)?;
        let mut staged = self.package.clone();
        clear_drawable_comment_in_package(&mut staged, self.application, drawable_object_id)?;
        validate_package_round_trip(&staged)?;
        self.package = staged;
        Ok(())
    }

    /// Adds a reply and returns its new comment-storage object identifier.
    ///
    /// The root storage is copy-on-written, matching native iWork saves and
    /// isolating a drawable when multiple drawables share one thread.
    pub fn add_reply<D>(
        &mut self,
        drawable_object_id: D,
        text: impl Into<String>,
    ) -> Result<CommentStorageId>
    where
        D: TryInto<DrawableObjectId>,
        D::Error: fmt::Debug,
    {
        self.add_reply_id(drawable_object_id, text)
    }

    /// Adds a reply and returns its validated comment-storage identifier.
    pub fn add_reply_id<D>(
        &mut self,
        drawable_object_id: D,
        text: impl Into<String>,
    ) -> Result<CommentStorageId>
    where
        D: TryInto<DrawableObjectId>,
        D::Error: fmt::Debug,
    {
        let drawable_object_id = normalize_drawable_object_id(drawable_object_id)?;
        let mut staged = self.package.clone();
        let reply_id = add_drawable_comment_reply_in_package(
            &mut staged,
            self.application,
            drawable_object_id,
            text.into(),
        )?;
        validate_package_round_trip(&staged)?;
        let reply_id = CommentStorageId::from_object_id(reply_id)?;
        self.package = staged;
        Ok(reply_id)
    }

    /// Updates one direct reply and returns its current storage identifier.
    ///
    /// A changed reply and its root are copy-on-written. The returned value can
    /// therefore differ from `reply_storage_object_id`.
    pub fn set_reply<D, S>(
        &mut self,
        drawable_object_id: D,
        reply_storage_object_id: S,
        text: impl Into<String>,
    ) -> Result<CommentStorageId>
    where
        D: TryInto<DrawableObjectId>,
        D::Error: fmt::Debug,
        S: TryInto<CommentStorageId>,
        S::Error: fmt::Debug,
    {
        self.set_reply_id(drawable_object_id, reply_storage_object_id, text)
    }

    /// Updates one direct reply and returns its validated comment-storage
    /// identifier.
    pub fn set_reply_id<D, S>(
        &mut self,
        drawable_object_id: D,
        reply_storage_object_id: S,
        text: impl Into<String>,
    ) -> Result<CommentStorageId>
    where
        D: TryInto<DrawableObjectId>,
        D::Error: fmt::Debug,
        S: TryInto<CommentStorageId>,
        S::Error: fmt::Debug,
    {
        let drawable_object_id = normalize_drawable_object_id(drawable_object_id)?;
        let reply_storage_object_id = normalize_comment_storage_id(reply_storage_object_id)?;
        let mut staged = self.package.clone();
        let reply_id = set_drawable_comment_reply_in_package(
            &mut staged,
            self.application,
            drawable_object_id,
            reply_storage_object_id,
            text.into(),
        )?;
        validate_package_round_trip(&staged)?;
        let reply_id = CommentStorageId::from_object_id(reply_id)?;
        self.package = staged;
        Ok(reply_id)
    }

    /// Removes one direct reply from a drawable's comment thread.
    pub fn remove_reply<D, S>(
        &mut self,
        drawable_object_id: D,
        reply_storage_object_id: S,
    ) -> Result<()>
    where
        D: TryInto<DrawableObjectId>,
        D::Error: fmt::Debug,
        S: TryInto<CommentStorageId>,
        S::Error: fmt::Debug,
    {
        let drawable_object_id = normalize_drawable_object_id(drawable_object_id)?;
        let reply_storage_object_id = normalize_comment_storage_id(reply_storage_object_id)?;
        let mut staged = self.package.clone();
        remove_drawable_comment_reply_in_package(
            &mut staged,
            self.application,
            drawable_object_id,
            reply_storage_object_id,
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

fn normalize_drawable_object_id<D>(value: D) -> Result<u64>
where
    D: TryInto<DrawableObjectId>,
    D::Error: fmt::Debug,
{
    value
        .try_into()
        .map(DrawableObjectId::object_id)
        .map_err(|error| {
            Error::ParseError(format!("invalid drawable object identifier: {error:?}"))
        })
}

fn normalize_comment_storage_id<S>(value: S) -> Result<u64>
where
    S: TryInto<CommentStorageId>,
    S::Error: fmt::Debug,
{
    value
        .try_into()
        .map(CommentStorageId::object_id)
        .map_err(|error| {
            Error::ParseError(format!(
                "invalid comment-storage object identifier: {error:?}"
            ))
        })
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
                let Some(application) = detect_application_from_document(&message.data) else {
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
                    let Some(payload) = DrawablePayload::decode(
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
                        comment_storage_object_id: payload.comment_identifier(),
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
) -> Result<Option<DrawableCommentInfo>> {
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
    let drawable_object_id = DrawableObjectId::from_object_id(drawable_object_id)?;
    let storage_object_id = CommentStorageId::from_object_id(storage_object_id)?;
    Ok(Some(DrawableCommentInfo {
        drawable_object_id,
        storage_object_id,
        comment,
    }))
}

fn drawable_comment_replies_in_package(
    package: &IWorkPackage,
    application: Application,
    drawable_object_id: u64,
) -> Result<Vec<DrawableCommentReplyInfo>> {
    let Some(root) = drawable_comment_in_package(package, application, drawable_object_id)? else {
        return Ok(Vec::new());
    };
    let drawable_object_id = DrawableObjectId::from_object_id(drawable_object_id)?;
    let root_storage_id = root.storage_object_id.object_id();
    let locations = object_locations(package)?;
    let mut seen = HashSet::new();
    let mut replies = Vec::with_capacity(root.comment.reply_object_ids.len());
    for reply_id in root.comment.reply_object_ids {
        if reply_id == root_storage_id || !seen.insert(reply_id) {
            return Err(Error::InvalidFormat(format!(
                "comment storage {} contains a duplicate or cyclic reply reference to {reply_id}",
                root_storage_id
            )));
        }
        let storage_object_id = CommentStorageId::from_object_id(reply_id)?;
        replies.push(DrawableCommentReplyInfo {
            drawable_object_id,
            root_storage_object_id: root.storage_object_id,
            storage_object_id,
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
        let storage_entry = locations.get(&storage_id).cloned().ok_or_else(|| {
            Error::InvalidFormat(format!("comment storage object {storage_id} is missing"))
        })?;
        let direct_users = drawables
            .values()
            .filter(|candidate| candidate.comment_storage_object_id == Some(storage_id))
            .count();
        if direct_users == 1 {
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
        archive.insert_object(object)
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
    root: &IWorkComment,
    root_storage_id: u64,
    reply_storage_id: u64,
) -> Result<()> {
    if root_storage_id == reply_storage_id {
        return Err(Error::InvalidFormat(format!(
            "comment storage {root_storage_id} references itself as a reply"
        )));
    }
    match root
        .reply_object_ids
        .iter()
        .filter(|identifier| **identifier == reply_storage_id)
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
    root: &IWorkComment,
) -> Result<()> {
    let mut seen = HashSet::new();
    for reply_id in &root.reply_object_ids {
        if *reply_id == root_storage_id || !seen.insert(*reply_id) {
            return Err(Error::InvalidFormat(format!(
                "comment storage {root_storage_id} contains a duplicate or cyclic reply reference to {reply_id}"
            )));
        }
        read_comment_storage(package, locations, *reply_id)?;
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
        let payload = DrawablePayload::decode(application, message.type_, message.data.as_slice())?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "object {} payload {} is no longer a drawable",
                    location.object_id, location.message_index
                ))
            })?;
        if payload.comment_identifier() != old {
            return Err(Error::InvalidFormat(format!(
                "drawable object {} comment changed during mutation",
                location.object_id
            )));
        }
        let message_type = message.type_;
        let replacement = new.map(|identifier| {
            tsp::Reference {
                identifier,
                ..Default::default()
            }
            .encode_to_vec()
        });
        let data = patch_nested_length_delimited_field(
            message.data.as_slice(),
            payload.comment_wire_path(),
            old.is_some(),
            replacement.as_deref(),
        )?;
        let verified = DrawablePayload::decode(application, message_type, data.as_slice())?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "object {} stopped decoding as a drawable after comment patch",
                    location.object_id
                ))
            })?;
        if verified.comment_identifier() != new {
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

fn read_comment_storage(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    storage_id: u64,
) -> Result<IWorkComment> {
    let archive_name = locations.get(&storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!("comment storage object {storage_id} is missing"))
    })?;
    package.with_parsed_archive(archive_name, |archive| {
        let object = archive.object(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("comment storage object {storage_id} is missing"))
        })?;
        let messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if messages.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "object {storage_id} must contain exactly one TSD comment-storage payload"
            )));
        }
        let comment = tsd::CommentStorageArchive::decode(messages[0].data.as_slice())?;
        Ok(IWorkComment {
            text: comment.text.unwrap_or_default(),
            creation_date_seconds: comment.creation_date.map(|date| date.seconds),
            author_object_id: comment.author.map(|author| author.identifier),
            reply_object_ids: comment
                .replies
                .into_iter()
                .map(|reply| reply.identifier)
                .collect(),
            storage_uuid: comment.storage_uuid.map(|uuid| IWorkCommentUuid {
                lower: uuid.lower,
                upper: uuid.upper,
            }),
        })
    })
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
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "object {storage_id} must contain exactly one TSD comment-storage payload"
            )));
        }
        let index = indexes[0];
        let comment = tsd::CommentStorageArchive::decode(object.messages[index].data.as_slice())?;
        let data = patch_length_delimited_field(
            object.messages[index].data.as_slice(),
            1,
            comment.text.is_some(),
            Some(text.as_bytes()),
        )?;
        let verified = tsd::CommentStorageArchive::decode(data.as_slice())?;
        if verified.text.as_deref() != Some(text.as_str()) {
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
    let comment = tsd::CommentStorageArchive::decode(source.messages[0].data.as_slice())?;
    let data = patch_length_delimited_field(
        source.messages[0].data.as_slice(),
        1,
        comment.text.is_some(),
        Some(text.as_bytes()),
    )?;
    let uuid = storage_uuid.encode_to_vec();
    let data = patch_length_delimited_field(
        data.as_slice(),
        5,
        comment.storage_uuid.is_some(),
        Some(uuid.as_slice()),
    )?;
    let verified = tsd::CommentStorageArchive::decode(data.as_slice())?;
    if verified.text.as_deref() != Some(text.as_str())
        || verified.storage_uuid != Some(storage_uuid)
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
    package.update_archive(archive_name, |archive| archive.insert_object(clone))
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
    tsd::CommentStorageArchive::decode(source.messages[0].data.as_slice())?;
    let mut clone = ArchiveObject::new(new_storage_id, source.messages.clone())?;
    clone.archive_info.should_merge = source.archive_info.should_merge;
    clone.archive_info.message_infos = source.archive_info.message_infos.clone();
    package.update_archive(&archive_name, |archive| archive.insert_object(clone))?;
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
        archive.insert_object(object)
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
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "object {root_storage_id} must contain exactly one TSD comment-storage payload"
            )));
        }
        let index = indexes[0];
        let original = object.messages[index].data.as_slice();
        let before = tsd::CommentStorageArchive::decode(original)?;
        let old_count = old_reply_id.map_or(0, |identifier| {
            before
                .replies
                .iter()
                .filter(|reference| reference.identifier == identifier)
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
            && before
                .replies
                .iter()
                .any(|reference| reference.identifier == identifier)
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
                transform_repeated_length_delimited_fields(original, 4, |payload| {
                    let reference = tsp::Reference::decode(payload)?;
                    if reference.identifier == old {
                        Ok(object_reference(new).encode_to_vec())
                    } else {
                        Ok(payload.to_vec())
                    }
                })?
            },
            (Some(identifier), None) => {
                remove_repeated_length_delimited_field_where(original, 4, |payload| {
                    Ok(tsp::Reference::decode(payload)?.identifier == identifier)
                })?
            },
            (None, None) => return Ok(()),
        };
        let verified = tsd::CommentStorageArchive::decode(data.as_slice())?;
        let expected = before
            .replies
            .iter()
            .filter_map(|reference| {
                if old_reply_id == Some(reference.identifier) {
                    new_reply_id
                } else {
                    Some(reference.identifier)
                }
            })
            .chain((old_reply_id.is_none()).then_some(new_reply_id).flatten())
            .collect::<Vec<_>>();
        if verified
            .replies
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>()
            != expected
        {
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

#[derive(Debug, Clone)]
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
            for message in object
                .messages
                .iter()
                .filter(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            {
                if let Some(uuid) =
                    tsd::CommentStorageArchive::decode(message.data.as_slice())?.storage_uuid
                {
                    existing.insert((uuid.lower, uuid.upper));
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
                tsk::AnnotationAuthorStorageArchive::decode(message.data.as_slice())?;
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

fn annotation_author(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    author_id: u64,
) -> Result<tsk::AnnotationAuthorArchive> {
    let archive_name = locations.get(&author_id).ok_or_else(|| {
        Error::InvalidFormat(format!("annotation author object {author_id} is missing"))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(author_id).ok_or_else(|| {
        Error::InvalidFormat(format!("annotation author object {author_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == ANNOTATION_AUTHOR_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "object {author_id} must contain exactly one annotation-author payload"
        )));
    }
    Ok(tsk::AnnotationAuthorArchive::decode(
        messages[0].data.as_slice(),
    )?)
}

fn generated_author_public_id() -> String {
    "4C495443-4849-4957-8100-000000000001:058e44481db1c6fdeeac88af010136d7f8949f54bde61ef9af3e078562b968b6".to_owned()
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
        .map(|field_number| tsp::FieldInfo {
            path: tsp::FieldPath {
                path: vec![field_number],
            },
            unknown_field_rule: Some(tsp::field_info::UnknownFieldRule::IgnoreAndPreserve as i32),
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
    let storage = {
        let archive = package.archive(&location.archive_name)?;
        let object = archive.object(location.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "annotation-author storage object {} is missing",
                location.object_id
            ))
        })?;
        tsk::AnnotationAuthorStorageArchive::decode(
            object.messages[location.message_index].data.as_slice(),
        )?
    };
    if storage.annotation_author.is_empty() {
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
    let storage = {
        let archive = package.archive(&location.archive_name)?;
        let object = archive.object(location.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "annotation-author storage object {} is missing",
                location.object_id
            ))
        })?;
        tsk::AnnotationAuthorStorageArchive::decode(
            object.messages[location.message_index].data.as_slice(),
        )?
    };
    let mut seen = HashSet::new();
    let locations = object_locations(package)?;
    for author in &storage.annotation_author {
        if !seen.insert(author.identifier) {
            return Err(Error::InvalidFormat(format!(
                "annotation-author storage duplicates object {}",
                author.identifier
            )));
        }
        if annotation_author(package, &locations, author.identifier)? == generated {
            return Ok((Some(author.identifier), Some(location.archive_name), false));
        }
    }

    let author_id = next_object_identifier(package)?;
    let mut expected_authors = storage
        .annotation_author
        .iter()
        .map(|reference| reference.identifier)
        .collect::<Vec<_>>();
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
            let verified = tsk::AnnotationAuthorStorageArchive::decode(data.as_slice())?;
            if verified
                .annotation_author
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>()
                != expected_authors
            {
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
        archive.insert_object(generated_annotation_author_object(author_id, &generated)?)
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
    let storage = {
        let archive = package.archive(&location.archive_name)?;
        let object = archive.object(location.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "annotation-author storage object {} is missing",
                location.object_id
            ))
        })?;
        tsk::AnnotationAuthorStorageArchive::decode(
            object.messages[location.message_index].data.as_slice(),
        )?
    };
    if storage.annotation_author.is_empty() {
        return Err(Error::InvalidFormat(
            "annotation-author storage unexpectedly has no registered authors".to_owned(),
        ));
    }
    let locations = object_locations(package)?;
    let mut seen = HashSet::new();
    for author in &storage.annotation_author {
        if !seen.insert(author.identifier) {
            return Err(Error::InvalidFormat(format!(
                "annotation-author storage duplicates object {}",
                author.identifier
            )));
        }
        annotation_author(package, &locations, author.identifier)?;
    }
    Ok((
        Some(storage.annotation_author[0].identifier),
        Some(location.archive_name),
        false,
    ))
}

pub(crate) fn remove_generated_annotation_author_if_unused(
    package: &mut IWorkPackage,
    author_id: u64,
) -> Result<bool> {
    let locations = object_locations(package)?;
    let author = annotation_author(package, &locations, author_id)?;
    if author != generated_annotation_author() && author != generated_local_annotation_author() {
        return Ok(false);
    }
    for name in package.iwa_entry_names() {
        for object in package.archive(name)?.objects {
            for message in object
                .messages
                .iter()
                .filter(|message| message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
            {
                if tsd::CommentStorageArchive::decode(message.data.as_slice())?
                    .author
                    .is_some_and(|reference| reference.identifier == author_id)
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
        let storage = tsk::AnnotationAuthorStorageArchive::decode(original)?;
        if storage
            .annotation_author
            .iter()
            .filter(|reference| reference.identifier == author_id)
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "generated annotation author {author_id} is not registered exactly once"
            )));
        }
        let data = remove_repeated_length_delimited_field_where(original, 1, |payload| {
            Ok(tsp::Reference::decode(payload)?.identifier == author_id)
        })?;
        if tsk::AnnotationAuthorStorageArchive::decode(data.as_slice())?
            .annotation_author
            .iter()
            .any(|reference| reference.identifier == author_id)
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
            let comment = tsd::CommentStorageArchive::decode(message.data.as_slice())?;
            if let Some(author) = comment.author {
                removed.author_ids.insert(author.identifier);
            }
            replies.extend(
                comment
                    .replies
                    .into_iter()
                    .map(|reference| reference.identifier),
            );
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
    if drawable_locations(package, application)?
        .values()
        .any(|drawable| drawable.comment_storage_object_id == Some(identifier))
    {
        return Ok(true);
    }
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ == COMMENT_STORAGE_MESSAGE_TYPE {
                    let comment = tsd::CommentStorageArchive::decode(message.data.as_slice())?;
                    if comment
                        .replies
                        .iter()
                        .any(|reply| reply.identifier == identifier)
                    {
                        return Ok(true);
                    }
                }
            }
            if object.archive_info.message_infos.iter().any(|info| {
                info.object_references.contains(&identifier)
                    || info
                        .field_infos
                        .iter()
                        .any(|field| field.object_references.contains(&identifier))
            }) {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

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

    fn comment_wire_path(&self) -> &'static [u32] {
        match self {
            Self::Drawable(_) => &[6],
            Self::Shape(_)
            | Self::Image(_)
            | Self::Mask(_)
            | Self::Movie(_)
            | Self::Group(_)
            | Self::Chart(_)
            | Self::Table(_) => &[1, 6],
            Self::ConnectionLine(_) | Self::ShapeInfo(_) | Self::WpTable(_) => &[1, 1, 6],
            Self::CommentInfo(_)
            | Self::PagesPlaceholder(_)
            | Self::KeynotePlaceholder(_)
            | Self::NumbersPlaceholder(_) => &[1, 1, 1, 6],
        }
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
mod tests;
