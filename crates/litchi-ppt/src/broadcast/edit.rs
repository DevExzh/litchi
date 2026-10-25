//! Package-level broadcast edits (record 0357).
//!
//! [`Package::edit_broadcast`] stages edits of one `BroadcastDocInfo9Container`
//! in the live document's `PowerPoint` 9 programmable tag, [`Edit::commit`]
//! publishes a new [`Package`] through one append-only user edit that
//! rewrites only the live document, and [`Package::apply_broadcast`] replays
//! a [`Patch`] onto the exact package it was made from. Dropping an [`Edit`]
//! discards it. Broadcast metadata stays inert: nothing is contacted,
//! opened or started.

use litchi_core::Position;

use super::codec::BROADCAST_CONTAINER_RECORD_TYPE;
use super::transaction::{Snapshot, Transaction as RecordEdit};
use super::{Broadcast, BroadcastProperties, Commit, Patch, Result, UnknownRecord};
use crate::Package;
use crate::consts::RecordType;
use crate::record_wire as wire;
use crate::scoped::{self, Base, Failure, Lazy};
use crate::slide_order::{self, records::Owner};

const OPERATION: &str = "edit_broadcast";

/// A staged edit of one broadcast description over one immutable
/// [`Package`].
pub struct Edit {
    base: Base,
    position: usize,
    state: Lazy<State>,
}

struct State {
    native: slide_order::Snapshot,
    document: Vec<u8>,
    inner: RecordEdit,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("position", &self.position)
            .field("initialized", &self.state.get().is_some())
            .finish_non_exhaustive()
    }
}

impl Edit {
    fn new(package: &Package, position: Position) -> Result<Self> {
        let base = Base::capture(package)?;
        Ok(Self {
            base,
            position: position.get(),
            state: Lazy::new(),
        })
    }

    fn initialize_state(&self) -> std::result::Result<State, Failure> {
        let native = self.base.native()?;
        let document = self.base.owner_record(&native, Owner::Document)?;
        let located = Located::find(&document)?;
        let source = located.container(self.position)?;
        let limits = self.base.open_options().source_limits_value();
        let inner = Snapshot::parse(source, &limits)
            .map_err(scoped::publication(OPERATION))?
            .edit();
        Ok(State {
            native,
            document,
            inner,
        })
    }

    fn state(&self) -> std::result::Result<&State, Failure> {
        self.state.get_or_try_init(|| self.initialize_state())
    }

    fn state_mut(&mut self) -> std::result::Result<&mut State, Failure> {
        let base = self.base.clone();
        let position = self.position;
        self.state.get_or_try_init_mut(|| {
            let native = base.native()?;
            let document = base.owner_record(&native, Owner::Document)?;
            let located = Located::find(&document)?;
            let source = located.container(position)?;
            let limits = base.open_options().source_limits_value();
            let inner = Snapshot::parse(source, &limits)
                .map_err(scoped::publication(OPERATION))?
                .edit();
            Ok(State {
                native,
                document,
                inner,
            })
        })
    }

    /// Current typed broadcast projection.
    #[must_use = "handle the fallible broadcast projection result"]
    pub fn broadcast(&self) -> Result<&Broadcast> {
        Ok(self.state()?.inner.broadcast())
    }

    /// Opaque source records retained by every candidate publication.
    #[must_use = "handle the fallible broadcast records result"]
    pub fn unknown_records(&self) -> Result<&[UnknownRecord]> {
        Ok(self.state()?.inner.unknown_records())
    }

    /// Replace all typed broadcast fields atomically.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn replace(&mut self, replacement: Broadcast) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .replace(replacement)
            .map_err(scoped::staging("broadcast.replace"))?)
    }

    /// Set one optional inert title.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_title(&mut self, value: Option<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_title(value)
            .map_err(scoped::staging("broadcast.set_title"))?)
    }

    /// Set one optional inert description.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_description(&mut self, value: Option<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_description(value)
            .map_err(scoped::staging("broadcast.set_description"))?)
    }

    /// Set one optional inert speaker name.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_speaker(&mut self, value: Option<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_speaker(value)
            .map_err(scoped::staging("broadcast.set_speaker"))?)
    }

    /// Set one optional inert contact name.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_contact(&mut self, value: Option<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_contact(value)
            .map_err(scoped::staging("broadcast.set_contact"))?)
    }

    /// Set one optional inert remote server name.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_remote_server_name(&mut self, value: Option<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_remote_server_name(value)
            .map_err(scoped::staging("broadcast.set_remote_server_name"))?)
    }

    /// Set one optional inert email address.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_email_address(&mut self, value: Option<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_email_address(value)
            .map_err(scoped::staging("broadcast.set_email_address"))?)
    }

    /// Set one optional inert email display name.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_email_name(&mut self, value: Option<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_email_name(value)
            .map_err(scoped::staging("broadcast.set_email_name"))?)
    }

    /// Set one optional inert HTTP chat URL.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_chat_url(&mut self, value: Option<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_chat_url(value)
            .map_err(scoped::staging("broadcast.set_chat_url"))?)
    }

    /// Set one optional inert archive directory.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_archive_directory(&mut self, value: Option<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_archive_directory(value)
            .map_err(scoped::staging("broadcast.set_archive_directory"))?)
    }

    /// Set one optional inert `NetShow` base directory.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_netshow_files_base_directory(&mut self, value: Option<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_netshow_files_base_directory(value)
            .map_err(scoped::staging(
                "broadcast.set_netshow_files_base_directory",
            ))?)
    }

    /// Set one optional inert `NetShow` directory.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_netshow_files_directory(&mut self, value: Option<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_netshow_files_directory(value)
            .map_err(scoped::staging("broadcast.set_netshow_files_directory"))?)
    }

    /// Set one optional inert `NetShow` server name.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_netshow_server_name(&mut self, value: Option<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_netshow_server_name(value)
            .map_err(scoped::staging("broadcast.set_netshow_server_name"))?)
    }

    /// Set one optional inert calendar entry identifier.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_entry_id(&mut self, value: Option<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_entry_id(value)
            .map_err(scoped::staging("broadcast.set_entry_id"))?)
    }

    /// Set the inert `PowerPoint` files base directory.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_ppt_files_base_directory(&mut self, value: impl Into<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_ppt_files_base_directory(value)
            .map_err(scoped::staging("broadcast.set_ppt_files_base_directory"))?)
    }

    /// Set the inert `PowerPoint` files directory.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_ppt_files_directory(&mut self, value: impl Into<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_ppt_files_directory(value)
            .map_err(scoped::staging("broadcast.set_ppt_files_directory"))?)
    }

    /// Set the inert `PowerPoint` files base URL or UNC path.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_ppt_files_base_url(&mut self, value: impl Into<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_ppt_files_base_url(value)
            .map_err(scoped::staging("broadcast.set_ppt_files_base_url"))?)
    }

    /// Set the inert user-name file fragment.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_user_name(&mut self, value: impl Into<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_user_name(value)
            .map_err(scoped::staging("broadcast.set_user_name"))?)
    }

    /// Set the inert broadcast date/time file fragment.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_broadcast_date_time(&mut self, value: impl Into<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_broadcast_date_time(value)
            .map_err(scoped::staging("broadcast.set_broadcast_date_time"))?)
    }

    /// Set the inert presentation-name file fragment.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_presentation_name(&mut self, value: impl Into<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_presentation_name(value)
            .map_err(scoped::staging("broadcast.set_presentation_name"))?)
    }

    /// Set the inert ASD file UNC path.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_asd_file_name(&mut self, value: impl Into<String>) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_asd_file_name(value)
            .map_err(scoped::staging("broadcast.set_asd_file_name"))?)
    }

    /// Replace the typed broadcast flags and inert timestamps atomically.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// request, leaving the edit unchanged.
    pub fn set_properties(&mut self, value: BroadcastProperties) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .set_properties(value)
            .map_err(scoped::staging("broadcast.set_properties"))?)
    }

    /// Publish the staged broadcast as a new package. An edit with no net
    /// change publishes the source package unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Signed`] for a changed write to a signed
    /// package and [`super::Error::Package`] when the result cannot be
    /// republished.
    pub fn commit(self) -> Result<Commit> {
        self.state
            .get_or_try_init(|| self.initialize_state())
            .map_err(super::Error::from)?;
        let Self {
            base,
            position,
            state,
        } = self;
        let State {
            native,
            document,
            inner,
        } = state.into_inner().ok_or_else(|| {
            Failure::Package(crate::package::Error::Corrupted(
                "PPT broadcast edit state was not initialized".into(),
            ))
        })?;
        let (snapshot, _) = inner
            .commit()
            .map_err(scoped::publication("broadcast.commit"))?
            .into_parts();
        let located = Located::find(&document)?;
        let after = located.replace(&document, position, snapshot.bytes())?;
        Ok(super::published(base.publish_owner(
            &native,
            Owner::Document,
            document,
            after,
        )?))
    }
}

/// The broadcast containers of the live document's `___PPT9` tag.
struct Located<'a> {
    children: Vec<&'a [u8]>,
    info: usize,
    blob: crate::binary_tag_edit::Blob<'a>,
    broadcasts: Vec<usize>,
}

impl<'a> Located<'a> {
    fn find(document: &'a [u8]) -> std::result::Result<Self, Failure> {
        let children = wire::child_list(document)?;
        let mut info = None;
        for (index, child) in children.iter().enumerate() {
            if wire::record_type(child)? == RecordType::DocInfoList as u16
                && info.replace(index).is_some()
            {
                return Err(scoped::unsupported_target(
                    OPERATION,
                    "the document has several information lists",
                ));
            }
        }
        let info = info.ok_or_else(|| {
            scoped::not_found(OPERATION, "the document has no broadcast description")
        })?;
        let blob = crate::binary_tag_edit::locate(
            children[info],
            RecordType::DocInfoList as u16,
            "___PPT9",
            "broadcast",
        )?;
        let mut broadcasts = Vec::new();
        for (index, record) in blob.records().iter().enumerate() {
            if wire::record_type(record)? == BROADCAST_CONTAINER_RECORD_TYPE {
                broadcasts.push(index);
            }
        }
        Ok(Self {
            children,
            info,
            blob,
            broadcasts,
        })
    }

    fn container(&self, position: usize) -> std::result::Result<&'a [u8], Failure> {
        self.broadcasts
            .get(position)
            .map(|index| self.blob.records()[*index])
            .ok_or_else(|| scoped::not_found(OPERATION, "no broadcast description at the position"))
    }

    fn replace(
        &self,
        document: &[u8],
        position: usize,
        replacement: &[u8],
    ) -> std::result::Result<Vec<u8>, Failure> {
        let target = *self.broadcasts.get(position).ok_or_else(|| {
            scoped::not_found(OPERATION, "no broadcast description at the position")
        })?;
        let records = self
            .blob
            .records()
            .iter()
            .enumerate()
            .map(|(index, record)| {
                if index == target {
                    replacement.to_vec()
                } else {
                    record.to_vec()
                }
            })
            .collect::<Vec<_>>();
        let list = self.blob.rebuild(self.children[self.info], &records, &[])?;
        let mut data = Vec::with_capacity(document.len().saturating_add(replacement.len()));
        for (index, child) in self.children.iter().enumerate() {
            if index == self.info {
                data.extend_from_slice(&list);
            } else {
                data.extend_from_slice(child);
            }
        }
        Ok(wire::rebuild(document, &data)?)
    }
}

impl Package {
    /// Start an edit of the broadcast description at a zero-based position
    /// among the live document's `PowerPoint` 9 broadcast descriptions.
    ///
    /// # Errors
    ///
    /// An encrypted presentation is refused immediately. A missing or
    /// unreadable description is reported by the first projection, staging
    /// operation, or commit.
    pub fn edit_broadcast(&self, position: Position) -> Result<Edit> {
        Edit::new(self, position)
    }

    /// Apply a broadcast patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::PatchConflict`] when this package is not the
    /// patch's source.
    pub fn apply_broadcast(&self, patch: &Patch) -> Result<Commit> {
        super::apply_patch(self, patch)
    }
}
