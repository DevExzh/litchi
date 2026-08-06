//! Typed PowerPoint presentation-broadcast metadata.

use crate::records::Record;
use crate::slide_sync::SystemTime;

/// Fixed `BroadcastDocInfoAtom` values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BroadcastProperties {
    pub send_audio: bool,
    pub send_video: bool,
    pub camera_remote: bool,
    pub use_netshow: bool,
    pub use_other_server: bool,
    pub can_email: bool,
    pub can_chat: bool,
    pub archive: bool,
    pub speaker_notes: bool,
    pub quarter_screen: bool,
    pub show_tools: bool,
    pub record_only: bool,
    pub start_time: SystemTime,
    pub end_time: SystemTime,
}

/// One fully typed `BroadcastDocInfo9Container`.
///
/// Paths, server names, URLs, calendar identifiers, and all capability flags
/// are inert metadata. Parsing or writing this value never starts a broadcast,
/// connects to a server, sends mail, opens a URL, or reads an ASD file.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Broadcast {
    pub title: Option<String>,
    pub description: Option<String>,
    pub speaker: Option<String>,
    pub contact: Option<String>,
    pub remote_server_name: Option<String>,
    pub email_address: Option<String>,
    pub email_name: Option<String>,
    pub chat_url: Option<String>,
    pub archive_directory: Option<String>,
    pub netshow_files_base_directory: Option<String>,
    pub netshow_files_directory: Option<String>,
    pub netshow_server_name: Option<String>,
    pub ppt_files_base_directory: String,
    pub ppt_files_directory: String,
    pub ppt_files_base_url: String,
    pub user_name: String,
    pub broadcast_date_time: String,
    pub presentation_name: String,
    pub asd_file_name: String,
    pub entry_id: Option<String>,
    pub properties: BroadcastProperties,
}

/// All PowerPoint 9 broadcast descriptions in document order.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Broadcasts {
    pub broadcasts: Vec<Broadcast>,
}

/// One unmodeled child retained by a broadcast snapshot.
///
/// The record is inert and is never interpreted as a command, path, URL, or
/// network target. Its header and payload are retained so a semantic edit to
/// known broadcast fields does not discard future-version or producer-
/// specific children.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnknownRecord {
    pub(crate) record: Record,
    pub(crate) known_before: usize,
}

impl UnknownRecord {
    /// Original raw record type value.
    pub fn record_type(&self) -> u16 {
        self.record.record_type_raw
    }

    /// Original record version nibble.
    pub fn version(&self) -> u16 {
        self.record.version
    }

    /// Original record instance.
    pub fn instance(&self) -> u16 {
        self.record.instance
    }

    /// Borrow the opaque record payload.
    pub fn data(&self) -> &[u8] {
        &self.record.data
    }

    /// Reconstruct the exact opaque record header and payload.
    pub fn to_record_bytes(&self) -> crate::package::Result<Vec<u8>> {
        super::codec::record_bytes(
            self.record.version,
            self.record.instance,
            self.record.record_type_raw,
            &self.record.data,
        )
    }

    pub(crate) fn from_record(record: &Record, known_before: usize) -> Self {
        Self {
            record: record.clone(),
            known_before,
        }
    }
}
