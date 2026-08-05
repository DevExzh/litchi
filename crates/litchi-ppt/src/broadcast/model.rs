//! Typed PowerPoint presentation-broadcast metadata.

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
