//! Immutable semantic view used by the authoring pipeline.
//!
//! `Builder` remains the ergonomic editing facade.  This borrowed view keeps
//! validation and package emission from reaching through the facade's fields
//! ad hoc, without cloning slides, media, or metadata.

use super::Builder;
use crate::Slide;
use crate::model::declaration;
use crate::model::media::EmbeddedMedia;
use crate::model::page_layout;
use crate::model::page_metadata;
use crate::model::settings::Settings;
use std::collections::BTreeMap;

#[derive(Clone, Copy)]
pub(super) struct Snapshot<'a> {
    pub(super) slides: &'a [Slide],
    pub(super) media_files: &'a BTreeMap<String, EmbeddedMedia>,
    pub(super) settings: Option<&'a Settings>,
    pub(super) declarations: Option<&'a declaration::Collection>,
    pub(super) page_metadata: Option<&'a page_metadata::Collection>,
    pub(super) page_layouts: &'a page_layout::Collection,
}

impl Builder {
    pub(super) fn snapshot(&self) -> Snapshot<'_> {
        Snapshot {
            slides: &self.slides,
            media_files: &self.media_files,
            settings: self.settings.as_ref(),
            declarations: self.declarations.as_ref(),
            page_metadata: self.page_metadata.as_ref(),
            page_layouts: &self.page_layouts,
        }
    }
}
