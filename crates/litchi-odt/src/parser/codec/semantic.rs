//! Typed in-flight state and model construction for the ODT XML codecs.

use super::super::model::{ChangeType, Comment, Section, TrackChange};
use litchi_core::{Error, Result};
use std::collections::HashMap;

pub(super) struct ActiveTrackedChange {
    pub(super) id: String,
    pub(super) xml_id: Option<String>,
    pub(super) author: Option<String>,
    pub(super) date: Option<String>,
    pub(super) comment: String,
    pub(super) change_type: Option<ChangeType>,
    pub(super) style_name: Option<String>,
    pub(super) merge_last_paragraph: Option<bool>,
    pub(super) content: String,
    pub(super) depth: usize,
    pub(super) kind_depth: Option<usize>,
    pub(super) change_info_depth: Option<usize>,
    pub(super) change_info_seen: bool,
    pub(super) creator_depth: Option<usize>,
    pub(super) date_depth: Option<usize>,
    pub(super) comment_depth: Option<usize>,
    pub(super) comment_seen: bool,
    pub(super) paragraph_depth: Option<usize>,
    pub(super) seen_paragraph: bool,
}

impl ActiveTrackedChange {
    pub(super) fn new(id: String, xml_id: Option<String>) -> Self {
        Self {
            id,
            xml_id,
            author: None,
            date: None,
            comment: String::new(),
            change_type: None,
            style_name: None,
            merge_last_paragraph: None,
            content: String::new(),
            depth: 1,
            kind_depth: None,
            change_info_depth: None,
            change_info_seen: false,
            creator_depth: None,
            date_depth: None,
            comment_depth: None,
            comment_seen: false,
            paragraph_depth: None,
            seen_paragraph: false,
        }
    }

    pub(super) fn finish(self) -> Result<TrackChange> {
        let change_type = self.change_type.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "changed region '{}' has no change declaration",
                self.id
            ))
        })?;
        if !self.change_info_seen {
            return Err(Error::InvalidFormat(format!(
                "changed region '{}' has no office:change-info",
                self.id
            )));
        }
        Ok(TrackChange {
            id: self.id,
            xml_id: self.xml_id,
            author: self.author,
            date: self.date,
            comment: self.comment_seen.then_some(self.comment),
            change_type,
            style_name: self.style_name,
            merge_last_paragraph: self.merge_last_paragraph,
            content: self.content,
        })
    }
}

pub(super) struct PendingChangeRange {
    pub(super) text: String,
    pub(super) seen_paragraph: bool,
}

#[derive(Default)]
pub(super) struct ChangeRangeState {
    pub(super) pending: HashMap<String, PendingChangeRange>,
    pub(super) completed: HashMap<String, Vec<String>>,
    pub(super) completed_count: usize,
}

pub(super) struct ActiveComment {
    pub(super) comment: Comment,
    pub(super) depth: usize,
    pub(super) creator_depth: Option<usize>,
    pub(super) date_depth: Option<usize>,
    pub(super) fallback_date_depth: Option<usize>,
    pub(super) fallback_date: String,
    pub(super) paragraph_depth: Option<usize>,
    pub(super) seen_paragraph: bool,
}

impl ActiveComment {
    pub(super) fn new(id: String) -> Self {
        Self {
            comment: Comment {
                id,
                author: None,
                date: None,
                content: String::new(),
                reference: None,
            },
            depth: 1,
            creator_depth: None,
            date_depth: None,
            fallback_date_depth: None,
            fallback_date: String::new(),
            paragraph_depth: None,
            seen_paragraph: false,
        }
    }

    pub(super) fn finish(mut self) -> Comment {
        if self.comment.date.is_none() && !self.fallback_date.is_empty() {
            self.comment.date = Some(self.fallback_date);
        }
        self.comment
    }
}

pub(super) struct PendingAnnotation {
    pub(super) name: String,
    pub(super) text: String,
    pub(super) seen_paragraph: bool,
}

impl PendingAnnotation {
    pub(super) fn new(name: String, seen_paragraph: bool) -> Self {
        Self {
            name,
            text: String::new(),
            seen_paragraph,
        }
    }
}

pub(super) struct ActiveSection {
    pub(super) section: Section,
    pub(super) depth: usize,
    pub(super) paragraph_depth: Option<usize>,
    pub(super) seen_paragraph: bool,
    pub(super) order: usize,
    pub(super) source_depth: Option<usize>,
}

impl ActiveSection {
    pub(super) fn new(section: Section, order: usize) -> Self {
        Self {
            section,
            depth: 1,
            paragraph_depth: None,
            seen_paragraph: false,
            order,
            source_depth: None,
        }
    }

    pub(super) fn into_ordered(self) -> (usize, Section) {
        (self.order, self.section)
    }
}
