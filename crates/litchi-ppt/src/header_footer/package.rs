//! Presentation record-tree collection and placement for header/footer data.

use super::codec::{
    LOCAL_INSTANCE, MAX_AGGREGATE_TEXT_BYTES, MAX_HEADER_FOOTER_ENTRIES, MAX_SCANNED_RECORDS,
    NOTES_AND_HANDOUTS_INSTANCE, PRESENTATION_SLIDES_INSTANCE, corrupted, validated_encoded_len,
};
use super::model::{
    PowerPointHeaderFooter, PowerPointHeaderFooterDisplayText, PowerPointHeaderFooterParent,
    PowerPointHeaderFooterParentOrdinal, PowerPointHeaderFooterScope, PowerPointHeaderFooters,
    PowerPointScopedHeaderFooterDisplayText,
};
use crate::consts::PptRecordType;
use crate::package::Result;
use crate::records::PptRecord;

impl PowerPointHeaderFooters {
    pub(crate) fn parse_record_tree(records: &[&PptRecord]) -> Result<Self> {
        if records.len() > MAX_SCANNED_RECORDS {
            return Err(corrupted(
                "PowerPoint record tree exceeds the header/footer scan limit",
            ));
        }
        let document_count = records
            .iter()
            .filter(|record| record.record_type == PptRecordType::Document)
            .count();
        if document_count != 1 {
            return Err(corrupted(
                "PowerPoint must contain exactly one Document container",
            ));
        }

        let total_containers = records
            .iter()
            .filter(|record| record.record_type == PptRecordType::HeadersFooters)
            .count();
        if total_containers > MAX_HEADER_FOOTER_ENTRIES {
            return Err(corrupted("too many PowerPoint header/footer containers"));
        }

        let mut entries = Vec::with_capacity(total_containers);
        let mut located = 0usize;
        let mut slide_ordinal = 0usize;
        let mut master_ordinal = 0usize;
        let mut aggregate = 0usize;

        for parent in records {
            match parent.record_type {
                PptRecordType::Document => {
                    let mut saw_slides = false;
                    let mut saw_notes = false;
                    for child in parent
                        .children
                        .iter()
                        .filter(|child| child.record_type == PptRecordType::HeadersFooters)
                    {
                        located += 1;
                        let scope = match child.instance {
                            PRESENTATION_SLIDES_INSTANCE if !saw_slides => {
                                saw_slides = true;
                                PowerPointHeaderFooterScope::PresentationSlides
                            },
                            NOTES_AND_HANDOUTS_INSTANCE if !saw_notes => {
                                saw_notes = true;
                                PowerPointHeaderFooterScope::NotesAndHandouts
                            },
                            PRESENTATION_SLIDES_INSTANCE | NOTES_AND_HANDOUTS_INSTANCE => {
                                return Err(corrupted(
                                    "duplicate document-level header/footer container",
                                ));
                            },
                            _ => {
                                return Err(corrupted(
                                    "invalid document-level header/footer instance",
                                ));
                            },
                        };
                        entries.push(PowerPointHeaderFooter::parse_record_bounded(
                            child,
                            scope,
                            &mut aggregate,
                        )?);
                    }
                },
                PptRecordType::Slide => {
                    locate_local(
                        parent,
                        PowerPointHeaderFooterParent::Slide,
                        slide_ordinal,
                        &mut located,
                        &mut aggregate,
                        &mut entries,
                    )?;
                    slide_ordinal += 1;
                },
                PptRecordType::MainMaster => {
                    locate_local(
                        parent,
                        PowerPointHeaderFooterParent::MainMaster,
                        master_ordinal,
                        &mut located,
                        &mut aggregate,
                        &mut entries,
                    )?;
                    master_ordinal += 1;
                },
                _ => {},
            }
        }
        if located != total_containers {
            return Err(corrupted(
                "HeadersFooters container has an invalid direct parent",
            ));
        }
        Ok(Self {
            entries,
            placeholder_displays: Vec::new(),
            placeholder_display_bytes: 0,
        })
    }

    pub(crate) fn attach_placeholder_display(
        &mut self,
        scope: PowerPointHeaderFooterScope,
        display: PowerPointHeaderFooterDisplayText,
    ) -> Result<()> {
        if self.placeholder_displays.len() == MAX_HEADER_FOOTER_ENTRIES {
            return Err(corrupted("too many PowerPoint placeholder displays"));
        }
        if self
            .placeholder_displays
            .iter()
            .any(|existing| existing.scope == scope)
        {
            return Err(corrupted("duplicate PowerPoint placeholder display scope"));
        }
        let mut display_bytes = 0usize;
        for value in [
            display.user_date.as_deref(),
            display.header.as_deref(),
            display.footer.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            display_bytes = display_bytes
                .checked_add(validated_encoded_len(value)?)
                .ok_or_else(|| corrupted("placeholder display aggregate size overflow"))?;
        }
        self.placeholder_display_bytes = self
            .placeholder_display_bytes
            .checked_add(display_bytes)
            .ok_or_else(|| corrupted("placeholder display aggregate size overflow"))?;
        if self.placeholder_display_bytes > MAX_AGGREGATE_TEXT_BYTES {
            return Err(corrupted(
                "placeholder display strings exceed the aggregate resource limit",
            ));
        }
        if let Some(entry) = self.entries.iter_mut().find(|entry| entry.scope == scope) {
            entry.placeholder_display = Some(display.clone());
        }
        self.placeholder_displays
            .push(PowerPointScopedHeaderFooterDisplayText {
                scope,
                text: display,
            });
        Ok(())
    }

    pub(crate) fn has_scope(&self, scope: PowerPointHeaderFooterScope) -> bool {
        self.entries.iter().any(|entry| entry.scope == scope)
    }
}

fn locate_local(
    parent_record: &PptRecord,
    parent: PowerPointHeaderFooterParent,
    ordinal: usize,
    located: &mut usize,
    aggregate: &mut usize,
    entries: &mut Vec<PowerPointHeaderFooter>,
) -> Result<()> {
    let mut containers = parent_record
        .children
        .iter()
        .filter(|child| child.record_type == PptRecordType::HeadersFooters);
    let Some(container) = containers.next() else {
        return Ok(());
    };
    if containers.next().is_some() {
        return Err(corrupted(
            "slide or main master has duplicate header/footer containers",
        ));
    }
    if container.instance != LOCAL_INSTANCE {
        return Err(corrupted(
            "local header/footer container has a nonzero instance",
        ));
    }
    *located += 1;
    let scope = PowerPointHeaderFooterScope::Local {
        parent,
        parent_ordinal: PowerPointHeaderFooterParentOrdinal(ordinal),
    };
    entries.push(PowerPointHeaderFooter::parse_record_bounded(
        container, scope, aggregate,
    )?);
    Ok(())
}
