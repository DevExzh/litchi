//! Semantic validation for the RTF writer.
//!
//! This layer validates snapshot ownership, ordering, and cross-story
//! invariants before the output layer emits bytes.

use super::{
    BodyStoryEvent, CellStoryEvent, Field, FieldOwner, HeaderFooterType, NavigationEntry, Revision,
    RevisionType, RtfDocument, RtfWriter, Section, StoryEvent, StoryField, Table, Write, field, io,
};

impl<W: Write> RtfWriter<W> {
    pub(super) fn validate_section_boundary_mapping(
        sections: &[Section<'_>],
        body_story_events: &[BodyStoryEvent],
    ) -> io::Result<()> {
        let first_section_is_boundary_scoped = body_story_events.iter().any(|event| {
            matches!(
                event,
                BodyStoryEvent::SectionBreak(section_break)
                    if section_break.next_section == Some(0)
            )
        });
        let mut next_section_index = if first_section_is_boundary_scoped {
            0
        } else {
            usize::from(!sections.is_empty())
        };
        for event in body_story_events {
            if let BodyStoryEvent::SectionBreak(section_break) = *event
                && let Some(index) = section_break.next_section
            {
                if index != next_section_index || index >= sections.len() {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF section boundary has an invalid or out-of-order section reference",
                    ));
                }
                next_section_index += 1;
            }
        }
        if next_section_index != sections.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF section definitions are missing main-story boundaries",
            ));
        }
        Ok(())
    }

    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    pub(super) fn validate_table_story_metadata_ownership(doc: &RtfDocument<'_>) -> io::Result<()> {
        let mut navigation_owners = vec![0u8; doc.navigation_entries().len()];
        let mut revision_owners = vec![0u8; doc.revisions().len()];
        let mut body_starts = vec![false; doc.revisions().len()];
        let mut body_ends = vec![false; doc.revisions().len()];
        let mut body_deletions = vec![false; doc.revisions().len()];
        for event in doc.body_story_events() {
            match *event {
                BodyStoryEvent::NavigationEntry(index) => {
                    let owner = navigation_owners.get_mut(index).ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF body navigation index is out of range",
                        )
                    })?;
                    *owner = owner.checked_add(1).ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF navigation ownership overflow",
                        )
                    })?;
                },
                BodyStoryEvent::RevisionStart(index) => {
                    let revision = doc.revisions().get(index).ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF body revision index is out of range",
                        )
                    })?;
                    let seen = body_starts.get_mut(index).ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF body revision index is out of range",
                        )
                    })?;
                    if revision.revision_type != RevisionType::Insertion
                        || std::mem::replace(seen, true)
                    {
                        return Err(io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF body revision start has the wrong kind or is duplicated",
                        ));
                    }
                },
                BodyStoryEvent::RevisionEnd(index) => {
                    let revision = doc.revisions().get(index).ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF body revision index is out of range",
                        )
                    })?;
                    let seen = body_ends.get_mut(index).ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF body revision index is out of range",
                        )
                    })?;
                    if revision.revision_type != RevisionType::Insertion
                        || std::mem::replace(seen, true)
                    {
                        return Err(io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF body revision end has the wrong kind or is duplicated",
                        ));
                    }
                },
                BodyStoryEvent::RevisionDeletion(index) => {
                    let revision = doc.revisions().get(index).ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF body revision index is out of range",
                        )
                    })?;
                    let seen = body_deletions.get_mut(index).ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF body revision index is out of range",
                        )
                    })?;
                    if revision.revision_type != RevisionType::Deletion
                        || std::mem::replace(seen, true)
                    {
                        return Err(io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF body deletion has the wrong kind or is duplicated",
                        ));
                    }
                },
                _ => {},
            }
        }
        for ((((revision, start), end), deletion), owner) in doc
            .revisions()
            .iter()
            .zip(&body_starts)
            .zip(&body_ends)
            .zip(&body_deletions)
            .zip(&mut revision_owners)
        {
            let owned = match revision.revision_type {
                RevisionType::Insertion => *start || *end,
                RevisionType::Deletion => *deletion,
                RevisionType::FormatChange | RevisionType::MovedFrom | RevisionType::MovedTo => {
                    false
                },
            };
            if owned {
                let valid = match revision.revision_type {
                    RevisionType::Insertion => *start && *end && !*deletion,
                    RevisionType::Deletion => *deletion && !*start && !*end,
                    RevisionType::FormatChange
                    | RevisionType::MovedFrom
                    | RevisionType::MovedTo => false,
                };
                if !valid {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF body revision ownership is incomplete or conflicting",
                    ));
                }
                *owner = 1;
            }
        }
        for table in doc.tables() {
            Self::validate_table_metadata_tree(
                table,
                doc.navigation_entries(),
                doc.revisions(),
                &mut navigation_owners,
                &mut revision_owners,
            )?;
        }
        if navigation_owners.iter().any(|owners| *owners != 1) {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "every RTF navigation entry must be owned by exactly one story",
            ));
        }
        if revision_owners.iter().any(|owners| *owners != 1) {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "every RTF revision must be owned by exactly one story",
            ));
        }
        Ok(())
    }

    pub(super) fn validate_table_metadata_tree(
        table: &Table<'_>,
        navigation_entries: &[NavigationEntry<'_>],
        revisions: &[Revision<'_>],
        navigation_owners: &mut [u8],
        revision_owners: &mut [u8],
    ) -> io::Result<()> {
        for row in table.rows() {
            for cell in row.cells() {
                let mut starts = vec![false; revisions.len()];
                let mut ends = vec![false; revisions.len()];
                let mut deletions = vec![false; revisions.len()];
                for event in cell.story_events() {
                    match *event {
                        CellStoryEvent::NavigationEntry(reference) => {
                            let entry =
                                navigation_entries.get(reference.index).ok_or_else(|| {
                                    io::Error::new(
                                        io::ErrorKind::InvalidInput,
                                        "RTF table-cell navigation index is out of range",
                                    )
                                })?;
                            if entry.position() != reference.position
                                || cell
                                    .text()
                                    .get(reference.position..reference.position)
                                    .is_none()
                            {
                                return Err(io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF table-cell navigation anchor is invalid",
                                ));
                            }
                            let owners =
                                navigation_owners.get_mut(reference.index).ok_or_else(|| {
                                    io::Error::new(
                                        io::ErrorKind::InvalidInput,
                                        "RTF table-cell navigation index is out of range",
                                    )
                                })?;
                            *owners = owners.checked_add(1).ok_or_else(|| {
                                io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF navigation ownership overflow",
                                )
                            })?;
                        },
                        CellStoryEvent::RevisionStart(reference) => {
                            let revision = revisions.get(reference.index).ok_or_else(|| {
                                io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF table-cell revision index is out of range",
                                )
                            })?;
                            let seen = starts.get_mut(reference.index).ok_or_else(|| {
                                io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF table-cell revision index is out of range",
                                )
                            })?;
                            if revision.revision_type != RevisionType::Insertion
                                || revision.position != reference.position
                                || cell.text().get(revision.position..revision.range_end)
                                    != Some(revision.content.as_ref())
                                || std::mem::replace(seen, true)
                            {
                                return Err(io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF table-cell revision start is invalid, duplicated, or has the wrong kind",
                                ));
                            }
                        },
                        CellStoryEvent::RevisionEnd(reference) => {
                            let revision = revisions.get(reference.index).ok_or_else(|| {
                                io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF table-cell revision index is out of range",
                                )
                            })?;
                            let seen = ends.get_mut(reference.index).ok_or_else(|| {
                                io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF table-cell revision index is out of range",
                                )
                            })?;
                            if revision.revision_type != RevisionType::Insertion
                                || revision.range_end != reference.position
                                || std::mem::replace(seen, true)
                            {
                                return Err(io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF table-cell revision end is invalid, duplicated, or has the wrong kind",
                                ));
                            }
                        },
                        CellStoryEvent::RevisionDeletion(reference) => {
                            let revision = revisions.get(reference.index).ok_or_else(|| {
                                io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF table-cell revision index is out of range",
                                )
                            })?;
                            let seen = deletions.get_mut(reference.index).ok_or_else(|| {
                                io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF table-cell revision index is out of range",
                                )
                            })?;
                            if revision.revision_type != RevisionType::Deletion
                                || revision.position != reference.position
                                || cell
                                    .text()
                                    .get(reference.position..reference.position)
                                    .is_none()
                                || std::mem::replace(seen, true)
                            {
                                return Err(io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF table-cell deletion is invalid, duplicated, or has the wrong kind",
                                ));
                            }
                        },
                        CellStoryEvent::NestedTable(_)
                        | CellStoryEvent::Drawing(_)
                        | CellStoryEvent::Field(_)
                        | CellStoryEvent::PageBreak(_)
                        | CellStoryEvent::ColumnBreak(_) => {},
                    }
                }
                for ((((revision, start), end), deletion), owners) in revisions
                    .iter()
                    .zip(&starts)
                    .zip(&ends)
                    .zip(&deletions)
                    .zip(&mut *revision_owners)
                {
                    let touched = *start || *end || *deletion;
                    if touched {
                        let valid = match revision.revision_type {
                            RevisionType::Insertion => *start && *end && !*deletion,
                            RevisionType::Deletion => *deletion && !*start && !*end,
                            RevisionType::FormatChange
                            | RevisionType::MovedFrom
                            | RevisionType::MovedTo => false,
                        };
                        if !valid {
                            return Err(io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF table-cell revision ownership is incomplete or conflicting",
                            ));
                        }
                        *owners = owners.checked_add(1).ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF revision ownership overflow",
                            )
                        })?;
                    }
                }
                for nested in cell.nested_tables() {
                    Self::validate_table_metadata_tree(
                        &nested.table,
                        navigation_entries,
                        revisions,
                        navigation_owners,
                        revision_owners,
                    )?;
                }
            }
        }
        Ok(())
    }

    pub(super) fn mark_owned_field(
        reference: StoryField,
        owner: FieldOwner,
        fields: &[Field<'_>],
        seen: &mut [bool],
    ) -> io::Result<()> {
        let field = fields
            .get(reference.field_index)
            .filter(|field| {
                field.owner == owner
                    && field.position == reference.position
                    && field.range_end == reference.position
            })
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF story has an invalid generic-field owner or reference",
                )
            })?;
        field
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        let seen_slot = seen.get_mut(reference.field_index).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF generic-field index is out of range",
            )
        })?;
        if std::mem::replace(seen_slot, true) {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF generic field is referenced by multiple owning stories",
            ));
        }
        Ok(())
    }

    pub(super) fn mark_table_fields(
        table: &Table<'_>,
        depth: u8,
        fields: &[Field<'_>],
        seen: &mut [bool],
    ) -> io::Result<()> {
        for row in table.rows() {
            for cell in row.cells() {
                for event in cell.story_events() {
                    if let CellStoryEvent::Field(reference) = *event {
                        Self::mark_owned_field(
                            reference,
                            FieldOwner::TableCell(depth),
                            fields,
                            seen,
                        )?;
                    }
                }
                for nested in cell.nested_tables() {
                    Self::mark_table_fields(
                        &nested.table,
                        depth.checked_add(1).ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF table nesting depth overflow",
                            )
                        })?,
                        fields,
                        seen,
                    )?;
                }
            }
        }
        Ok(())
    }

    pub(super) fn validate_generic_field_ownership(doc: &RtfDocument<'_>) -> io::Result<()> {
        let fields = doc.fields();
        if fields.len() > field::MAX_GENERIC_FIELDS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF generic field count exceeds the safety limit",
            ));
        }
        let mut seen = vec![false; fields.len()];
        for event in doc.body_story_events() {
            if let BodyStoryEvent::Field(index) = *event {
                let position = fields.get(index).map_or(usize::MAX, |field| field.position);
                Self::mark_owned_field(
                    StoryField {
                        field_index: index,
                        position,
                    },
                    FieldOwner::Body,
                    fields,
                    &mut seen,
                )?;
            }
        }
        for section in doc.sections() {
            for hf in &section.headers_footers {
                let owner = match hf.header_type {
                    HeaderFooterType::Header
                    | HeaderFooterType::HeaderFirst
                    | HeaderFooterType::HeaderLeft
                    | HeaderFooterType::HeaderRight => FieldOwner::Header,
                    HeaderFooterType::Footer
                    | HeaderFooterType::FooterFirst
                    | HeaderFooterType::FooterLeft
                    | HeaderFooterType::FooterRight => FieldOwner::Footer,
                };
                for event in &hf.story_events {
                    if let StoryEvent::Field(reference) = *event {
                        Self::mark_owned_field(reference, owner, fields, &mut seen)?;
                    }
                }
            }
        }
        for note in doc.notes() {
            for event in &note.story_events {
                if let StoryEvent::Field(reference) = *event {
                    Self::mark_owned_field(
                        reference,
                        if note.is_footnote {
                            FieldOwner::Footnote
                        } else {
                            FieldOwner::Endnote
                        },
                        fields,
                        &mut seen,
                    )?;
                }
            }
        }
        for table in doc.tables() {
            Self::mark_table_fields(table, 1, fields, &mut seen)?;
        }
        for field in fields {
            for event in &field.result_events {
                if let StoryEvent::Field(reference) = *event {
                    Self::mark_owned_field(reference, FieldOwner::FieldResult, fields, &mut seen)?;
                }
            }
        }
        if fields.iter().zip(seen).any(|(field, was_seen)| {
            matches!(
                field.owner,
                FieldOwner::Body
                    | FieldOwner::Header
                    | FieldOwner::Footer
                    | FieldOwner::Footnote
                    | FieldOwner::Endnote
                    | FieldOwner::TableCell(_)
                    | FieldOwner::FieldResult
            ) && !was_seen
        }) {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF generic field lacks its concrete owning-story event",
            ));
        }
        Ok(())
    }
}
