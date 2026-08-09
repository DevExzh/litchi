//! Read-side speaker-notes support for binary `PowerPoint` files.

use super::directory::SlideDirectory;
use super::types::Slide;
use crate::consts::RecordType;
use crate::odraw::ShapeExt as _;
use crate::package::{Error, RecordLimits, Result};
use crate::persist::PersistMapping;
use crate::records::Record;
use crate::shapes::ShapeEnum;
use once_cell::unsync::OnceCell;
use std::collections::HashMap;

const NOTES_LIST_INSTANCE: u16 = 2;
const SLIDE_ATOM_SIZE: usize = 24;
const NOTES_PERSIST_ATOM_SIZE: usize = 20;
const NOTES_ATOM_SIZE: usize = 8;

#[derive(Debug, Clone, Copy)]
pub(crate) struct NoteDescriptor {
    pub notes_id: u32,
    pub persist_id: u32,
    pub offset: usize,
    pub expected_slide_id: Option<u32>,
}

#[derive(Debug, Default)]
pub(crate) struct NotesIndex {
    notes: HashMap<u32, u32>,
    slide_ids: HashMap<u32, u32>,
    error: Option<String>,
}

impl NotesIndex {
    #[cfg(all(test, feature = "vba-inspection"))]
    #[allow(
        dead_code,
        reason = "retained for focused presentation fixture construction"
    )]
    pub(crate) fn build(document_data: &[u8], slide_directory: &SlideDirectory) -> Self {
        Self::build_with_limits(document_data, slide_directory, RecordLimits::default())
    }

    pub(crate) fn build_with_limits(
        document_data: &[u8],
        slide_directory: &SlideDirectory,
        limits: RecordLimits,
    ) -> Self {
        match Self::try_build_with_limits(document_data, slide_directory, limits) {
            Ok(index) => index,
            Err(error) => Self {
                error: Some(error.to_string()),
                ..Self::default()
            },
        }
    }

    #[cfg(test)]
    fn try_build(document_data: &[u8], slide_directory: &SlideDirectory) -> Result<Self> {
        Self::try_build_with_limits(document_data, slide_directory, RecordLimits::default())
    }

    fn try_build_with_limits(
        document_data: &[u8],
        slide_directory: &SlideDirectory,
        limits: RecordLimits,
    ) -> Result<Self> {
        let (document, _) =
            Record::parse_with_limits(document_data, slide_directory.document_offset(), limits)?;
        if document.record_type != RecordType::Document {
            return Err(Error::Corrupted(
                "live document persist object is not a DocumentContainer".to_string(),
            ));
        }
        let mut index = Self::default();
        index
            .slide_ids
            .try_reserve(slide_directory.entries().len())
            .map_err(|_err| Error::AllocationFailed("PPT notes slide index"))?;
        for entry in slide_directory.entries() {
            index.slide_ids.insert(entry.persist_id(), entry.slide_id());
        }
        let mut saw_notes_list = false;

        for list in document.extract_slide_list_with_texts() {
            let is_notes = list.get_instance() == NOTES_LIST_INSTANCE;
            if !is_notes {
                continue;
            }
            if saw_notes_list {
                return Err(Error::Corrupted(
                    "duplicate NotesListWithTextContainer".to_string(),
                ));
            }
            saw_notes_list = true;

            for atom in &list.children {
                if atom.record_type != RecordType::SlidePersistAtom {
                    continue;
                }
                Self::validate_persist_atom(atom)?;
                let persist_id = read_u32(&atom.data, 0, "persistIdRef")?;
                let identifier = read_u32(&atom.data, 12, "slide/notes identifier")?;

                if persist_id == 0 {
                    return Err(Error::Corrupted(
                        "SlidePersistAtom has a null persistIdRef".to_string(),
                    ));
                }
                if !(0x100..=0x7fff_ffff).contains(&identifier) {
                    return Err(Error::Corrupted(format!(
                        "invalid NotesId {identifier:#010x}"
                    )));
                }
                if index.notes.insert(identifier, persist_id).is_some() {
                    return Err(Error::Corrupted(format!(
                        "duplicate NotesId {identifier:#010x}"
                    )));
                }
            }
        }

        Ok(index)
    }

    fn validate_persist_atom(atom: &Record) -> Result<()> {
        if atom.version != 0 || atom.instance != 0 || atom.data.len() != NOTES_PERSIST_ATOM_SIZE {
            return Err(Error::Corrupted(format!(
                "invalid SlidePersistAtom header or length: version={}, instance={}, length={}",
                atom.version,
                atom.instance,
                atom.data.len()
            )));
        }
        Ok(())
    }

    pub(crate) fn descriptor(
        &self,
        slide: &Record,
        slide_persist_id: u32,
        persist_mapping: &PersistMapping,
    ) -> std::result::Result<Option<NoteDescriptor>, String> {
        let Some(slide_atom) = slide.find_child(RecordType::SlideAtom) else {
            return Err("SlideContainer is missing SlideAtom".to_string());
        };
        if slide_atom.data.len() != SLIDE_ATOM_SIZE {
            return Err(format!(
                "invalid SlideAtom length {}; expected {SLIDE_ATOM_SIZE}",
                slide_atom.data.len()
            ));
        }
        let notes_id = u32::from_le_bytes([
            slide_atom.data[16],
            slide_atom.data[17],
            slide_atom.data[18],
            slide_atom.data[19],
        ]);
        if notes_id == 0 {
            return Ok(None);
        }
        if let Some(error) = &self.error {
            return Err(error.clone());
        }
        let persist_id = self.notes.get(&notes_id).copied().ok_or_else(|| {
            format!("notesIdRef {notes_id:#010x} is not present in the notes list")
        })?;
        let offset = persist_mapping
            .get_offset(persist_id)
            .ok_or_else(|| format!("notes persist ID {persist_id} has no directory entry"))?;
        let offset_usize = usize::try_from(offset)
            .map_err(|_err| format!("notes persist offset does not fit usize: {offset}"))?;

        Ok(Some(NoteDescriptor {
            notes_id,
            persist_id,
            offset: offset_usize,
            expected_slide_id: self.slide_ids.get(&slide_persist_id).copied(),
        }))
    }
}

/// A notes page associated with a presentation slide.
#[allow(
    clippy::module_name_repetitions,
    reason = "`SpeakerNotes` is the established public API name re-exported as `slide::SpeakerNotes`; renaming it would break downstream crates"
)]
pub struct SpeakerNotes {
    notes_id: u32,
    persist_id: u32,
    slide_id_ref: u32,
    record: Record,
    shapes: OnceCell<Vec<ShapeEnum<'static>>>,
    text: OnceCell<String>,
}

impl SpeakerNotes {
    pub(crate) fn parse_with_limits(
        descriptor: NoteDescriptor,
        document_data: &[u8],
        limits: RecordLimits,
    ) -> Result<Self> {
        if descriptor
            .offset
            .checked_add(8)
            .is_none_or(|end| end > document_data.len())
        {
            return Err(Error::Corrupted(format!(
                "notes persist offset {} is outside the PowerPoint Document stream",
                descriptor.offset
            )));
        }
        let (record, _) = Record::parse_with_limits(document_data, descriptor.offset, limits)?;
        if record.record_type != RecordType::Notes || record.version != 0x0f || record.instance != 0
        {
            return Err(Error::Corrupted(format!(
                "persist ID {} does not resolve to a valid NotesContainer",
                descriptor.persist_id
            )));
        }
        let notes_atom = record
            .find_child(RecordType::NotesAtom)
            .ok_or_else(|| Error::Corrupted("NotesContainer is missing NotesAtom".to_string()))?;
        if notes_atom.version != 1
            || notes_atom.instance != 0
            || notes_atom.data.len() != NOTES_ATOM_SIZE
        {
            return Err(Error::Corrupted(format!(
                "invalid NotesAtom header or length: version={}, instance={}, length={}",
                notes_atom.version,
                notes_atom.instance,
                notes_atom.data.len()
            )));
        }
        let slide_id_ref = read_u32(&notes_atom.data, 0, "NotesAtom.slideIdRef")?;
        if slide_id_ref == 0 {
            return Err(Error::Corrupted(
                "notes slide has a null NotesAtom.slideIdRef".to_string(),
            ));
        }
        if descriptor
            .expected_slide_id
            .is_some_and(|expected| expected != slide_id_ref)
        {
            return Err(Error::Corrupted(format!(
                "NotesAtom.slideIdRef {slide_id_ref:#010x} does not match its slide"
            )));
        }

        Ok(Self {
            notes_id: descriptor.notes_id,
            persist_id: descriptor.persist_id,
            slide_id_ref,
            record,
            shapes: OnceCell::new(),
            text: OnceCell::new(),
        })
    }

    pub fn notes_id(&self) -> u32 {
        self.notes_id
    }

    pub fn persist_id(&self) -> u32 {
        self.persist_id
    }

    pub fn slide_id_ref(&self) -> u32 {
        self.slide_id_ref
    }

    /// Return this notes page's typed slide-level programmable tags (MS-PPT
    /// 2.5.19), when the `NotesContainer` carries a `SlideProgTagsContainer`.
    ///
    /// Tag payloads are inert: they are parsed and preserved, never executed,
    /// loaded, or resolved. Use
    /// [`crate::ProgTags::slide_extensions`] to decode the
    /// versioned binary-tag payloads into typed extension structs.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn programmable_tags(&self) -> Result<Option<crate::ProgTags>> {
        self.programmable_tags_with_limits(crate::ProgTagLimits::default())
    }

    /// Return notes programmable tags with caller-supplied resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn programmable_tags_with_limits(
        &self,
        limits: crate::ProgTagLimits,
    ) -> Result<Option<crate::ProgTags>> {
        crate::ProgTags::parse_slide(&self.record, limits)
    }

    /// Return the shapes drawn on this notes page.
    ///
    /// Shapes are parsed lazily on first access and then cached. Each
    /// returned shape is bound to its parsed source and refuses public
    /// mutation.
    ///
    /// # Errors
    ///
    /// Returns an error if the `PPDrawing` payload is malformed or a shape
    /// fails to decode.
    pub fn shapes(&self) -> Result<&[ShapeEnum<'static>]> {
        self.shapes
            .get_or_try_init(|| self.parse_shapes())
            .map(Vec::as_slice)
    }

    /// Return the body text of this notes page.
    ///
    /// The text is extracted lazily on first access and then cached.
    ///
    /// # Errors
    ///
    /// Returns an error if the `PPDrawing` payload is malformed or the notes
    /// shape tree cannot be traversed.
    pub fn text(&self) -> Result<&str> {
        self.text
            .get_or_try_init(|| self.extract_notes_body_text())
            .map(String::as_str)
    }

    fn parse_shapes(&self) -> Result<Vec<ShapeEnum<'static>>> {
        let Some(drawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(Vec::new());
        };
        let parsed = crate::odraw::parse(&drawing.data)?;
        let mut shapes = Vec::with_capacity(parsed.len());
        for shape in &parsed {
            if let Some(mut converted_shape) = Slide::<'static>::convert_odraw_to_shape_enum(shape)?
            {
                converted_shape.mark_source_bound_recursive();
                shapes.push(converted_shape);
            }
        }
        Ok(shapes)
    }

    fn extract_notes_body_text(&self) -> Result<String> {
        let Some(drawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(String::new());
        };
        let shapes = crate::odraw::parse(&drawing.data)?;
        let mut text = Vec::new();
        for shape in &shapes {
            collect_notes_body_text(shape, &mut text)?;
        }
        Ok(text.join("\n"))
    }
}

fn collect_notes_body_text(
    shape: &litchi_odraw::shape::Shape<'_>,
    text: &mut Vec<String>,
) -> Result<()> {
    const MAX_SHAPES: usize = 1_000_000;

    let mut pending = vec![shape];
    let mut visited = 0usize;
    while let Some(current) = pending.pop() {
        visited = visited
            .checked_add(1)
            .ok_or_else(|| Error::Corrupted("Notes shape count overflow".to_string()))?;
        if visited > MAX_SHAPES {
            return Err(Error::Corrupted(
                "Notes page exceeds the PPT shape limit".to_string(),
            ));
        }
        if current
            .placeholder()?
            .is_some_and(|placeholder| placeholder.kind == crate::PlaceholderKind::NotesBody)
            && let Some(value) = current.text()?.filter(|value| !value.is_empty())
        {
            text.push(value);
        }
        pending.extend(current.children().iter().rev());
    }
    Ok(())
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| Error::Corrupted(format!("{field} offset overflow")))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| Error::Corrupted(format!("truncated {field}")))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use crate::Package;

    fn record(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(8 + data.len());
        bytes.extend_from_slice(&(version | (instance << 4)).to_le_bytes());
        bytes.extend_from_slice(&record_type.to_le_bytes());
        bytes.extend_from_slice(&u32::try_from(data.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(data);
        bytes
    }

    fn persist_atom(persist_id: u32, identifier: u32) -> Vec<u8> {
        let mut data = Vec::with_capacity(20);
        data.extend_from_slice(&persist_id.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&identifier.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        record(0, 0, 0x03f3, &data)
    }

    fn document_with_notes(entries: &[(u32, u32)]) -> Vec<u8> {
        let children: Vec<u8> = entries
            .iter()
            .flat_map(|&(persist_id, notes_id)| persist_atom(persist_id, notes_id))
            .collect();
        let notes_list = record(0x0f, 2, 0x0ff0, &children);
        record(0x0f, 0, 0x03e8, &notes_list)
    }

    #[test]
    fn notes_index_uses_ids_not_list_order() {
        let directory = SlideDirectory::new_for_test(0);
        let index =
            NotesIndex::try_build(&document_with_notes(&[(9, 0x345), (7, 0x123)]), &directory)
                .expect("valid notes index");
        assert_eq!(index.notes.get(&0x123), Some(&7));
        assert_eq!(index.notes.get(&0x345), Some(&9));
    }

    #[test]
    fn notes_index_rejects_duplicate_and_malformed_atoms() {
        let duplicate = document_with_notes(&[(7, 0x123), (9, 0x123)]);
        let directory = SlideDirectory::new_for_test(0);
        assert!(NotesIndex::try_build(&duplicate, &directory).is_err());

        let malformed = record(0, 0, 0x03f3, &[0; 16]);
        let notes_list = record(0x0f, 2, 0x0ff0, &malformed);
        let document = record(0x0f, 0, 0x03e8, &notes_list);
        assert!(NotesIndex::try_build(&document, &directory).is_err());
    }

    #[test]
    fn notes_index_rejects_dangling_persist_reference() {
        let directory = SlideDirectory::new_for_test(0);
        let index = NotesIndex::try_build(&document_with_notes(&[(9, 0x345)]), &directory)
            .expect("valid notes index");
        let mut slide_atom_data = [0u8; SLIDE_ATOM_SIZE];
        slide_atom_data[16..20].copy_from_slice(&0x345u32.to_le_bytes());
        let slide_atom = record(2, 0, 0x03ef, &slide_atom_data);
        let slide_bytes = record(0x0f, 0, 0x03ee, &slide_atom);
        let (slide, _) = Record::parse(&slide_bytes, 0).expect("parse synthetic slide");
        let error = index
            .descriptor(&slide, 4, &PersistMapping::new())
            .expect_err("dangling persist reference must fail");
        assert!(error.contains("no directory entry"));
    }

    #[test]
    fn poi_basic_fixture_exposes_exact_speaker_notes() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/poi/test-data/slideshow/basic_test_ppt_file.ppt"
        );
        let mut package = Package::open(path).expect("open POI fixture");
        let presentation = package.presentation().expect("parse POI fixture");
        let slides = presentation.slides().expect("read slides");
        let mut notes = Vec::new();
        for slide in &slides {
            if let Some(value) = slide.speaker_notes().expect("read speaker notes") {
                notes.push(value.text().expect("extract notes text").to_string());
            }
        }
        assert_eq!(
            notes,
            [
                "These are the notes for page 1",
                "These are the notes on page two, again lacking formatting",
            ]
        );
    }

    #[test]
    fn writer_notes_round_trip_through_reader_api() {
        let mut writer = crate::writer::Writer::new();
        let slide = writer.add_slide().expect("add slide");
        let notes_page = crate::writer::NotesPage::simple(0, "reader writer notes round trip")
            .with_header("excluded notes header")
            .with_footer("excluded notes footer");
        writer.set_notes_page(slide, notes_page).expect("set notes");
        let path = std::env::temp_dir().join(format!(
            "litchi-ppt-notes-{}-{}.ppt",
            std::process::id(),
            slide
        ));
        writer.save(&path).expect("write presentation");

        let mut package = Package::open(&path).expect("open written presentation");
        let presentation = package.presentation().expect("parse written presentation");
        let slides = presentation.slides().expect("read written slides");
        let notes = slides[0]
            .speaker_notes()
            .expect("read written notes")
            .expect("notes exist");
        assert_eq!(
            notes.text().expect("extract written notes"),
            "reader writer notes round trip"
        );
        let _removed = std::fs::remove_file(path);
    }

    #[test]
    fn parsed_notes_shape_clone_refuses_public_mutation() {
        use crate::shapes::{Mutation, MutationError, Shape, ShapeEnum};

        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/poi/test-data/slideshow/basic_test_ppt_file.ppt"
        );
        let package = Package::open(path);
        assert!(package.is_ok());
        let Ok(mut opened_package) = package else {
            return;
        };
        let presentation = opened_package.presentation();
        assert!(presentation.is_ok());
        let Ok(parsed_presentation) = presentation else {
            return;
        };
        let slides = parsed_presentation.slides();
        assert!(slides.is_ok());
        let Ok(parsed_slides) = slides else {
            return;
        };

        let mut candidate = None;
        'slides: for slide in &parsed_slides {
            let notes = slide.speaker_notes();
            assert!(notes.is_ok());
            let Ok(Some(parsed_notes)) = notes else {
                continue;
            };
            let shapes = parsed_notes.shapes();
            assert!(shapes.is_ok());
            let Ok(notes_shapes) = shapes else {
                continue;
            };
            for shape in notes_shapes {
                if !matches!(
                    shape,
                    ShapeEnum::TextBox(_) | ShapeEnum::Placeholder(_) | ShapeEnum::AutoShape(_)
                ) {
                    continue;
                }
                let text = shape.text();
                assert!(text.is_ok());
                let Ok(shape_text) = text else {
                    continue;
                };
                if !shape_text.is_empty() {
                    candidate = Some(shape.clone());
                    break 'slides;
                }
            }
        }

        assert!(candidate.is_some());
        let Some(mut shape) = candidate else {
            return;
        };
        let before_text = shape.text();
        assert!(before_text.is_ok());
        let Ok(text_before_mutation) = before_text else {
            return;
        };
        let before_fill = match &shape {
            ShapeEnum::TextBox(text_box) => text_box.properties().fill_color,
            ShapeEnum::Placeholder(placeholder) => placeholder.properties().fill_color,
            ShapeEnum::AutoShape(auto_shape) => auto_shape.properties().fill_color,
            ShapeEnum::Picture(_)
            | ShapeEnum::Table(_)
            | ShapeEnum::Group(_)
            | ShapeEnum::Line(_) => {
                return;
            },
        };
        let mutation = match &mut shape {
            ShapeEnum::TextBox(text_box) => Shape::set_fill_color(text_box, Some(0x0012_3456)),
            ShapeEnum::Placeholder(placeholder) => {
                Shape::set_fill_color(placeholder, Some(0x0012_3456))
            },
            ShapeEnum::AutoShape(auto_shape) => {
                Shape::set_fill_color(auto_shape, Some(0x0012_3456))
            },
            ShapeEnum::Picture(_)
            | ShapeEnum::Table(_)
            | ShapeEnum::Group(_)
            | ShapeEnum::Line(_) => {
                return;
            },
        };

        assert_eq!(
            mutation,
            Err(MutationError::SourceBound {
                mutation: Mutation::Fill,
            })
        );
        let after_text = shape.text();
        assert!(after_text.is_ok());
        let Ok(text_after_mutation) = after_text else {
            return;
        };
        assert_eq!(text_after_mutation, text_before_mutation);
        let after_fill = match &shape {
            ShapeEnum::TextBox(text_box) => text_box.properties().fill_color,
            ShapeEnum::Placeholder(placeholder) => placeholder.properties().fill_color,
            ShapeEnum::AutoShape(auto_shape) => auto_shape.properties().fill_color,
            ShapeEnum::Picture(_)
            | ShapeEnum::Table(_)
            | ShapeEnum::Group(_)
            | ShapeEnum::Line(_) => {
                return;
            },
        };
        assert_eq!(after_fill, before_fill);
    }
}
