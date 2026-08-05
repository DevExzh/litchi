//! Hyperlink and bookmark-reference field models.

use super::{Field, Switch};

use crate::error::{Error, Result};

use super::super::codec::{field_instruction_remainder, parse_field_operand_and_switches};

use super::super::{MAX_HYPERLINK_FIELD_INSTRUCTION_BYTES, MAX_REFERENCE_FIELD_INSTRUCTION_BYTES};

/// A typed, inert Word `HYPERLINK` field.
///
/// This type retains only stored link metadata, a cached result, and field
/// state. It never opens, resolves, follows, activates, or refreshes a link.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Hyperlink {
    instruction: String,
    external_target: Option<String>,
    bookmark: Option<String>,
    screen_tip: Option<String>,
    target_frame: Option<String>,
    appends_image_map_coordinates: bool,
    opens_new_window: bool,
    unknown_switches: Vec<Switch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Hyperlink {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        if !field.is_hyperlink_field() {
            return Ok(None);
        }
        if field.instruction().len() > MAX_HYPERLINK_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "HYPERLINK field instruction exceeds {MAX_HYPERLINK_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }
        let Some((external_target, switches)) =
            parse_field_operand_and_switches(field.instruction(), "HYPERLINK")?
        else {
            unreachable!("hyperlink-field recognition and parsing must agree");
        };
        let external_target = external_target
            .map(|target| {
                (!target.is_empty()).then_some(target).ok_or_else(|| {
                    Error::Invalid("HYPERLINK external target must not be empty".to_string())
                })
            })
            .transpose()?;

        let mut bookmark = None;
        let mut screen_tip = None;
        let mut target_frame = None;
        let mut appends_image_map_coordinates = false;
        let mut opens_new_window = false;
        let mut unknown_switches = Vec::new();
        for switch in switches {
            let (slot, switch_name) = match switch.name {
                'l' => (&mut bookmark, 'l'),
                'o' => (&mut screen_tip, 'o'),
                't' => (&mut target_frame, 't'),
                'm' => {
                    if appends_image_map_coordinates {
                        return Err(Error::Invalid(
                            "HYPERLINK \\m switch is duplicated".to_string(),
                        ));
                    }
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "HYPERLINK \\m switch does not take an argument".to_string(),
                        ));
                    }
                    appends_image_map_coordinates = true;
                    continue;
                },
                'n' => {
                    if opens_new_window {
                        return Err(Error::Invalid(
                            "HYPERLINK field has duplicate \\n switches".to_string(),
                        ));
                    }
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "HYPERLINK \\n switch does not take an argument".to_string(),
                        ));
                    }
                    opens_new_window = true;
                    continue;
                },
                _ => {
                    unknown_switches.push(switch);
                    continue;
                },
            };
            let value = switch.argument.ok_or_else(|| {
                Error::Invalid(format!(
                    "HYPERLINK \\{switch_name} switch requires an argument"
                ))
            })?;
            if value.is_empty() {
                return Err(Error::Invalid(format!(
                    "HYPERLINK \\{switch_name} switch argument must not be empty"
                )));
            }
            if slot.replace(value).is_some() {
                return Err(Error::Invalid(format!(
                    "HYPERLINK field has duplicate \\{switch_name} switches"
                )));
            }
        }
        if external_target.is_none() && bookmark.is_none() {
            return Err(Error::Invalid(
                "HYPERLINK field requires an external target or \\l bookmark".to_string(),
            ));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            external_target,
            bookmark,
            screen_tip,
            target_frame,
            appends_image_map_coordinates,
            opens_new_window,
            unknown_switches,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored external target without resolving or opening it.
    pub fn external_target(&self) -> Option<&str> {
        self.external_target.as_deref()
    }

    /// Return the stored internal bookmark target without resolving it.
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Return the stored screen-tip text, if present.
    ///
    /// This is metadata only and is never displayed by the library.
    pub fn screen_tip(&self) -> Option<&str> {
        self.screen_tip.as_deref()
    }

    /// Return the stored target frame, if present.
    ///
    /// This is metadata only and is never used to open a window or frame.
    pub fn target_frame(&self) -> Option<&str> {
        self.target_frame.as_deref()
    }

    /// Whether the target receives click coordinates for a server-side image map.
    ///
    /// This records producer intent only; no navigation or hit testing occurs.
    pub fn appends_image_map_coordinates(&self) -> bool {
        self.appends_image_map_coordinates
    }

    /// Whether the field requests opening the target in a new window.
    ///
    /// This records producer intent only; no window is opened.
    pub fn opens_new_window(&self) -> bool {
        self.opens_new_window
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[Switch] {
        &self.unknown_switches
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by resolving a link.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }
}

/// The stored category of a Word bookmark-reference field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReferenceKind {
    /// A `REF` field.
    Reference,
    /// A `PAGEREF` field.
    PageReference,
    /// A historical `FTNREF` field.
    FootnoteReference,
    /// A `NOTEREF` field.
    NoteReference,
}

impl ReferenceKind {
    fn from_instruction(instruction: &str) -> Option<(Self, &'static str)> {
        for (kind, field_type) in [
            (Self::Reference, "REF"),
            (Self::PageReference, "PAGEREF"),
            (Self::FootnoteReference, "FTNREF"),
            (Self::NoteReference, "NOTEREF"),
        ] {
            if field_instruction_remainder(instruction, field_type).is_some() {
                return Some((kind, field_type));
            }
        }
        None
    }

    fn is_note_reference(self) -> bool {
        matches!(self, Self::FootnoteReference | Self::NoteReference)
    }
}

/// One recognized stored option of a Word bookmark-reference field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ReferenceOption {
    /// The `\d` `REF` separator between sequence and page numbers.
    SequencePageSeparator(String),
    /// The `\f` `REF` request for referenced note or comment content.
    ReferencedNoteContent,
    /// The `\h` request for a link to the stored bookmark.
    Hyperlink,
    /// The `\n` `REF` request for a paragraph number without context.
    ParagraphNumberWithoutContext,
    /// The `\p` request for relative-position text.
    RelativePosition,
    /// The `\r` `REF` request for a paragraph number in relative context.
    ParagraphNumberRelativeContext,
    /// The `\t` `REF` request to suppress non-number text.
    SuppressNonNumberText,
    /// The `\w` `REF` request for a paragraph number in full context.
    ParagraphNumberFullContext,
    /// The `\f` `FTNREF` or `NOTEREF` request to format the note mark.
    NoteMarkFormatting,
}

/// A typed, inert Word bookmark-reference field.
///
/// This model preserves only stored categories, targets, options, switches,
/// cached results, and field state. It never looks up a bookmark, reads a
/// referenced range or note, resolves a page number, creates a link,
/// calculates a relative position, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    instruction: String,
    kind: ReferenceKind,
    bookmark: String,
    options: Vec<ReferenceOption>,
    unknown_switches: Vec<Switch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Reference {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((kind, field_type)) = ReferenceKind::from_instruction(field.instruction()) else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_REFERENCE_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "{field_type} field instruction exceeds {MAX_REFERENCE_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }
        let Some((bookmark, switches)) =
            parse_field_operand_and_switches(field.instruction(), field_type)?
        else {
            unreachable!("bookmark-reference recognition and parsing must agree");
        };
        let bookmark = bookmark
            .filter(|bookmark| !bookmark.is_empty())
            .ok_or_else(|| {
                Error::Invalid(format!("{field_type} field is missing its bookmark target"))
            })?;

        let mut options = Vec::new();
        let mut unknown_switches = Vec::new();
        for switch in switches {
            match switch.name {
                'd' if kind == ReferenceKind::Reference => {
                    let separator = switch.argument.ok_or_else(|| {
                        Error::Invalid("REF \\d switch requires a separator".to_string())
                    })?;
                    options.push(ReferenceOption::SequencePageSeparator(separator));
                },
                'f' if kind == ReferenceKind::Reference => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "REF \\f switch does not take an argument".to_string(),
                        ));
                    }
                    options.push(ReferenceOption::ReferencedNoteContent);
                },
                'f' if kind.is_note_reference() => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(format!(
                            "{field_type} \\f switch does not take an argument"
                        )));
                    }
                    options.push(ReferenceOption::NoteMarkFormatting);
                },
                'h' => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(format!(
                            "{field_type} \\h switch does not take an argument"
                        )));
                    }
                    options.push(ReferenceOption::Hyperlink);
                },
                'n' if kind == ReferenceKind::Reference => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "REF \\n switch does not take an argument".to_string(),
                        ));
                    }
                    options.push(ReferenceOption::ParagraphNumberWithoutContext);
                },
                'p' => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(format!(
                            "{field_type} \\p switch does not take an argument"
                        )));
                    }
                    options.push(ReferenceOption::RelativePosition);
                },
                'r' if kind == ReferenceKind::Reference => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "REF \\r switch does not take an argument".to_string(),
                        ));
                    }
                    options.push(ReferenceOption::ParagraphNumberRelativeContext);
                },
                't' if kind == ReferenceKind::Reference => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "REF \\t switch does not take an argument".to_string(),
                        ));
                    }
                    options.push(ReferenceOption::SuppressNonNumberText);
                },
                'w' if kind == ReferenceKind::Reference => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "REF \\w switch does not take an argument".to_string(),
                        ));
                    }
                    options.push(ReferenceOption::ParagraphNumberFullContext);
                },
                _ => unknown_switches.push(switch),
            }
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            kind,
            bookmark,
            options,
            unknown_switches,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored reference-field category.
    pub const fn kind(&self) -> ReferenceKind {
        self.kind
    }

    /// Return the stored bookmark or note target without resolving it.
    pub fn bookmark(&self) -> &str {
        &self.bookmark
    }

    /// Return recognized stored options in source order.
    ///
    /// This metadata is never used to navigate, resolve, or activate a link.
    pub fn options(&self) -> &[ReferenceOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[Switch] {
        &self.unknown_switches
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by resolving a
    /// bookmark, page number, or note reference.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }
}

impl Field {
    /// Check whether this is a bookmark-reference field.
    ///
    /// This recognizes `REF`, `PAGEREF`, `FTNREF`, and `NOTEREF` stored
    /// instructions. It never looks up a bookmark, reads a referenced range or
    /// note, resolves a page number, creates a link, calculates a relative
    /// position, or refreshes the result.
    pub fn is_reference_field(&self) -> bool {
        ReferenceKind::from_instruction(&self.instruction).is_some()
    }

    /// Parse this field as inert bookmark-reference metadata.
    ///
    /// Returns `Ok(None)` for fields other than `REF`, `PAGEREF`, `FTNREF`, and
    /// `NOTEREF`. The stored kind, target, options, unknown switches, cached
    /// content, and dirty/lock state are metadata only; this method never looks
    /// up a bookmark, reads a referenced range or note, resolves a page number,
    /// creates a link, calculates a relative position, or refreshes a field.
    pub fn reference_field(&self) -> Result<Option<Reference>> {
        Reference::from_field(self)
    }

    /// Check whether this is a `HYPERLINK` field.
    ///
    /// Recognition is limited to stored field metadata. It never opens,
    /// resolves, follows, or refreshes a hyperlink target.
    pub fn is_hyperlink_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "HYPERLINK").is_some()
    }

    /// Parse this field as inert `HYPERLINK` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `HYPERLINK`. The stored target,
    /// bookmark, tooltip, frame, image-map-coordinate request, switches, cached
    /// content, and dirty/lock state are metadata only; this method never opens, resolves,
    /// follows, activates, or refreshes a link.
    pub fn hyperlink_field(&self) -> Result<Option<Hyperlink>> {
        Hyperlink::from_field(self)
    }
}
