use super::core::Field;

/// A typed, inert legacy Word `MACROBUTTON` field.
///
/// ECMA-376 Part 1 §17.16.5.34 defines two stored field arguments: a macro or
/// command name and the text or graphic used as its button.
///
/// This preserves the stored macro or command name, button text, cached
/// result, and field-marker state. It never resolves, loads, invokes, or
/// otherwise executes the named macro or command.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MacroButtonField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) macro_name: String,
    pub(in crate::parts::fields) display_text: String,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl MacroButtonField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored macro or command name without resolving or invoking it.
    pub fn macro_name(&self) -> &str {
        &self.macro_name
    }

    /// Return the stored button text.
    ///
    /// This is source metadata, not a generated result.
    pub fn display_text(&self) -> &str {
        &self.display_text
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from the macro or command.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// A typed, inert legacy Word `GOTOBUTTON` field.
///
/// [MS-DOC] §2.9.90 identifies its native field-type byte, and ECMA-376 Part
/// 1 §17.16.5.23 defines two stored field arguments: a destination and the
/// text or graphic used as its button. This type exposes stored text only; it
/// never resolves a destination, changes the insertion point, or activates a
/// jump.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GoToButtonField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) target: String,
    pub(in crate::parts::fields) button_text: String,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl GoToButtonField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored destination without resolving or navigating to it.
    ///
    /// A destination can be a bookmark, page reference, annotation, footnote,
    /// line, page, or section expression.
    pub fn target(&self) -> &str {
        &self.target
    }

    /// Return the stored text or graphic-label expression for the button.
    ///
    /// This is source metadata, not an activated control.
    pub fn button_text(&self) -> &str {
        &self.button_text
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from the destination.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The stored category of a legacy Word active-content field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ActiveContentFieldKind {
    /// An `ADDIN` field that stores add-in-created data.
    AddIn,
    /// A `CONTROL` field that represents an OCX control.
    OcxControl,
    /// An `HTMLCONTROL` field that represents an HTML control.
    HtmlControl,
}

/// Typed, inert metadata for a legacy Word add-in or control field.
///
/// [MS-DOC] §2.9.90 identifies the native `ADDIN`, `CONTROL`, and
/// `HTMLCONTROL` field types. This type retains only the stored category,
/// instruction, cached result, and field state. It never loads an add-in,
/// instantiates an OCX or HTML control, invokes code, executes script, renders
/// a control, or accesses an external resource.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ActiveContentField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) kind: ActiveContentFieldKind,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl ActiveContentField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never interpreted.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this stores add-in, OCX-control, or HTML-control metadata.
    pub fn kind(&self) -> ActiveContentFieldKind {
        self.kind
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by loading or running content.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `PRINT` field.
///
/// [MS-DOC] §2.9.90 identifies native `PRINT` fields with type `0x30`.
/// This type retains opaque printer-instruction text, a cached result, and
/// field-marker state only. It never interprets printer-control codes, opens a
/// printer, sends output, changes print settings, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PrintField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) printer_instructions: String,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl PrintField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never interpreted or sent to
    /// a printer.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored printer-instruction text after the `PRINT` keyword.
    ///
    /// This can include printer-control or PostScript text. It is never parsed,
    /// interpreted, or sent to a printer.
    pub fn printer_instructions(&self) -> &str {
        &self.printer_instructions
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by printing.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `EMBED` field.
///
/// [MS-DOC] §2.9.90 identifies native `EMBED` fields with type `0x3A`.
/// This type retains opaque object-instruction text, a cached result, and
/// field-marker state only. It never loads, inspects, deserializes, activates,
/// renders, or executes an embedded object, accesses an external resource, or
/// refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbedField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) object_instructions: String,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl EmbedField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `EMBED` field instruction.
    ///
    /// This string remains opaque metadata and is never used to load or
    /// activate an object.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored opaque object-instruction text after `EMBED`.
    ///
    /// It is never parsed, used to locate an object, or used to load, inspect,
    /// deserialize, activate, render, or execute object content.
    pub fn object_instructions(&self) -> &str {
        &self.object_instructions
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is cached text only and is never regenerated from an object.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `BARCODE` field.
///
/// [MS-DOC] §2.9.90 identifies native `BARCODE` fields with type `0x3F`.
/// This type retains opaque barcode-instruction text, a cached result, and
/// field-marker state only. It never parses or validates barcode data or
/// symbology, generates or renders a barcode, accesses an external resource,
/// or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BarcodeField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) barcode_instructions: String,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl BarcodeField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `BARCODE` field instruction.
    ///
    /// This string remains opaque metadata and is never used to generate or
    /// render a barcode.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored opaque barcode-instruction text after `BARCODE`.
    ///
    /// It is never parsed, validated, interpreted, or used to generate or
    /// render barcode content.
    pub fn barcode_instructions(&self) -> &str {
        &self.barcode_instructions
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is cached text only and is never regenerated from barcode
    /// data.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `BIDIOUTLINE` field.
///
/// [MS-DOC] §2.9.90 identifies native `BIDIOUTLINE` fields with type
/// `0x5C`. This type retains opaque instruction text, a cached result, and
/// field-marker state only. It never reads right-to-left language, paragraph
/// outline, or layout state; chooses a numbering system; calculates a result;
/// or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BidiOutlineField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) opaque_instructions: String,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl BidiOutlineField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `BIDIOUTLINE` field instruction.
    ///
    /// This string remains opaque metadata and is never used to calculate an
    /// outline number.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return opaque stored instruction text after `BIDIOUTLINE`.
    ///
    /// It is never parsed, interpreted, or used to resolve language, outline,
    /// numbering, or layout state.
    pub fn opaque_instructions(&self) -> &str {
        &self.opaque_instructions
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is cached text only and is never regenerated from document
    /// state.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `SHAPE` field.
///
/// [MS-DOC] §2.9.90 identifies native `SHAPE` fields with type `0x5F`.
/// Word uses this legacy field as a drawing-canvas anchor. This type retains
/// opaque instruction text, a cached result, and field-marker state only. It
/// never locates, links, loads, positions, lays out, or renders a drawing or
/// canvas, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) opaque_instructions: String,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl ShapeField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `SHAPE` field instruction.
    ///
    /// This string remains opaque metadata and is never used to locate or
    /// position a drawing canvas.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return opaque stored instruction text after `SHAPE`.
    ///
    /// It is never parsed, interpreted, or used to link a field to a drawing,
    /// resolve an anchor, or calculate layout.
    pub fn opaque_instructions(&self) -> &str {
        &self.opaque_instructions
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is cached metadata only and is never regenerated from a
    /// drawing canvas.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The stored kind of a legacy Word form-code field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyFormFieldKind {
    /// A `FORMTEXT` text-box form field.
    Text,
    /// A `FORMCHECKBOX` checkbox form field.
    CheckBox,
    /// A `FORMDROPDOWN` drop-down-list form field.
    DropDown,
}

/// Typed, inert metadata for a legacy Word form-code field.
///
/// [MS-DOC] §2.9.90 identifies native `FORMTEXT`, `FORMCHECKBOX`, and
/// `FORMDROPDOWN` fields with types `0x46`, `0x47`, and `0x53`. This type
/// retains only the stored kind, opaque instruction text, cached result, and
/// field-marker state, plus the stored `FFData` form state when it could be
/// located and parsed. It never fills a form, changes a selection or checkbox
/// state, invokes entry or exit macros, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyFormField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) kind: LegacyFormFieldKind,
    pub(in crate::parts::fields) opaque_instructions: String,
    pub(in crate::parts::fields) cached_result: Option<String>,
    pub(in crate::parts::fields) form_data: Option<crate::parts::form_fields::FormFieldData>,
}

impl LegacyFormField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never used to change a form
    /// field.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is a text, checkbox, or drop-down form-code field.
    pub const fn kind(&self) -> LegacyFormFieldKind {
        self.kind
    }

    /// Return opaque stored instruction text after the form-code keyword.
    ///
    /// It is never parsed, interpreted, or used to fill a form, change a
    /// checkbox or selection, or invoke a macro.
    pub fn opaque_instructions(&self) -> &str {
        &self.opaque_instructions
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is cached metadata only and is never regenerated from form
    /// state.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Return the parsed stored form state (`FFData`, MS-DOC 2.9.78), when the
    /// field's `NilPICFAndBinData` could be located in the Data stream and was
    /// well-formed.
    ///
    /// The returned data is inert: entry and exit macro names are stored
    /// verbatim and never invoked, the form is never filled, and checkbox or
    /// selection state is never changed. Fields constructed without Data
    /// stream access (or whose stored binary data is invalid, which MS-DOC
    /// §2.9.158 says MUST be ignored) return `None`.
    pub fn form_data(&self) -> Option<&crate::parts::form_fields::FormFieldData> {
        self.form_data.as_ref()
    }

    /// Attach the parsed stored form state. Crate-internal: only the document
    /// layer can locate the `NilPICFAndBinData` in the Data stream.
    pub(crate) fn set_form_data(
        &mut self,
        form_data: Option<crate::parts::form_fields::FormFieldData>,
    ) {
        self.form_data = form_data;
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}
