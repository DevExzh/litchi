//! XML and typed-part codecs exposed through the document facade.

use crate::content_control::ContentControl;
use crate::error::{Error, Result};
use crate::field::Compare;
use crate::field::{
    ActiveContent, Advance, AutoNumber, AutoText, AutoTextList, Barcode, Bibliography, BidiOutline,
    Citation, Context, Database, Dde, Embed, Equation, Field, Formula, GoToButton, Hyperlink, If,
    Include, Index, IndexEntry, Info, Information, LegacyForm, Link, ListNumber, MacroButton,
    Merge, MergeControl, MergeCounter, MergeData, MergeNext, Print, Private, Prompt, Property,
    Quote, Recipient, Reference, Sequence, Set, Shape, StyleReference, SubDocument, Symbol, Toa,
    ToaEntry, Toc, TocEntry, UserIdentity, Variable,
};
use crate::section::Sections;
use crate::styles::Styles;
use litchi_opc::constants::relationship_type;

use super::super::codec;
use super::super::model::Document;

impl<'a> Document<'a> {
    /// Get all sections in the document.
    ///
    /// Returns a `Sections` collection providing access to each section's
    /// page properties, margins, orientation, etc.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let mut sections = doc.sections()?;
    ///
    /// println!("Document has {} sections", sections.len());
    /// for section in sections.iter_mut() {
    ///     println!("Orientation: {}", section.orientation());
    ///     if let Some(width) = section.page_width() {
    ///         println!("  Page width: {} inches", width.to_inches());
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn sections(&self) -> Result<Sections> {
        codec::extract_sections(self.part.xml_bytes())
    }

    /// Get the document styles.
    ///
    /// Returns a `Styles` object providing access to all paragraph, character,
    /// table, and list styles defined in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let mut styles = doc.styles()?;
    ///
    /// // Find a style by name
    /// if let Some(style) = styles.get_by_name("Heading 1")? {
    ///     println!("Found style: {} (id: {})",
    ///         style.name().unwrap_or(""),
    ///         style.style_id());
    /// }
    ///
    /// // Iterate all styles
    /// for style in styles.iter()? {
    ///     println!("Style: {} - Type: {}",
    ///         style.name().unwrap_or("<unnamed>"),
    ///         style.style_type());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn styles(&self) -> Result<Styles<'a>> {
        // Try to find the styles part through the main document part's relationships
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        // Look for a relationship to the styles part
        if let Ok(rel) = rels.part_with_reltype(relationship_type::STYLES) {
            let target = rel.target_partname()?;
            let styles_part = self.opc.get_part(&target)?;
            return Ok(Styles::from_part(styles_part));
        }

        // If no styles part is found, return an empty Styles object
        // This can happen in minimal documents
        Err(Error::PartNotFound("styles part not found".to_string()))
    }

    pub fn fields(&self) -> Result<Vec<Field>> {
        let xml_bytes = self.part.xml_bytes();
        Ok(Field::extract_from_document(xml_bytes)?)
    }

    /// Get the number of fields in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// println!("Document has {} fields", doc.field_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn field_count(&self) -> Result<usize> {
        Ok(self.fields()?.len())
    }

    /// Get typed, inert `HYPERLINK` fields in document order.
    ///
    /// Returned values expose stored targets, bookmarks, display metadata,
    /// switches, cached content, and dirty/lock state only. This method never
    /// opens, resolves, follows, activates, or refreshes a link.
    pub fn hyperlink_fields(&self) -> Result<Vec<Hyperlink>> {
        self.fields()?
            .iter()
            .map(|field| Field::hyperlink_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `HYPERLINK` fields in the main document.
    pub fn hyperlink_field_count(&self) -> Result<usize> {
        Ok(self.hyperlink_fields()?.len())
    }

    /// Get typed, inert bibliography citation (`CITATION`) fields in document order.
    ///
    /// Returned values expose stored source tags, switches, cached content, and
    /// dirty/lock state only. This method never looks up bibliography sources,
    /// formats citations, or refreshes fields.
    pub fn citations(&self) -> Result<Vec<Citation>> {
        self.fields()?
            .iter()
            .map(|field| Field::citation(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of bibliography citation fields in the main document.
    pub fn citation_count(&self) -> Result<usize> {
        Ok(self.citations()?.len())
    }

    /// Get typed, inert `BIBLIOGRAPHY` fields in document order.
    ///
    /// Returned values expose stored switches and cached content only. This
    /// method never loads source XML, sorts sources, or regenerates a
    /// bibliography.
    pub fn bibliographies(&self) -> Result<Vec<Bibliography>> {
        self.fields()?
            .iter()
            .map(|field| Field::bibliography(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of bibliography fields in the main document.
    pub fn bibliography_count(&self) -> Result<usize> {
        Ok(self.bibliographies()?.len())
    }

    /// Get typed, inert `DOCVARIABLE` fields in document order.
    ///
    /// Returned values expose stored names, switches, cached content, and
    /// dirty/lock state only. This method never reads the settings part,
    /// resolves document-variable values, or refreshes fields.
    pub fn document_variable_fields(&self) -> Result<Vec<Variable>> {
        self.fields()?
            .iter()
            .map(|field| Field::document_variable(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of document-variable fields in the main document.
    pub fn document_variable_field_count(&self) -> Result<usize> {
        Ok(self.document_variable_fields()?.len())
    }

    /// Get typed, inert `DOCPROPERTY` fields in document order.
    ///
    /// Returned values expose stored property names, switches, cached content,
    /// and dirty/lock state only. This method never reads core, extended, or
    /// custom package properties, resolves a value, or refreshes fields.
    pub fn document_property_fields(&self) -> Result<Vec<Property>> {
        self.fields()?
            .iter()
            .map(|field| Field::document_property(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `DOCPROPERTY` fields in the main document.
    pub fn document_property_field_count(&self) -> Result<usize> {
        Ok(self.document_property_fields()?.len())
    }

    /// Get typed, inert explicit legacy `INFO` fields in document order.
    ///
    /// Returned values expose stored property selectors, optional replacement
    /// values, switches, cached content, and dirty/lock state only. This method
    /// never reads, resolves, modifies, or writes document or template
    /// properties, or refreshes a field.
    pub fn info_fields(&self) -> Result<Vec<Info>> {
        self.fields()?
            .iter()
            .map(|field| Field::info_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert explicit legacy `INFO` fields.
    pub fn info_field_count(&self) -> Result<usize> {
        Ok(self.info_fields()?.len())
    }

    /// Get typed, inert built-in document-information fields in document order.
    ///
    /// Returned values expose only stored kinds, switches, cached content, and
    /// dirty/lock state. This method never reads package metadata or host
    /// identity data, calculates dates, revisions, or statistics, resolves a
    /// value, or refreshes fields.
    pub fn document_information_fields(&self) -> Result<Vec<Information>> {
        self.fields()?
            .iter()
            .map(|field| Field::document_information(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert built-in document-information fields.
    pub fn document_information_field_count(&self) -> Result<usize> {
        Ok(self.document_information_fields()?.len())
    }

    /// Get typed, inert built-in document-context and runtime fields in document order.
    ///
    /// Returned values expose only stored kinds, switches, cached content, and
    /// dirty/lock state. This method never reads a document path, attached
    /// template, host filesystem state or file size, current clock, or page and
    /// section layout; resolves a value; or refreshes fields.
    pub fn document_context_fields(&self) -> Result<Vec<Context>> {
        self.fields()?
            .iter()
            .map(|field| Field::document_context(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert built-in document-context and runtime fields.
    pub fn document_context_field_count(&self) -> Result<usize> {
        Ok(self.document_context_fields()?.len())
    }

    /// Get typed, inert `MACROBUTTON` fields in document order.
    ///
    /// Returned values expose only stored macro or command names, button text,
    /// cached results, and dirty/lock state. This method never resolves, loads,
    /// invokes, or otherwise executes a macro or command.
    pub fn macro_button_fields(&self) -> Result<Vec<MacroButton>> {
        self.fields()?
            .iter()
            .map(|field| Field::macro_button(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `MACROBUTTON` fields in the main document.
    pub fn macro_button_field_count(&self) -> Result<usize> {
        Ok(self.macro_button_fields()?.len())
    }

    /// Get typed, inert `ADDIN`, `CONTROL`, and `HTMLCONTROL` fields in document
    /// order.
    ///
    /// Returned values expose stored kinds, instructions, cached content, and
    /// dirty/lock state only. This method never loads an add-in, instantiates
    /// an OCX or HTML control, invokes code, executes script, renders content,
    /// accesses an external resource, or refreshes a field.
    pub fn active_content_fields(&self) -> Result<Vec<ActiveContent>> {
        self.fields()?
            .iter()
            .map(|field| Field::active_content_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert active-content fields in the main document.
    pub fn active_content_field_count(&self) -> Result<usize> {
        Ok(self.active_content_fields()?.len())
    }

    /// Get typed, inert `GLOSSARY` and `AUTOTEXT` fields in document order.
    ///
    /// Returned values expose stored kinds, entry names, switches, cached
    /// content, and dirty/lock state only. This method never looks up a
    /// building block, reads a template, inserts content, changes bookmarks,
    /// accesses an external resource, or refreshes a field.
    pub fn auto_text_fields(&self) -> Result<Vec<AutoText>> {
        self.fields()?
            .iter()
            .map(|field| Field::auto_text_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert building-block fields in the main document.
    pub fn auto_text_field_count(&self) -> Result<usize> {
        Ok(self.auto_text_fields()?.len())
    }

    /// Get typed, inert `AUTOTEXTLIST` fields in document order.
    ///
    /// Returned values expose stored display text, style/tip options, unknown
    /// switches, cached content, and dirty/lock state only. This method never
    /// shows a selection UI, looks up a building block, reads a template,
    /// inserts content, accesses an external resource, or refreshes a field.
    pub fn auto_text_list_fields(&self) -> Result<Vec<AutoTextList>> {
        self.fields()?
            .iter()
            .map(|field| Field::auto_text_list_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `AUTOTEXTLIST` fields in the main document.
    pub fn auto_text_list_field_count(&self) -> Result<usize> {
        Ok(self.auto_text_list_fields()?.len())
    }

    /// Get typed, inert `GOTOBUTTON` fields in document order.
    ///
    /// Returned values expose only stored destinations, button text, cached
    /// results, and dirty/lock state. This method never resolves a destination,
    /// changes the insertion point, activates a jump, or refreshes a field.
    pub fn go_to_button_fields(&self) -> Result<Vec<GoToButton>> {
        self.fields()?
            .iter()
            .map(|field| Field::go_to_button(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `GOTOBUTTON` fields in the main document.
    pub fn go_to_button_field_count(&self) -> Result<usize> {
        Ok(self.go_to_button_fields()?.len())
    }

    /// Get typed, inert `PRINT` fields in document order.
    ///
    /// Returned values expose only stored printer-instruction text, cached
    /// results, and dirty/lock state. This method never interprets control
    /// codes, opens a printer, sends output, changes print settings, or
    /// refreshes a field.
    pub fn print_fields(&self) -> Result<Vec<Print>> {
        self.fields()?
            .iter()
            .map(|field| Field::print_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `PRINT` fields in the main document.
    pub fn print_field_count(&self) -> Result<usize> {
        Ok(self.print_fields()?.len())
    }

    /// Get typed, inert `EMBED` fields in document order.
    ///
    /// Returned values expose only stored opaque object instructions, cached
    /// content, and dirty/lock state. This method never loads, inspects,
    /// deserializes, activates, renders, or executes an embedded object,
    /// accesses an external resource, or refreshes a field.
    pub fn embed_fields(&self) -> Result<Vec<Embed>> {
        self.fields()?
            .iter()
            .map(|field| Field::embed_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `EMBED` fields in the main document.
    pub fn embed_field_count(&self) -> Result<usize> {
        Ok(self.embed_fields()?.len())
    }

    /// Get typed, inert `BARCODE` fields in document order.
    ///
    /// Returned values expose only stored opaque barcode instructions, cached
    /// content, and dirty/lock state. This method never parses or validates
    /// barcode data or symbology, generates or renders a barcode, accesses an
    /// external resource, or refreshes a field.
    pub fn barcode_fields(&self) -> Result<Vec<Barcode>> {
        self.fields()?
            .iter()
            .map(|field| Field::barcode_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `BARCODE` fields in the main document.
    pub fn barcode_field_count(&self) -> Result<usize> {
        Ok(self.barcode_fields()?.len())
    }

    /// Get typed, inert `BIDIOUTLINE` fields in document order.
    ///
    /// Returned values expose only stored opaque instructions, cached content,
    /// and dirty/lock state. This method never reads right-to-left language,
    /// paragraph outline, or layout state; chooses a numbering system;
    /// calculates a result; or refreshes a field.
    pub fn bidi_outline_fields(&self) -> Result<Vec<BidiOutline>> {
        self.fields()?
            .iter()
            .map(|field| Field::bidi_outline_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `BIDIOUTLINE` fields in the main document.
    pub fn bidi_outline_field_count(&self) -> Result<usize> {
        Ok(self.bidi_outline_fields()?.len())
    }

    /// Get typed, inert `SHAPE` drawing-canvas anchor fields in document order.
    ///
    /// Returned values expose only stored opaque instructions, cached content,
    /// and dirty/lock state. This method never locates, links, loads, positions,
    /// lays out, or renders a drawing or canvas, or refreshes a field.
    pub fn shape_fields(&self) -> Result<Vec<Shape>> {
        self.fields()?
            .iter()
            .map(|field| Field::shape_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `SHAPE` drawing-canvas anchor fields.
    pub fn shape_field_count(&self) -> Result<usize> {
        Ok(self.shape_fields()?.len())
    }

    /// Get typed, inert legacy form-code fields in document order.
    ///
    /// Returned values expose only stored text/checkbox/drop-down kind, opaque
    /// instructions, cached content, and dirty/lock state. This method never
    /// reads associated form-property XML, fills a form, changes a selection or
    /// checkbox state, invokes entry or exit macros, or refreshes a field.
    pub fn legacy_form_fields(&self) -> Result<Vec<LegacyForm>> {
        self.fields()?
            .iter()
            .map(|field| Field::legacy_form_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert legacy form-code fields.
    pub fn legacy_form_field_count(&self) -> Result<usize> {
        Ok(self.legacy_form_fields()?.len())
    }

    /// Get typed, inert `PRIVATE` conversion-data fields in document order.
    ///
    /// Returned values expose only stored opaque instructions, cached content,
    /// and dirty/lock state. This method never converts a document, interprets
    /// field data, changes hidden-text visibility or layout, or refreshes a
    /// field. `PRIVATE` is not treated as a confidentiality mechanism.
    pub fn private_fields(&self) -> Result<Vec<Private>> {
        self.fields()?
            .iter()
            .map(|field| Field::private_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `PRIVATE` conversion-data fields.
    pub fn private_field_count(&self) -> Result<usize> {
        Ok(self.private_fields()?.len())
    }

    /// Get typed, inert `DATABASE` query fields in document order.
    ///
    /// Returned values expose only stored opaque instructions, cached content,
    /// and dirty/lock state. This method never opens a data source or database,
    /// uses connection information, executes SQL, generates or inserts a table,
    /// changes layout, or refreshes a field.
    pub fn database_fields(&self) -> Result<Vec<Database>> {
        self.fields()?
            .iter()
            .map(|field| Field::database_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `DATABASE` query fields.
    pub fn database_field_count(&self) -> Result<usize> {
        Ok(self.database_fields()?.len())
    }

    /// Get typed, inert user-identity fields in document order.
    ///
    /// Returned values expose only stored kind, override, formatting, cached
    /// content, and dirty/lock state. This method never reads or modifies a host
    /// user's identity, applies formatting, or refreshes a field.
    pub fn user_identity_fields(&self) -> Result<Vec<UserIdentity>> {
        self.fields()?
            .iter()
            .map(|field| Field::user_identity_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert user-identity fields in the main document.
    pub fn user_identity_field_count(&self) -> Result<usize> {
        Ok(self.user_identity_fields()?.len())
    }

    /// Get typed, inert `ADVANCE` fields in document order.
    ///
    /// Returned values expose stored point adjustments, cached content, and
    /// dirty/lock state only. This method never moves text, changes layout,
    /// reflows content, or refreshes a field.
    pub fn advance_fields(&self) -> Result<Vec<Advance>> {
        self.fields()?
            .iter()
            .map(|field| Field::advance_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `ADVANCE` fields in the main document.
    pub fn advance_field_count(&self) -> Result<usize> {
        Ok(self.advance_fields()?.len())
    }

    /// Get typed, inert `DDE` and `DDEAUTO` fields in document order.
    ///
    /// Returned fields expose stored application, source, item, representation,
    /// storage, cached content, and dirty/lock metadata only. This method never
    /// launches an application, initiates a DDE conversation, opens a source,
    /// requests data, refreshes, converts, evaluates, or executes anything.
    pub fn dde_links(&self) -> Result<Vec<Dde>> {
        self.fields()?
            .iter()
            .map(|field| Field::dde_link(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `DDE` and `DDEAUTO` fields in the main document.
    pub fn dde_link_count(&self) -> Result<usize> {
        Ok(self.dde_links()?.len())
    }

    /// Get typed, inert `INCLUDETEXT`/`INCLUDEPICTURE` fields and historical
    /// `INCLUDE`/`IMPORT` aliases in document order.
    ///
    /// Returned fields expose stored source, bookmark, converter, XML, cached,
    /// and dirty/lock metadata only. This method never opens, resolves,
    /// imports, fetches, refreshes, converts, transforms, evaluates, or
    /// executes anything.
    pub fn external_includes(&self) -> Result<Vec<Include>> {
        self.fields()?
            .iter()
            .map(|field| Field::external_include(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert external-include fields in the main document.
    pub fn external_include_count(&self) -> Result<usize> {
        Ok(self.external_includes()?.len())
    }

    /// Get typed, inert RD referenced-document fields in document order.
    ///
    /// Returned fields expose stored paths, relative-path requests, switches,
    /// cached content, and dirty/lock metadata only. This method never opens,
    /// resolves, reads, imports, refreshes, evaluates, or executes a referenced
    /// document.
    pub fn referenced_documents(&self) -> Result<Vec<SubDocument>> {
        self.fields()?
            .iter()
            .map(|field| Field::referenced_document(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert referenced-document fields in the main document.
    pub fn referenced_document_count(&self) -> Result<usize> {
        Ok(self.referenced_documents()?.len())
    }

    /// Get typed, inert `LINK` fields in document order.
    ///
    /// Returned fields expose stored application, source, item, result,
    /// formatting, cached content, and dirty/lock metadata only. This method
    /// never activates an OLE server, launches an application, opens a source,
    /// requests data, refreshes, converts, evaluates, or executes anything.
    pub fn link_fields(&self) -> Result<Vec<Link>> {
        self.fields()?
            .iter()
            .map(|field| Field::link(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `LINK` fields in the main document.
    pub fn link_field_count(&self) -> Result<usize> {
        Ok(self.link_fields()?.len())
    }

    /// Get typed, inert table-of-contents fields in document order.
    ///
    /// Both simple (`w:fldSimple`) and complex (`w:fldChar`) TOC fields are
    /// discovered. Returned values expose the stored instruction, switches,
    /// cached result, and dirty/lock state; this method never paginates,
    /// regenerates a table of contents, follows its links, or executes fields.
    pub fn table_of_contents(&self) -> Result<Vec<Toc>> {
        self.fields()?
            .iter()
            .map(|field| Field::table_of_contents(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of table-of-contents fields in the main document.
    pub fn table_of_contents_count(&self) -> Result<usize> {
        Ok(self.table_of_contents()?.len())
    }

    /// Get typed, inert table-of-contents entry (`TC`) fields in document order.
    ///
    /// Returned fields expose only stored entry text, list identifiers, levels,
    /// page-number omission requests, switches, cached content, and dirty/lock
    /// state. This method never changes hidden text, calculates page numbers,
    /// generates a table of contents, or refreshes fields.
    pub fn table_of_contents_entries(&self) -> Result<Vec<TocEntry>> {
        self.fields()?
            .iter()
            .map(|field| Field::table_of_contents_entry(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert table-of-contents entry fields.
    pub fn table_of_contents_entry_count(&self) -> Result<usize> {
        Ok(self.table_of_contents_entries()?.len())
    }

    /// Get typed, inert table-of-authorities fields in document order.
    ///
    /// Returned fields expose stored switches and cached content only. This
    /// method never locates citation text, paginates the document, generates a
    /// table of authorities, or refreshes fields.
    pub fn tables_of_authorities(&self) -> Result<Vec<Toa>> {
        self.fields()?
            .iter()
            .map(|field| Field::table_of_authorities(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of table-of-authorities fields in the main document.
    pub fn table_of_authorities_count(&self) -> Result<usize> {
        Ok(self.tables_of_authorities()?.len())
    }

    /// Get typed, inert table-of-authorities entry (`TA`) fields in document order.
    ///
    /// These are stored citation markers. This method does not search for
    /// matching visible text, change hidden-text state, or generate a `TOA`.
    pub fn table_of_authorities_entries(&self) -> Result<Vec<ToaEntry>> {
        self.fields()?
            .iter()
            .map(|field| Field::table_of_authorities_entry(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of table-of-authorities entry fields in the main document.
    pub fn table_of_authorities_entry_count(&self) -> Result<usize> {
        Ok(self.table_of_authorities_entries()?.len())
    }

    /// Get typed, inert generated-index (`INDEX`) fields in document order.
    ///
    /// Returned fields expose stored switches and cached content only. This
    /// method never searches for index markers, sorts entries, calculates page
    /// references, generates an index, or refreshes fields.
    pub fn indexes(&self) -> Result<Vec<Index>> {
        self.fields()?
            .iter()
            .map(|field| Field::index(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of generated-index fields in the main document.
    pub fn index_count(&self) -> Result<usize> {
        Ok(self.indexes()?.len())
    }

    /// Get typed, inert index-entry (`XE`) fields in document order.
    ///
    /// These are stored index markers. This method does not change hidden text,
    /// resolve page-range bookmarks, sort entries, or generate an `INDEX`.
    pub fn index_entries(&self) -> Result<Vec<IndexEntry>> {
        self.fields()?
            .iter()
            .map(|field| Field::index_entry(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of index-entry fields in the main document.
    pub fn index_entry_count(&self) -> Result<usize> {
        Ok(self.index_entries()?.len())
    }

    /// Get typed, inert `MERGEFIELD` fields in document order.
    ///
    /// Returned values expose stored data-column names, switches, cached
    /// content, and dirty/lock state only. This method never opens a data
    /// source, resolves records, performs a merge, or refreshes field results.
    ///
    /// For backward-compatible access to the raw fields, use
    /// [`Self::merge_fields`].
    pub fn typed_merge_fields(&self) -> Result<Vec<Merge>> {
        self.fields()?
            .iter()
            .map(|field| Field::merge_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `MERGEFIELD` fields in the main document.
    pub fn typed_merge_field_count(&self) -> Result<usize> {
        Ok(self.typed_merge_fields()?.len())
    }

    /// Get typed, inert `DATA` mail-merge source fields in document order.
    ///
    /// Returned values expose only stored data-source and header-source
    /// identifiers, switches, cached content, and dirty/lock state. This method
    /// never opens, reads, connects to, resolves, or modifies either source; it
    /// never selects a record, performs a merge, or refreshes a field result.
    pub fn mail_merge_data_fields(&self) -> Result<Vec<MergeData>> {
        self.fields()?
            .iter()
            .map(|field| Field::mail_merge_data(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `DATA` mail-merge source fields.
    pub fn mail_merge_data_field_count(&self) -> Result<usize> {
        Ok(self.mail_merge_data_fields()?.len())
    }

    /// Get typed, inert `MERGEREC` and `MERGESEQ` fields in document order.
    ///
    /// Returned values expose stored kind, cached content, and dirty/lock state
    /// only. This method never selects or counts records, opens a data source,
    /// performs a merge, or refreshes field results.
    pub fn mail_merge_counters(&self) -> Result<Vec<MergeCounter>> {
        self.fields()?
            .iter()
            .map(|field| Field::mail_merge_counter(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert mail-merge counter fields in the main document.
    pub fn mail_merge_counter_count(&self) -> Result<usize> {
        Ok(self.mail_merge_counters()?.len())
    }

    /// Get typed, inert `NEXT` mail-merge control fields in document order.
    ///
    /// Returned values expose stored cached content and dirty/lock state only.
    /// This method never advances a record, opens a data source, performs a
    /// merge, or refreshes field results.
    pub fn mail_merge_next_fields(&self) -> Result<Vec<MergeNext>> {
        self.fields()?
            .iter()
            .map(|field| Field::mail_merge_next(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `NEXT` mail-merge control fields.
    pub fn mail_merge_next_field_count(&self) -> Result<usize> {
        Ok(self.mail_merge_next_fields()?.len())
    }

    /// Get typed, inert `NEXTIF` and `SKIPIF` fields in document order.
    ///
    /// Returned values expose stored comparison text, cached content, and
    /// dirty/lock state only. This method never evaluates a comparison, changes
    /// record selection, opens a data source, performs a merge, or refreshes
    /// field results.
    pub fn mail_merge_conditional_controls(&self) -> Result<Vec<MergeControl>> {
        self.fields()?
            .iter()
            .map(|field| Field::mail_merge_conditional_control(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert conditional mail-merge control fields.
    pub fn mail_merge_conditional_control_count(&self) -> Result<usize> {
        Ok(self.mail_merge_conditional_controls()?.len())
    }

    /// Get typed, inert `IF` fields in document order.
    ///
    /// Returned values expose stored expression text, cached content, and
    /// dirty/lock state only. This method never parses or evaluates an
    /// expression, resolves field values, or refreshes a field result.
    pub fn if_fields(&self) -> Result<Vec<If>> {
        self.fields()?
            .iter()
            .map(|field| Field::if_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `IF` fields.
    pub fn if_field_count(&self) -> Result<usize> {
        Ok(self.if_fields()?.len())
    }

    /// Get typed, inert `COMPARE` fields in document order.
    ///
    /// Returned values expose stored comparisons, cached content, and
    /// dirty/lock state only. This method never parses or evaluates a
    /// comparison, resolves nested field values, or refreshes a field.
    pub fn compare_fields(&self) -> Result<Vec<Compare>> {
        self.fields()?
            .iter()
            .map(|field| Field::compare_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `COMPARE` fields in the main document.
    pub fn compare_field_count(&self) -> Result<usize> {
        Ok(self.compare_fields()?.len())
    }

    /// Get typed, inert bookmark-reference fields in document order.
    ///
    /// Returned values expose stored kinds, targets, options, unknown switches,
    /// cached content, and dirty/lock state only. This method never looks up a
    /// bookmark, reads a referenced range or note, resolves a page number,
    /// creates a link, calculates a relative position, or refreshes a field.
    pub fn reference_fields(&self) -> Result<Vec<Reference>> {
        self.fields()?
            .iter()
            .map(|field| Field::reference_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert bookmark-reference fields in the main document.
    pub fn reference_field_count(&self) -> Result<usize> {
        Ok(self.reference_fields()?.len())
    }

    /// Get typed, inert `SET` fields in document order.
    ///
    /// Returned values expose stored target names, opaque expressions, cached
    /// content, and dirty/lock state only. This method never evaluates an
    /// expression, looks up or changes a bookmark, changes document state, or
    /// refreshes a field.
    pub fn set_fields(&self) -> Result<Vec<Set>> {
        self.fields()?
            .iter()
            .map(|field| Field::set_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `SET` fields in the main document.
    pub fn set_field_count(&self) -> Result<usize> {
        Ok(self.set_fields()?.len())
    }

    /// Get typed, inert `=` formula fields in document order.
    ///
    /// Returned values expose stored formulas, cached content, and dirty/lock
    /// state only. This method never parses or evaluates a formula, reads table
    /// cells or bookmarks, resolves field values, or refreshes a field.
    pub fn formula_fields(&self) -> Result<Vec<Formula>> {
        self.fields()?
            .iter()
            .map(|field| Field::formula_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert formula fields in the main document.
    pub fn formula_field_count(&self) -> Result<usize> {
        Ok(self.formula_fields()?.len())
    }

    /// Get typed, inert `EQ` equation fields in document order.
    ///
    /// Returned values expose stored expressions, cached content, and dirty/lock
    /// state only. This method never parses, calculates, formats, renders, or
    /// refreshes an equation.
    pub fn equations(&self) -> Result<Vec<Equation>> {
        self.fields()?
            .iter()
            .map(|field| Field::equation(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `EQ` fields in the main document.
    pub fn equation_count(&self) -> Result<usize> {
        Ok(self.equations()?.len())
    }

    /// Get typed, inert `SEQ` fields in document order.
    ///
    /// Returned values expose stored identifiers, optional bookmarks, opaque
    /// tails, cached content, and dirty/lock state only. This method never
    /// looks up a bookmark, increments or resets a sequence, calculates a
    /// number, or refreshes a field.
    pub fn sequence_fields(&self) -> Result<Vec<Sequence>> {
        self.fields()?
            .iter()
            .map(|field| Field::sequence_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `SEQ` fields in the main document.
    pub fn sequence_field_count(&self) -> Result<usize> {
        Ok(self.sequence_fields()?.len())
    }

    /// Get typed, inert `STYLEREF` fields in document order.
    ///
    /// Returned values expose stored style names, options, switches, cached
    /// content, and dirty/lock state only. This method never looks up styled
    /// text, searches document stories, calculates paragraph numbers or
    /// relative positions, resolves page layout, or refreshes a field.
    pub fn style_reference_fields(&self) -> Result<Vec<StyleReference>> {
        self.fields()?
            .iter()
            .map(|field| Field::style_reference_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `STYLEREF` fields in the main document.
    pub fn style_reference_field_count(&self) -> Result<usize> {
        Ok(self.style_reference_fields()?.len())
    }

    /// Get typed, inert `QUOTE` fields in document order.
    ///
    /// Returned values expose stored text arguments, switches, cached content,
    /// and dirty/lock state only. This method never interprets character codes,
    /// expands nested fields, inserts text, or refreshes a field result.
    pub fn quote_fields(&self) -> Result<Vec<Quote>> {
        self.fields()?
            .iter()
            .map(|field| Field::quote_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `QUOTE` fields in the main document.
    pub fn quote_field_count(&self) -> Result<usize> {
        Ok(self.quote_fields()?.len())
    }

    /// Get typed, inert `SYMBOL` fields in document order.
    ///
    /// Returned values expose stored character arguments, switches, cached
    /// content, and dirty/lock state only. This method never maps a character
    /// code, looks up a font, inserts a glyph, changes formatting or layout, or
    /// refreshes a field result.
    pub fn symbol_fields(&self) -> Result<Vec<Symbol>> {
        self.fields()?
            .iter()
            .map(|field| Field::symbol_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `SYMBOL` fields in the main document.
    pub fn symbol_field_count(&self) -> Result<usize> {
        Ok(self.symbol_fields()?.len())
    }

    /// Get typed, inert legacy automatic-numbering fields in document order.
    ///
    /// Returned values expose stored kinds, switches, cached content, and
    /// dirty/lock state only. This method never calculates paragraph numbers,
    /// reads heading or style state, changes paragraphs or layout, or refreshes
    /// a field result.
    pub fn auto_number_fields(&self) -> Result<Vec<AutoNumber>> {
        self.fields()?
            .iter()
            .map(|field| Field::auto_number_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert legacy automatic-numbering fields.
    pub fn auto_number_field_count(&self) -> Result<usize> {
        Ok(self.auto_number_fields()?.len())
    }

    /// Get typed, inert `LISTNUM` fields in document order.
    ///
    /// Returned values expose stored optional list names, switches, cached
    /// content, and dirty/lock state only. This method never looks up a list,
    /// determines a level or start value, calculates a number, changes layout,
    /// or refreshes a field result.
    pub fn list_number_fields(&self) -> Result<Vec<ListNumber>> {
        self.fields()?
            .iter()
            .map(|field| Field::list_number_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `LISTNUM` fields in the main document.
    pub fn list_number_field_count(&self) -> Result<usize> {
        Ok(self.list_number_fields()?.len())
    }

    /// Get typed, inert `ASK` and `FILLIN` fields in document order.
    ///
    /// Returned values expose stored prompt, bookmark, default-response, cached
    /// content, and dirty/lock state only. This method never displays a prompt,
    /// captures a response, creates or updates a bookmark, performs a merge, or
    /// refreshes a field result.
    pub fn prompt_fields(&self) -> Result<Vec<Prompt>> {
        self.fields()?
            .iter()
            .map(|field| Field::prompt_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `ASK` and `FILLIN` fields.
    pub fn prompt_field_count(&self) -> Result<usize> {
        Ok(self.prompt_fields()?.len())
    }

    /// Get typed, inert `ADDRESSBLOCK` and `GREETINGLINE` fields in document
    /// order.
    ///
    /// Returned values expose stored recipient layout, locale, country, fallback,
    /// cached-content, and dirty/lock state only. This method never opens a data
    /// source, selects a record, performs a merge, expands placeholders, generates
    /// text, or refreshes a field result.
    pub fn mail_merge_recipient_fields(&self) -> Result<Vec<Recipient>> {
        self.fields()?
            .iter()
            .map(|field| Field::mail_merge_recipient_field(field).map_err(Error::from))
            .filter_map(|result| result.transpose())
            .collect()
    }

    /// Get the number of typed, inert `ADDRESSBLOCK` and `GREETINGLINE` fields.
    pub fn mail_merge_recipient_field_count(&self) -> Result<usize> {
        Ok(self.mail_merge_recipient_fields()?.len())
    }

    /// Get all mail-merge fields in document order.
    ///
    /// This recognizes `MERGEFIELD` instructions represented by either
    /// `<w:fldSimple>` or complex `<w:fldChar>` field sequences.
    pub fn merge_fields(&self) -> Result<Vec<Field>> {
        Ok(self
            .fields()?
            .into_iter()
            .filter(Field::is_merge_field)
            .collect())
    }

    /// Get the data-source column names referenced by mail-merge fields.
    pub fn merge_field_names(&self) -> Result<Vec<String>> {
        Ok(self
            .merge_fields()?
            .into_iter()
            .filter_map(|field| field.merge_field_name().map(str::to_owned))
            .collect())
    }

    /// Get the numbering definitions for the document.
    ///
    /// Returns a numbering collection providing access to abstract numbering
    /// definitions and numbering instances used for lists.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if let Some(numbering) = doc.numbering()? {
    ///     println!("Document has {} numbering definitions", numbering.num_count());
    ///     for num in numbering.nums() {
    ///         println!("Num ID {}: references abstract num {}",
    ///             num.id(), num.abstract_num_id());
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    /// Load the bounded, inert classic-chart graph owned by this document.
    ///
    /// Returns every classic DrawingML chart anchored in the main document
    /// body together with its style, color-style, and embedded-workbook
    /// companion parts. See [`crate::chart::load`].
    pub fn chart_graph(&self) -> Result<crate::chart::Graph> {
        let main = self.opc.main_document_part()?.partname().clone();
        crate::chart::load(self.opc, &main)
    }
    /// Get all content controls in the document.
    ///
    /// Returns a vector of `ContentControl` objects representing structured
    /// content regions in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for control in doc.content_controls()? {
    ///     println!("Control ID {}", control.id());
    ///     if let Some(tag) = control.tag() {
    ///         println!("  Tag: {}", tag);
    ///     }
    ///     println!("  Type: {}", control.kind());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn content_controls(&self) -> Result<Vec<ContentControl>> {
        let xml_bytes = self.part.xml_bytes();
        ContentControl::extract_from_document(xml_bytes)
    }
}
