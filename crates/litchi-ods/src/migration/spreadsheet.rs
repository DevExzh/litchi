//! Main Spreadsheet structure and implementation.

use super::{
    CalculationSettings, Consolidation, ContentValidation, DataPilotTable, DatabaseRange, DdeLink,
    LabelRange, NamedDefinition, NamedDefinitionScope, NamedExpression, NamedRange, Sheet,
    SheetProtection, SpreadsheetProtection, SpreadsheetTrackedChanges, TableTemplate,
    calculation::parse_calculation_settings,
    consolidation::parse_consolidation,
    data_pilot::parse_data_pilot_tables,
    data_validation::parse_content_validations,
    database_range::parse_database_ranges,
    dde::parse_dde_links,
    label_range::parse_label_ranges,
    parser::Parser,
    protection::parse_protection,
    style_protection::{
        CellStyleProtection, CellStyleRegistry, ConditionalCellStyle, TableCellProtectionStyle,
    },
    table_template::parse_table_templates,
    tracked_changes::parse_tracked_changes,
};
use crate::core::{Content, Meta, OwnedPackage, Styles};
use litchi_core::{Error, Metadata, Result};
use std::path::Path;

/// An OpenDocument spreadsheet (.ods).
///
/// This struct represents a complete ODS spreadsheet and provides methods to access
/// its sheets, cells, and metadata.
///
/// # Examples
///
/// ```no_run
/// use litchi_odf::Spreadsheet;
///
/// # fn main() -> litchi_core::Result<()> {
/// let mut spreadsheet = Spreadsheet::open("data.ods")?;
///
/// // Get sheet count
/// println!("Sheets: {}", spreadsheet.sheet_count()?);
///
/// // Access first sheet
/// if let Some(sheet) = spreadsheet.sheet_by_index(0)? {
///     println!("Sheet: {}", sheet.name()?);
///     println!("Rows: {}, Columns: {}", sheet.row_count()?, sheet.column_count()?);
/// }
///
/// // Export to CSV
/// let csv = spreadsheet.to_csv()?;
/// # Ok(())
/// # }
/// ```
pub struct Spreadsheet {
    package: OwnedPackage,
    #[allow(dead_code)]
    content: Content,
    #[allow(dead_code)]
    styles: Option<Styles>,
    meta: Option<Meta>,
    named_definitions: Vec<NamedDefinition>,
    content_validations: Vec<ContentValidation>,
    database_ranges: Vec<DatabaseRange>,
    data_pilot_tables: Vec<DataPilotTable>,
    calculation_settings: Option<CalculationSettings>,
    label_ranges: Vec<LabelRange>,
    consolidation: Option<Consolidation>,
    dde_links: Vec<DdeLink>,
    protection: SpreadsheetProtection,
    sheet_protections: Vec<SheetProtection>,
    cell_styles: CellStyleRegistry,
    tracked_changes: Option<SpreadsheetTrackedChanges>,
    table_templates: Vec<TableTemplate>,
}

impl Spreadsheet {
    crate::script_package::script_facade_methods!();
    crate::annotation_package::annotation_facade_methods!(Spreadsheet);

    pub(crate) fn into_package(self) -> OwnedPackage {
        self.package
    }

    pub(crate) fn content_xml(&self) -> &str {
        self.content.xml_content()
    }

    pub(crate) fn styles_xml(&self) -> Option<&str> {
        self.styles.as_ref().map(Styles::xml_content)
    }

    /// Inspect named drawing fill-image definitions from spreadsheet styles.
    ///
    /// Links remain stored metadata: this does not follow them, load linked
    /// resources, or render images.
    pub fn drawing_fill_images(&self) -> Result<crate::drawing_fill_image::OdfDrawingFillImages> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::drawing_fill_image::parse_drawing_fill_images(styles.xml_content()),
        )
    }

    /// Inspect named legacy and SVG drawing gradients from spreadsheet styles.
    ///
    /// This does not resolve style use sites or render gradients.
    pub fn drawing_gradients(&self) -> Result<crate::drawing_gradient::OdfDrawingGradients> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::drawing_gradient::parse_drawing_gradients(styles.xml_content()),
        )
    }

    /// Inspect named drawing hatch definitions from spreadsheet styles.
    ///
    /// This does not resolve style use sites or render hatches.
    pub fn drawing_hatches(&self) -> Result<crate::drawing_hatch::OdfDrawingHatches> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::drawing_hatch::parse_drawing_hatches(styles.xml_content()),
        )
    }

    /// Inspect named drawing marker definitions from spreadsheet styles.
    ///
    /// This does not resolve style use sites or render marker paths.
    pub fn drawing_markers(&self) -> Result<crate::drawing_marker::OdfDrawingMarkers> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::drawing_marker::parse_drawing_markers(styles.xml_content()),
        )
    }

    /// Inspect named drawing opacity definitions from spreadsheet styles.
    ///
    /// This does not resolve style use sites or render opacity gradients.
    pub fn drawing_opacities(&self) -> Result<crate::drawing_opacity::OdfDrawingOpacities> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::drawing_opacity::parse_drawing_opacities(styles.xml_content()),
        )
    }

    /// Inspect named drawing stroke-dash definitions from spreadsheet styles.
    ///
    /// This does not resolve style use sites or render strokes.
    pub fn drawing_stroke_dashes(
        &self,
    ) -> Result<crate::drawing_stroke_dash::OdfDrawingStrokeDashes> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::drawing_stroke_dash::parse_drawing_stroke_dashes(styles.xml_content()),
        )
    }

    /// Open an ODS spreadsheet from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .ods file
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid ODS file.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let spreadsheet = Spreadsheet::open("data.ods")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let bytes = std::fs::read(path.as_ref())?;
        Self::from_bytes(bytes)
    }

    /// Open a password-encrypted ODS spreadsheet.
    pub fn open_with_password<P: AsRef<Path>>(
        path: P,
        password: impl Into<String>,
    ) -> Result<Self> {
        let bytes = std::fs::read(path.as_ref())?;
        Self::from_bytes_with_password(bytes, password)
    }

    /// Create a Spreadsheet from a byte buffer.
    ///
    /// # Arguments
    ///
    /// * `bytes` - Complete ODS file contents as bytes
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes do not represent a valid ODS file.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let bytes = std::fs::read("data.ods")?;
    /// let spreadsheet = Spreadsheet::from_bytes(bytes)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let owned_package = OwnedPackage::from_bytes(bytes)?;
        Self::from_owned_package(owned_package)
    }

    /// Create a spreadsheet from password-encrypted ODS bytes.
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        Self::from_owned_package(OwnedPackage::from_bytes_with_password(bytes, password)?)
    }

    fn from_owned_package(owned_package: OwnedPackage) -> Result<Self> {
        let package = owned_package.package()?;

        // Verify this is a spreadsheet
        let mime_type = package.mimetype();
        if !mime_type.contains("opendocument.spreadsheet") {
            return Err(Error::InvalidFormat(format!(
                "Not an ODS file: MIME type is {}",
                mime_type
            )));
        }

        // Parse core components
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;
        let named_definitions = Parser::parse_named_definitions(content.xml_content())?;
        super::named_expression::validate_named_definition_collection(&named_definitions)?;
        let content_validations = parse_content_validations(content.xml_content())?;
        let database_ranges = parse_database_ranges(content.xml_content())?;
        let data_pilot_tables = parse_data_pilot_tables(content.xml_content())?;
        let calculation_settings = parse_calculation_settings(content.xml_content())?;
        let label_ranges = parse_label_ranges(content.xml_content())?;
        let consolidation = parse_consolidation(content.xml_content())?;
        let dde_links = parse_dde_links(content.xml_content())?;
        let (protection, sheet_protections) = parse_protection(content.xml_content())?;
        let tracked_changes = parse_tracked_changes(content.xml_content())?;

        let styles = if package.has_file("styles.xml") {
            let styles_bytes = package.get_file("styles.xml")?;
            Some(Styles::from_bytes(&styles_bytes)?)
        } else {
            None
        };
        let cell_styles = CellStyleRegistry::parse(
            styles.as_ref().map(Styles::xml_content),
            content.xml_content(),
        )?;
        let mut template_parts = vec![content.xml_content()];
        if let Some(styles) = styles.as_ref().map(Styles::xml_content) {
            template_parts.push(styles);
        }
        let table_templates = parse_table_templates(&template_parts)?;

        let meta = if package.has_file("meta.xml") {
            let meta_bytes = package.get_file("meta.xml")?;
            Some(Meta::from_bytes(&meta_bytes)?)
        } else {
            None
        };

        Ok(Self {
            package: owned_package,
            content,
            styles,
            meta,
            named_definitions,
            content_validations,
            database_ranges,
            data_pilot_tables,
            calculation_settings,
            label_ranges,
            consolidation,
            dde_links,
            protection,
            sheet_protections,
            cell_styles,
            tracked_changes,
            table_templates,
        })
    }

    /// Return spreadsheet-wide formula calculation settings.
    pub fn calculation_settings(&self) -> Option<&CalculationSettings> {
        self.calculation_settings.as_ref()
    }

    /// Return spreadsheet row and column label ranges in document order.
    pub fn label_ranges(&self) -> &[LabelRange] {
        &self.label_ranges
    }

    /// Return the inert spreadsheet consolidation declaration.
    pub fn consolidation(&self) -> Option<&Consolidation> {
        self.consolidation.as_ref()
    }

    /// Return inert DDE declarations and their document-stored cached tables.
    pub fn dde_links(&self) -> &[DdeLink] {
        &self.dde_links
    }

    /// Return inert spreadsheet change-tracking metadata in document order.
    pub fn tracked_changes(&self) -> Option<&SpreadsheetTrackedChanges> {
        self.tracked_changes.as_ref()
    }

    /// Return named table-style templates from content and styles parts.
    pub fn table_templates(&self) -> &[TableTemplate] {
        &self.table_templates
    }

    /// Find a table-style template by its ODF name.
    pub fn table_template(&self, name: &str) -> Option<&TableTemplate> {
        self.table_templates
            .iter()
            .find(|template| template.name == name)
    }

    /// Create an ODS spreadsheet from raw bytes (ZIP archive data).
    ///
    /// This is used for single-pass parsing where the ZIP archive has already
    /// been validated during format detection. It avoids double-parsing.
    pub fn from_archive_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes(bytes)
    }

    /// Discover referenced, inline, missing, and inert linked images.
    pub fn images(&self) -> Result<Vec<crate::Image>> {
        let package = self.package.package()?;
        crate::media::scan_packaged_images(
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            |path| package.has_file(path),
            |path| package.manifest().get_media_type(path).map(str::to_string),
        )
    }

    /// Inspect classic forms without executing bindings, events, or external resources.
    pub fn forms(&self) -> Result<crate::OdfForms> {
        let mut parts = vec![(self.content.xml_content(), crate::OdfFormPart::Content)];
        if let Some(styles) = self.styles.as_ref().map(Styles::xml_content) {
            parts.push((styles, crate::OdfFormPart::Styles));
        }
        crate::form::parse_form_parts(&parts)
    }

    pub fn rdf_graphs(&self) -> Result<Vec<crate::rdf::Graph>> {
        crate::rdf::graphs(&self.package)
    }
    pub fn add_rdf_graph(
        &mut self,
        preferred_path: Option<&str>,
        triples: &[crate::rdf::Triple],
    ) -> Result<String> {
        let (bytes, path) = crate::rdf::add_graph(&self.package, preferred_path, triples)?;
        *self = Self::from_bytes(bytes)?;
        Ok(path)
    }
    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[crate::rdf::Triple]) -> Result<()> {
        let bytes = crate::rdf::replace_graph(&self.package, path, triples)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<()> {
        let bytes = crate::rdf::remove_graph(&self.package, path)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn add_rdf_triple(&mut self, path: &str, triple: &crate::rdf::Triple) -> Result<usize> {
        let index = self
            .rdf_graphs()?
            .into_iter()
            .find(|graph| graph.path == path)
            .ok_or_else(|| Error::InvalidFormat(format!("RDF graph '{path}' was not found")))?
            .triples
            .len();
        let (bytes, _) = crate::rdf::add_triple(&self.package, path, triple)?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }
    pub fn replace_rdf_triple(
        &mut self,
        path: &str,
        index: usize,
        triple: &crate::rdf::Triple,
    ) -> Result<()> {
        let bytes = crate::rdf::replace_triple(&self.package, path, index, triple)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn remove_rdf_triple(&mut self, path: &str, index: usize) -> Result<()> {
        let bytes = crate::rdf::remove_triple(&self.package, path, index)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> Result<()> {
        let bytes = crate::rdf::move_triple(&self.package, path, from, to)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn add_form(&mut self, group_index: usize, form: &crate::OdfAuthoredForm) -> Result<usize> {
        let (bytes, index) = crate::form_package::add_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::form_package::FormHost::Spreadsheet,
            group_index,
            None,
            form,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }
    pub fn add_nested_form(
        &mut self,
        parent_form: usize,
        form: &crate::OdfAuthoredForm,
    ) -> Result<usize> {
        let (bytes, index) = crate::form_package::add_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::form_package::FormHost::Spreadsheet,
            0,
            Some(parent_form),
            form,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }
    pub fn replace_form(&mut self, index: usize, form: &crate::OdfAuthoredForm) -> Result<()> {
        let bytes = crate::form_package::replace_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            form,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn remove_form(&mut self, index: usize) -> Result<()> {
        let bytes = crate::form_package::remove_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn move_form(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::form_package::move_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn add_form_control(
        &mut self,
        form_index: usize,
        control: &crate::OdfAuthoredFormControl,
    ) -> Result<usize> {
        let (bytes, index) = crate::form_package::add_control(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            form_index,
            control,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }
    pub fn replace_form_control(
        &mut self,
        index: usize,
        control: &crate::OdfAuthoredFormControl,
    ) -> Result<()> {
        let bytes = crate::form_package::replace_control(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            control,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn remove_form_control(&mut self, index: usize) -> Result<()> {
        let bytes = crate::form_package::remove_control(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn move_form_control(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::form_package::move_control(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Inspect ordered ODF variable declarations without evaluating fields or formulas.
    pub fn variable_declarations(&self) -> Result<crate::OdfVariableDeclarations> {
        let mut parts = vec![(self.content.xml_content(), crate::OdfVariablePart::Content)];
        if let Some(styles) = self.styles.as_ref().map(Styles::xml_content) {
            parts.push((styles, crate::OdfVariablePart::Styles));
        }
        crate::variable_declaration::parse_variable_declaration_parts(&parts)
    }

    /// Discover package, inline, missing, and inert linked embedded objects.
    pub fn embedded_objects(&self) -> Result<Vec<crate::OdfEmbeddedObject>> {
        let package = self.package.package()?;
        crate::embedded_object::scan_packaged_objects(
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            |path| package.has_file(path),
            |path| package.manifest().get_media_type(path).map(str::to_string),
        )
    }

    pub fn embedded_chart(&self, index: usize) -> Result<crate::ChartDocument> {
        crate::embedded_chart::open_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )
    }

    pub fn add_embedded_chart(
        &mut self,
        sheet_name: &str,
        definition: &crate::ChartDefinition,
    ) -> Result<usize> {
        self.add_embedded_chart_with_storage(
            sheet_name,
            definition,
            crate::OdfEmbeddedChartStorage::PackageSubdocument,
        )
    }

    pub fn add_embedded_chart_with_storage(
        &mut self,
        sheet_name: &str,
        definition: &crate::ChartDefinition,
        storage: crate::OdfEmbeddedChartStorage,
    ) -> Result<usize> {
        let (bytes, index) = crate::embedded_chart::add_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::embedded_chart::EmbeddedChartHost::Sheet(sheet_name),
            storage,
            definition,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }

    pub fn replace_embedded_chart(
        &mut self,
        index: usize,
        definition: &crate::ChartDefinition,
    ) -> Result<()> {
        let bytes = crate::embedded_chart::replace_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            definition,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn remove_embedded_chart(&mut self, index: usize) -> Result<()> {
        let bytes = crate::embedded_chart::remove_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn add_embedded_resource(
        &mut self,
        sheet_name: &str,
        resource: &crate::OdfEmbeddedResource,
    ) -> Result<usize> {
        let (bytes, index) = crate::embedded_package::add(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::embedded_chart::EmbeddedChartHost::Sheet(sheet_name),
            resource,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }

    pub fn replace_embedded_object(
        &mut self,
        index: usize,
        resource: &crate::OdfEmbeddedResource,
    ) -> Result<()> {
        let bytes = crate::embedded_package::replace(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::embedded_package::ResourceTarget::Object,
            resource,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn replace_embedded_image(
        &mut self,
        index: usize,
        resource: &crate::OdfEmbeddedResource,
    ) -> Result<()> {
        let bytes = crate::embedded_package::replace(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::embedded_package::ResourceTarget::Image,
            resource,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn remove_embedded_object(&mut self, index: usize) -> Result<()> {
        let bytes = crate::embedded_package::remove(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::embedded_package::ResourceTarget::Object,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn remove_embedded_image(&mut self, index: usize) -> Result<()> {
        let bytes = crate::embedded_package::remove(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::embedded_package::ResourceTarget::Image,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn move_embedded_object(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::embedded_package::reorder(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
            crate::embedded_package::ResourceTarget::Object,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn move_embedded_image(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::embedded_package::reorder(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
            crate::embedded_package::ResourceTarget::Image,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Return bytes only for inline or verified package-contained images.
    /// Linked images remain inert and are never fetched.
    pub fn image_bytes(&self, image: &crate::Image) -> Result<Option<Vec<u8>>> {
        match &image.source {
            crate::ImageSource::Inline { bytes, .. } => Ok(Some(bytes.clone())),
            crate::ImageSource::PackagePart { path, .. } => self.package.get_file(path).map(Some),
            _ => Ok(None),
        }
    }

    /// Get the number of sheets in the spreadsheet.
    pub fn sheet_count(&mut self) -> Result<usize> {
        let sheets = self.sheets()?;
        Ok(sheets.len())
    }

    /// Get all sheets in the spreadsheet.
    ///
    /// Returns a vector of `Sheet` objects representing all sheets in the document.
    pub fn sheets(&mut self) -> Result<Vec<Sheet>> {
        let package = self.package.package()?;
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        let mut sheets = Parser::parse_sheets(content.xml_content())?;
        if sheets.len() != self.sheet_protections.len() {
            return Err(Error::InvalidFormat(format!(
                "sheet protection count {} does not match sheet count {}",
                self.sheet_protections.len(),
                sheets.len()
            )));
        }
        for (sheet, protection) in sheets.iter_mut().zip(&self.sheet_protections) {
            sheet.protection = protection.clone();
        }
        for cell in sheets
            .iter()
            .flat_map(|sheet| sheet.rows.iter())
            .flat_map(|row| row.cells.iter())
        {
            if let Some(name) = cell.validation_name.as_deref()
                && self.content_validation(name).is_none()
            {
                return Err(Error::InvalidFormat(format!(
                    "cell references missing content validation '{name}'"
                )));
            }
        }
        Ok(sheets)
    }

    /// Create an immutable snapshot for the shared spreadsheet trait API.
    ///
    /// This parses the worksheet content once and keeps the resulting values
    /// in an owned read-only model.  Use it with consumers such as
    /// `litchi_eval::FormulaEvaluator` without repeatedly reparsing the ODS
    /// package.  The original `Spreadsheet` remains usable for package edits.
    pub fn evaluation_workbook(&mut self) -> Result<super::Workbook> {
        super::Workbook::from_sheets(self.sheets()?)
    }

    /// Consume this spreadsheet into an immutable shared-workbook snapshot.
    ///
    /// Unlike [`Self::evaluation_workbook`], this avoids retaining the ODS
    /// package after its sheets have been materialized.
    pub fn into_evaluation_workbook(mut self) -> Result<super::Workbook> {
        super::Workbook::from_sheets(self.sheets()?)
    }

    /// Return all named ranges and expressions in document order.
    pub fn named_definitions(&self) -> &[NamedDefinition] {
        &self.named_definitions
    }

    /// Find either named-definition kind by name and scope.
    pub fn find_named_definition(
        &self,
        name: &str,
        scope: &NamedDefinitionScope,
    ) -> Option<&NamedDefinition> {
        self.named_definitions
            .iter()
            .find(|value| value.name() == name && value.scope() == scope)
    }

    /// Add a global or sheet-local named definition atomically.
    pub fn add_named_definition(&mut self, definition: &NamedDefinition) -> Result<()> {
        let bytes = crate::ods_definition_package::add_named(
            &self.package,
            self.content.xml_content(),
            &self.named_definitions,
            definition,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Replace a named definition without changing its scope.
    pub fn replace_named_definition(
        &mut self,
        name: &str,
        scope: &NamedDefinitionScope,
        replacement: &NamedDefinition,
    ) -> Result<()> {
        let bytes = crate::ods_definition_package::replace_named(
            &self.package,
            self.content.xml_content(),
            &self.named_definitions,
            name,
            scope,
            replacement,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Update attributes while preserving unknown attributes and the original body.
    pub fn update_named_definition(
        &mut self,
        name: &str,
        scope: &NamedDefinitionScope,
        update: &crate::NamedDefinitionUpdate,
    ) -> Result<()> {
        let bytes = crate::ods_definition_package::update_named(
            &self.package,
            self.content.xml_content(),
            &self.named_definitions,
            name,
            scope,
            update,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Remove a named definition unless another named expression depends on it.
    pub fn remove_named_definition(
        &mut self,
        name: &str,
        scope: &NamedDefinitionScope,
    ) -> Result<()> {
        let bytes = crate::ods_definition_package::remove_named(
            &self.package,
            self.content.xml_content(),
            &self.named_definitions,
            name,
            scope,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Reorder named definitions within one global or sheet-local scope.
    pub fn reorder_named_definition(
        &mut self,
        scope: &NamedDefinitionScope,
        from: usize,
        to: usize,
    ) -> Result<()> {
        let bytes = crate::ods_definition_package::reorder_named(
            &self.package,
            self.content.xml_content(),
            &self.named_definitions,
            scope,
            from,
            to,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Return document-level spreadsheet content validations in document order.
    pub fn content_validations(&self) -> &[ContentValidation] {
        &self.content_validations
    }

    /// Return spreadsheet database ranges, filters, sort keys, and subtotal rules.
    ///
    /// External database sources are inert metadata and are never executed.
    pub fn database_ranges(&self) -> &[DatabaseRange] {
        &self.database_ranges
    }

    /// Find a uniquely named database range.
    pub fn find_database_range(&self, name: &str) -> Option<&DatabaseRange> {
        self.database_ranges
            .iter()
            .find(|range| range.name.as_deref() == Some(name))
    }

    /// Add a database range without refreshing or executing its source metadata.
    pub fn add_database_range(&mut self, range: &DatabaseRange) -> Result<()> {
        let bytes = crate::ods_definition_package::add_database(
            &self.package,
            self.content.xml_content(),
            &self.database_ranges,
            range,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Replace a database range atomically.
    pub fn replace_database_range(&mut self, index: usize, range: &DatabaseRange) -> Result<()> {
        let bytes = crate::ods_definition_package::replace_database(
            &self.package,
            self.content.xml_content(),
            &self.database_ranges,
            index,
            range,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Update attributes while preserving filter, sort, subtotal, source, and extension children.
    pub fn update_database_range(
        &mut self,
        index: usize,
        update: &crate::DatabaseRangeUpdate,
    ) -> Result<()> {
        let bytes = crate::ods_definition_package::update_database(
            &self.package,
            self.content.xml_content(),
            &self.database_ranges,
            index,
            update,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Remove a database range atomically.
    pub fn remove_database_range(&mut self, index: usize) -> Result<()> {
        let bytes = crate::ods_definition_package::remove_database(
            &self.package,
            self.content.xml_content(),
            &self.database_ranges,
            index,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Reorder database ranges inside their schema container.
    pub fn reorder_database_range(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::ods_definition_package::reorder_database(
            &self.package,
            self.content.xml_content(),
            &self.database_ranges,
            from,
            to,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Return data-pilot (pivot-table) declarations.
    pub fn data_pilot_tables(&self) -> &[DataPilotTable] {
        &self.data_pilot_tables
    }

    /// Find a data-pilot table by its collection-unique name.
    pub fn find_data_pilot_table(&self, name: &str) -> Option<&DataPilotTable> {
        self.data_pilot_tables
            .iter()
            .find(|table| table.name == name)
    }

    /// Add a data-pilot declaration without executing its database, query, or service source.
    pub fn add_data_pilot_table(&mut self, table: &DataPilotTable) -> Result<usize> {
        let (bytes, index) =
            crate::data_pilot_package::add(&self.package, self.content.xml_content(), table)?;
        let replacement = Self::from_bytes(bytes)?;
        *self = replacement;
        Ok(index)
    }

    /// Replace a complete data-pilot declaration atomically.
    pub fn replace_data_pilot_table(&mut self, index: usize, table: &DataPilotTable) -> Result<()> {
        let bytes = crate::data_pilot_package::replace(
            &self.package,
            self.content.xml_content(),
            index,
            table,
        )?;
        let replacement = Self::from_bytes(bytes)?;
        *self = replacement;
        Ok(())
    }

    /// Update top-level data-pilot metadata while preserving its original body and extensions.
    pub fn update_data_pilot_table(
        &mut self,
        index: usize,
        update: &crate::DataPilotTableUpdate,
    ) -> Result<()> {
        let bytes = crate::data_pilot_package::update(
            &self.package,
            self.content.xml_content(),
            index,
            update,
        )?;
        let replacement = Self::from_bytes(bytes)?;
        *self = replacement;
        Ok(())
    }

    /// Remove a data-pilot declaration atomically.
    pub fn remove_data_pilot_table(&mut self, index: usize) -> Result<()> {
        let bytes =
            crate::data_pilot_package::remove(&self.package, self.content.xml_content(), index)?;
        let replacement = Self::from_bytes(bytes)?;
        *self = replacement;
        Ok(())
    }

    /// Reorder data-pilot declarations within their schema container.
    pub fn reorder_data_pilot_table(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::data_pilot_package::reorder(
            &self.package,
            self.content.xml_content(),
            from,
            to,
        )?;
        let replacement = Self::from_bytes(bytes)?;
        *self = replacement;
        Ok(())
    }

    /// Find a content-validation definition by name.
    pub fn content_validation(&self, name: &str) -> Option<&ContentValidation> {
        self.content_validations
            .iter()
            .find(|validation| validation.name == name)
    }

    /// Return document-structure protection metadata.
    pub fn protection(&self) -> &SpreadsheetProtection {
        &self.protection
    }

    /// Resolve the inherited `style:cell-protect` value for a cell.
    pub fn cell_style_protection(&self, cell: &super::Cell) -> Result<Option<CellStyleProtection>> {
        self.cell_styles.resolve(cell.style_name())
    }

    /// Standard ODF conditional table-cell styles in effective document order.
    ///
    /// Conditions are returned as inert text and are never evaluated.
    pub fn conditional_cell_styles(&self) -> &[ConditionalCellStyle] {
        self.cell_styles.conditional_styles()
    }

    /// Find a standard ODF conditional table-cell style by style name.
    pub fn conditional_cell_style(&self, style_name: &str) -> Option<&ConditionalCellStyle> {
        self.cell_styles.conditional_style(style_name)
    }

    /// Automatic table-cell styles with an explicit `style:cell-protect` value.
    pub fn table_cell_protection_styles(&self) -> &[TableCellProtectionStyle] {
        self.cell_styles.automatic_protection_styles()
    }

    /// Return all named ranges, including global and sheet-local ranges.
    pub fn named_ranges(&self) -> impl Iterator<Item = &NamedRange> {
        self.named_definitions
            .iter()
            .filter_map(|definition| match definition {
                NamedDefinition::Range(range) => Some(range),
                NamedDefinition::Expression(_) => None,
            })
    }

    /// Find a named range by name and scope.
    pub fn named_range(&self, name: &str, scope: &NamedDefinitionScope) -> Option<&NamedRange> {
        self.named_definitions
            .iter()
            .find_map(|definition| match definition {
                NamedDefinition::Range(range) if range.name == name && &range.scope == scope => {
                    Some(range)
                },
                _ => None,
            })
    }

    /// Return all named expressions, including global and sheet-local expressions.
    pub fn named_expressions(&self) -> impl Iterator<Item = &NamedExpression> {
        self.named_definitions
            .iter()
            .filter_map(|definition| match definition {
                NamedDefinition::Expression(expression) => Some(expression),
                NamedDefinition::Range(_) => None,
            })
    }

    /// Find a named expression by name and scope.
    pub fn named_expression(
        &self,
        name: &str,
        scope: &NamedDefinitionScope,
    ) -> Option<&NamedExpression> {
        self.named_definitions
            .iter()
            .find_map(|definition| match definition {
                NamedDefinition::Expression(expression)
                    if expression.name == name && &expression.scope == scope =>
                {
                    Some(expression)
                },
                _ => None,
            })
    }

    /// Get a sheet by name.
    ///
    /// Returns `Some(sheet)` if a sheet with the given name exists, `None` otherwise.
    ///
    /// # Arguments
    ///
    /// * `name` - Name of the sheet to find
    pub fn sheet_by_name(&mut self, name: &str) -> Result<Option<Sheet>> {
        let sheets = self.sheets()?;
        Ok(sheets.into_iter().find(|sheet| sheet.name == name))
    }

    /// Get a sheet by index.
    ///
    /// Returns `Some(sheet)` if a sheet exists at the given index, `None` otherwise.
    ///
    /// # Arguments
    ///
    /// * `index` - 0-based index of the sheet
    pub fn sheet_by_index(&mut self, index: usize) -> Result<Option<Sheet>> {
        let sheets = self.sheets()?;
        Ok(sheets.into_iter().nth(index))
    }

    /// Extract all text content from the spreadsheet.
    ///
    /// Returns text from all cells, separated by newlines.
    pub fn text(&mut self) -> Result<String> {
        let sheets = self.sheets()?;
        let mut all_text = Vec::new();

        for sheet in sheets {
            for row in sheet.rows {
                for cell in row.cells {
                    if !cell.text.trim().is_empty() {
                        all_text.push(cell.text.trim().to_string());
                    }
                }
            }
        }

        Ok(all_text.join("\n"))
    }

    /// Export spreadsheet data as CSV.
    ///
    /// Converts all sheets to CSV format, with sheets separated by double newlines.
    /// Properly escapes CSV special characters.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut spreadsheet = Spreadsheet::open("data.ods")?;
    /// let csv = spreadsheet.to_csv()?;
    /// std::fs::write("output.csv", csv)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_csv(&mut self) -> Result<String> {
        let sheets = self.sheets()?;
        let mut csv_output = String::new();

        for (sheet_index, sheet) in sheets.iter().enumerate() {
            if sheet_index > 0 {
                csv_output.push_str("\n\n"); // Separate sheets with double newline
            }

            for (row_index, row) in sheet.rows.iter().enumerate() {
                if row_index > 0 {
                    csv_output.push('\n');
                }

                for (col_index, cell) in row.cells.iter().enumerate() {
                    if col_index > 0 {
                        csv_output.push(',');
                    }

                    // Escape CSV special characters and wrap in quotes if needed
                    let cell_text = &cell.text;
                    if cell_text.contains(',')
                        || cell_text.contains('"')
                        || cell_text.contains('\n')
                    {
                        let escaped = cell_text.replace('"', "\"\"");
                        csv_output.push('"');
                        csv_output.push_str(&escaped);
                        csv_output.push('"');
                    } else {
                        csv_output.push_str(cell_text);
                    }
                }
            }
        }

        Ok(csv_output)
    }

    /// Get document metadata.
    ///
    /// Extracts metadata from the meta.xml file.
    pub fn metadata(&self) -> Result<Metadata> {
        if let Some(meta) = &self.meta {
            meta.try_extract_metadata()
        } else {
            Ok(Metadata::default())
        }
    }

    /// Get the complete format-specific OpenDocument metadata model.
    pub fn odf_metadata(&self) -> Result<Option<crate::Metadata>> {
        self.meta.as_ref().map(Meta::odf_metadata).transpose()
    }

    // Note: For spreadsheet modification operations, see `MutableSpreadsheet` which provides
    // full CRUD operations on sheets, rows, and cells including set_cell, clear_cell, add/remove
    // rows and sheets.

    /// Save the spreadsheet to a new file.
    ///
    /// This method saves the current spreadsheet state to a new file.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODS file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let spreadsheet = Spreadsheet::open("input.ods")?;
    /// spreadsheet.save("output.ods")?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Note
    ///
    /// Full spreadsheet modification support is planned for future releases. For now,
    /// to modify a spreadsheet, use `SpreadsheetBuilder` to create a new one with
    /// the desired content.
    pub fn save<P: AsRef<std::path::Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert the spreadsheet to bytes.
    ///
    /// This method serializes the spreadsheet to an ODF-compliant ZIP archive.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let spreadsheet = Spreadsheet::open("data.ods")?;
    /// let bytes = spreadsheet.to_bytes()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        Ok(self.package.as_bytes().to_vec())
    }

    // Note: DELETE operations are available via `MutableSpreadsheet`. To modify this spreadsheet:
    //   1. Convert: `let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet)?`
    //   2. Modify: `mutable.remove_sheet(0)?`, `mutable.set_cell(0, 0, 0, value)?`, etc.
    //   3. Save: `mutable.save("output.ods")?`
    // Available methods: remove_sheet, remove_row, set_cell, clear_cell, clear_sheet, etc.
}
