//! Styles-part resources, presentation metadata, and master-page edits.

use super::super::super::model::MutableDocument;
use crate::core::Structure;
use crate::header_footer::{Master, read, set_text, set_xml};
use crate::master_page::{add, insert, remove, replace};
use crate::page_layout::{PageLayout, parse_page_layouts, set_page_layout_xml};
use litchi_core::Result;

impl MutableDocument {
    /// Return typed named ruby styles from the current `styles.xml`.
    pub fn ruby_styles(&self) -> Result<crate::ruby_family::Styles> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(Default::default()), crate::parse_ruby_styles)
    }

    /// Insert or replace one named ruby style definition and return the old value.
    pub fn set_ruby_style(
        &mut self,
        style: &crate::ruby_family::Style,
    ) -> Result<Option<crate::ruby_family::Style>> {
        style.validate()?;
        let old = self.ruby_styles()?.get(&style.name).cloned();
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(Structure::default_styles_xml);
        self.styles_xml = Some(crate::set_ruby_style_xml(&styles, style)?);
        Ok(old)
    }

    /// Remove one named ruby style definition and return the old value.
    ///
    /// Existing `text:ruby` style references are preserved verbatim, so callers
    /// can intentionally manage their lifecycle separately.
    pub fn remove_ruby_style(&mut self, name: &str) -> Result<Option<crate::ruby_family::Style>> {
        let old = self.ruby_styles()?.get(name).cloned();
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        self.styles_xml = Some(crate::remove_ruby_style_xml(styles, name)?);
        Ok(old)
    }

    /// Return font-face declarations from the current `styles.xml`.
    ///
    /// Linked font resources remain inert metadata. This does not fetch a URI,
    /// load a font, or inspect embedded font data.
    pub fn styles_font_face_declarations(&self) -> Result<Option<crate::font_face::Declarations>> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(None),
            crate::font_face::parse_styles_font_face_declarations,
        )
    }

    /// Replace styles-part font-face declarations and return the old value.
    ///
    /// This edits `styles.xml` only. It does not fetch linked font resources,
    /// load a font, or inspect embedded font data.
    pub fn set_styles_font_face_declarations(
        &mut self,
        declarations: &crate::font_face::Declarations,
    ) -> Result<Option<crate::font_face::Declarations>> {
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(Structure::default_styles_xml);
        let (updated, old) =
            crate::font_face::set_styles_font_face_declarations_xml(&styles, declarations)?;
        self.styles_xml = Some(updated);
        Ok(old)
    }

    /// Remove styles-part font-face declarations and return the old value.
    ///
    /// This edits `styles.xml` only. Existing style references remain
    /// verbatim so callers can manage their lifecycle separately.
    pub fn clear_styles_font_face_declarations(
        &mut self,
    ) -> Result<Option<crate::font_face::Declarations>> {
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        let (updated, old) = crate::font_face::remove_styles_font_face_declarations_xml(styles)?;
        self.styles_xml = Some(updated);
        Ok(old)
    }

    /// Return named legacy and SVG drawing gradients from current styles metadata.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites, load external data, or render gradients.
    pub fn drawing_gradients(&self) -> Result<crate::drawing::resources::gradient::Collection> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing::resources::gradient::parse_drawing_gradients,
        )
    }

    /// Return named drawing hatch resources from current styles metadata.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render hatches.
    pub fn drawing_hatches(&self) -> Result<crate::drawing::resources::hatch::Collection> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing::resources::hatch::parse_drawing_hatches,
        )
    }

    /// Return named drawing stroke-dash resources from current styles metadata.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render strokes.
    pub fn drawing_stroke_dashes(
        &self,
    ) -> Result<crate::drawing::resources::stroke_dash::Collection> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing::resources::stroke_dash::parse_drawing_stroke_dashes,
        )
    }

    /// Return named drawing fill-image definitions from current styles metadata.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites, follow links, load linked resources, or render images.
    pub fn drawing_fill_images(&self) -> Result<crate::drawing::resources::fill_image::Collection> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing::resources::fill_image::parse_drawing_fill_images,
        )
    }

    /// Return named drawing marker definitions from current styles metadata.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites or render marker paths.
    pub fn drawing_markers(&self) -> Result<crate::drawing::resources::marker::Collection> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing::resources::marker::parse_drawing_markers,
        )
    }

    /// Return named drawing opacity definitions from current styles metadata.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites or render opacity gradients.
    pub fn drawing_opacities(&self) -> Result<crate::drawing::resources::opacity::Collection> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing::resources::opacity::parse_drawing_opacities,
        )
    }

    /// Return stored footnote and endnote presentation configurations.
    ///
    /// The result describes style metadata only. It never renumbers, lays out,
    /// or renders notes.
    pub fn notes_configurations(&self) -> Result<crate::notes_configuration::Configurations> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(Default::default()), crate::notes_configuration::parse)
    }

    /// Return stored outline numbering styles from current styles metadata.
    ///
    /// The result does not apply styles to headings, generate labels, or
    /// update tables of contents.
    pub fn outline_styles(&self) -> Result<crate::outline_style::Styles> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::outline_style::parse_outline_styles,
        )
    }

    /// Insert or replace one named outline numbering style.
    ///
    /// This edits `styles.xml` only and returns the previous style with the
    /// same name. It does not alter heading structure or cached index content.
    pub fn set_outline_style(
        &mut self,
        style: &crate::outline_style::Style,
    ) -> Result<Option<crate::outline_style::Style>> {
        style.validate()?;
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(Structure::default_styles_xml);
        let (updated, old) = crate::outline_style::set_outline_style_xml(&styles, style)?;
        self.styles_xml = Some(updated);
        Ok(old)
    }

    /// Remove one named outline numbering style and return its prior value.
    ///
    /// Existing heading references are retained verbatim, allowing callers to
    /// manage those references separately.
    pub fn remove_outline_style(
        &mut self,
        name: &str,
    ) -> Result<Option<crate::outline_style::Style>> {
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        let (updated, old) = crate::outline_style::remove_outline_style_xml(styles, name)?;
        self.styles_xml = Some(updated);
        Ok(old)
    }

    /// Insert or replace one stored footnote or endnote configuration.
    ///
    /// This edits `styles.xml` only and returns the prior configuration for the
    /// same note class. It never changes note anchors, citations, or numbering.
    pub fn set_notes_configuration(
        &mut self,
        configuration: &crate::notes_configuration::Configuration,
    ) -> Result<Option<crate::notes_configuration::Configuration>> {
        configuration.validate()?;
        let old = self
            .notes_configurations()?
            .get(configuration.note_class)
            .cloned();
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(Structure::default_styles_xml);
        self.styles_xml = Some(crate::notes_configuration::set_xml(&styles, configuration)?);
        Ok(old)
    }

    /// Replace both stored note-class configurations and return the old values.
    ///
    /// An absent class is removed from `styles.xml`. This updates metadata only and
    /// never recalculates citations, sequence numbers, or page layout.
    pub fn set_notes_configurations(
        &mut self,
        configurations: &crate::notes_configuration::Configurations,
    ) -> Result<crate::notes_configuration::Configurations> {
        configurations.validate()?;
        let old = self.notes_configurations()?;
        if self.styles_xml.is_none()
            && configurations.footnote.is_none()
            && configurations.endnote.is_none()
        {
            return Ok(old);
        }
        let mut styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(Structure::default_styles_xml);
        for note_class in crate::notes_configuration::Class::ALL {
            styles = match configurations.get(note_class) {
                Some(configuration) => crate::notes_configuration::set_xml(&styles, configuration)?,
                None => crate::notes_configuration::remove_xml(&styles, note_class)?,
            };
        }
        self.styles_xml = Some(styles);
        Ok(old)
    }

    /// Remove one stored note-class configuration and return its prior value.
    ///
    /// This edits style metadata only. Existing notes and their cached citations
    /// are preserved verbatim.
    pub fn clear_notes_configuration(
        &mut self,
        note_class: crate::notes_configuration::Class,
    ) -> Result<Option<crate::notes_configuration::Configuration>> {
        let old = self.notes_configurations()?.get(note_class).cloned();
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        self.styles_xml = Some(crate::notes_configuration::remove_xml(styles, note_class)?);
        Ok(old)
    }

    /// Return the stored document-wide bibliography formatting policy.
    ///
    /// The policy is styles metadata only. It is never used to generate
    /// bibliography entries, resolve citations, or access external sources.
    pub fn bibliography_configuration(
        &self,
    ) -> Result<Option<crate::bibliography_configuration::Configuration>> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(None),
            crate::bibliography_configuration::parse_bibliography_configuration,
        )
    }

    /// Insert or replace the document-wide bibliography formatting policy.
    ///
    /// This edits `styles.xml` only and returns the prior policy. It does not
    /// regenerate bibliography entries or modify bibliography marks.
    pub fn set_bibliography_configuration(
        &mut self,
        configuration: &crate::bibliography_configuration::Configuration,
    ) -> Result<Option<crate::bibliography_configuration::Configuration>> {
        configuration.validate()?;
        let old = self.bibliography_configuration()?;
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(Structure::default_styles_xml);
        self.styles_xml = Some(
            crate::bibliography_configuration::set_bibliography_configuration_xml(
                &styles,
                configuration,
            )?,
        );
        Ok(old)
    }

    /// Remove the document-wide bibliography formatting policy.
    ///
    /// This edits styles metadata only. Existing bibliography entries and
    /// source marks are preserved verbatim.
    pub fn clear_bibliography_configuration(
        &mut self,
    ) -> Result<Option<crate::bibliography_configuration::Configuration>> {
        let old = self.bibliography_configuration()?;
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        self.styles_xml =
            Some(crate::bibliography_configuration::remove_bibliography_configuration_xml(styles)?);
        Ok(old)
    }

    /// Return stored document line-numbering configuration from current styles.
    ///
    /// The result is presentation metadata only. It is never used to paginate
    /// the document or generate line numbers.
    pub fn line_numbering_configuration(
        &self,
    ) -> Result<Option<crate::line_numbering::Configuration>> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(None), crate::line_numbering::parse)
    }

    /// Insert or replace document line-numbering configuration.
    ///
    /// This updates stored style metadata only. It never calculates page or
    /// line numbers.
    pub fn set_line_numbering_configuration(
        &mut self,
        configuration: &crate::line_numbering::Configuration,
    ) -> Result<Option<crate::line_numbering::Configuration>> {
        configuration.validate()?;
        let old = self.line_numbering_configuration()?;
        let styles = self
            .styles_xml
            .clone()
            .unwrap_or_else(Structure::default_styles_xml);
        self.styles_xml = Some(crate::line_numbering::set_xml(&styles, configuration)?);
        Ok(old)
    }

    /// Remove document line-numbering configuration and return its old value.
    pub fn clear_line_numbering_configuration(
        &mut self,
    ) -> Result<Option<crate::line_numbering::Configuration>> {
        let old = self.line_numbering_configuration()?;
        let Some(styles) = self.styles_xml.as_deref() else {
            return Ok(None);
        };
        self.styles_xml = Some(crate::line_numbering::remove_xml(styles)?);
        Ok(old)
    }

    /// Parse the document's master pages and current header/footer regions.
    pub fn master_pages(&self) -> Result<Vec<Master>> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(Vec::new()), read)
    }

    /// Parse automatic page layouts, their properties, and header/footer styles.
    pub fn page_layouts(&self) -> Result<Vec<PageLayout>> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(Vec::new()), parse_page_layouts)
    }

    /// Replace one page layout with a complete XML fragment.
    ///
    /// The fragment must be exactly one self-contained `style:page-layout`
    /// element whose `style:name` matches `page_layout_name`. This supports all
    /// page properties and nested header/footer styles while preserving every
    /// unrelated byte in `styles.xml`.
    pub fn set_page_layout_xml(
        &mut self,
        page_layout_name: &str,
        page_layout_xml: &str,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml page layout to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_page_layout_xml(
            styles,
            page_layout_name,
            page_layout_xml,
        )?);
        Ok(())
    }

    /// Create or replace typed header/footer properties in one page layout.
    pub fn set_page_layout_header_footer_properties(
        &mut self,
        page_layout_name: &str,
        region: crate::header_footer::properties::Region,
        properties: &crate::header_footer::properties::StyleProperties,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml page layout to modify".to_string(),
            )
        })?;
        let layouts = parse_page_layouts(styles)?;
        let layout = layouts
            .iter()
            .find(|layout| layout.name == page_layout_name)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "page layout '{page_layout_name}' does not exist"
                ))
            })?;
        let replacement = crate::header_footer::properties::replace_page_layout_region_properties(
            layout, region, properties,
        )?;
        self.styles_xml = Some(set_page_layout_xml(styles, page_layout_name, &replacement)?);
        Ok(())
    }

    /// Create or replace typed columns in one existing page layout.
    pub fn set_page_layout_columns(
        &mut self,
        page_layout_name: &str,
        columns: &crate::style::columns::Columns,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml page layout to modify".to_string(),
            )
        })?;
        let layouts = parse_page_layouts(styles)?;
        let layout = layouts
            .iter()
            .find(|layout| layout.name == page_layout_name)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "page layout '{page_layout_name}' does not exist"
                ))
            })?;
        let replacement = crate::style::columns::replace_page_layout_columns(layout, columns)?;
        self.styles_xml = Some(set_page_layout_xml(styles, page_layout_name, &replacement)?);
        Ok(())
    }

    /// Create or replace the typed footnote separator in one existing page layout.
    pub fn set_page_layout_footnote_separator(
        &mut self,
        page_layout_name: &str,
        separator: &crate::footnote_separator::Separator,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml page layout to modify".to_string(),
            )
        })?;
        let layouts = parse_page_layouts(styles)?;
        let layout = layouts
            .iter()
            .find(|layout| layout.name == page_layout_name)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "page layout '{page_layout_name}' does not exist"
                ))
            })?;
        let replacement =
            crate::footnote_separator::replace_page_layout_footnote_separator(layout, separator)?;
        self.styles_xml = Some(set_page_layout_xml(styles, page_layout_name, &replacement)?);
        Ok(())
    }

    /// Add an empty master page and its referenced page layout.
    /// Replace one existing named list level's modern label alignment.
    pub fn set_list_level_label_alignment(
        &mut self,
        item: &crate::list_label_alignment::Style,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml list style to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(crate::list_label_alignment::set_xml(styles, item)?);
        Ok(())
    }

    /// Add an empty master page and its referenced page layout.
    /// Replace, insert, or remove one existing paragraph style's direct drop cap.
    pub fn set_paragraph_style_drop_cap(
        &mut self,
        style: &crate::style::paragraph::drop_cap::Style,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml paragraph style to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(crate::style::paragraph::drop_cap::set_xml(styles, style)?);
        Ok(())
    }

    /// Replace, insert, or remove typed row properties on an existing table-row style.
    pub fn set_table_row_style_properties(
        &mut self,
        style: &crate::style::table::row::Style,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml table-row style to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(crate::style::table::row::set_xml(styles, style)?);
        Ok(())
    }

    /// Replace, insert, or remove typed properties on an existing table style.
    pub fn set_table_style_properties(
        &mut self,
        style: &crate::style::table::table::Style,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml table style to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(crate::style::table::table::set_xml(styles, style)?);
        Ok(())
    }

    /// Add an empty master page and its referenced page layout.
    ///
    /// A minimal page layout is created in `office:automatic-styles` when a
    /// layout with `page_layout_name` does not already exist.
    pub fn add_master_page(&mut self, name: &str, page_layout_name: &str) -> Result<()> {
        let styles = self
            .styles_xml
            .get_or_insert_with(Structure::default_styles_xml);
        *styles = add(styles, name, page_layout_name)?;
        Ok(())
    }

    /// Insert a complete typed master page without rewriting unrelated styles.
    pub fn insert_master_page(&mut self, page: &Master) -> Result<()> {
        let fragment = page.to_xml_fragment()?;
        let styles = self
            .styles_xml
            .get_or_insert_with(Structure::default_styles_xml);
        *styles = insert(styles, &fragment)?;
        Ok(())
    }

    /// Replace one named master page without rewriting unrelated styles.
    pub fn replace_master_page(&mut self, name: &str, page: &Master) -> Result<()> {
        let fragment = page.to_xml_fragment()?;
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(replace(styles, name, &fragment)?);
        Ok(())
    }

    /// Remove one named master page without rewriting unrelated styles.
    pub fn remove_master_page(&mut self, name: &str) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(remove(styles, name)?);
        Ok(())
    }

    /// Set plain text in one header/footer region of an existing master page.
    ///
    /// Only the selected region is rewritten; all unrelated style XML is preserved.
    pub fn set_header_footer_text(
        &mut self,
        master_page_name: &str,
        kind: crate::header_footer::Kind,
        text: &str,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_text(styles, master_page_name, kind, Some(text))?);
        Ok(())
    }

    /// Replace one header/footer region with a complete XML fragment.
    ///
    /// The fragment must be exactly one self-contained `style:header`,
    /// `style:footer`, or corresponding first/left variant matching `kind`.
    /// This preserves rich text, fields, tables, drawings, and extension content.
    pub fn set_header_footer_xml(
        &mut self,
        master_page_name: &str,
        kind: crate::header_footer::Kind,
        xml: &str,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_xml(styles, master_page_name, kind, xml)?);
        Ok(())
    }

    /// Remove one header/footer region from an existing master page.
    pub fn clear_header_footer(
        &mut self,
        master_page_name: &str,
        kind: crate::header_footer::Kind,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_text(styles, master_page_name, kind, None)?);
        Ok(())
    }
}
