//! XML generation for the builder's ODF document parts.

use super::model::{AnnotationInsertion, Builder, DocumentElement};
use litchi_core::{Error, Result, xml::escape_xml};

impl Builder {
    /// Generate the content.xml body
    pub(super) fn generate_content_body(&self) -> Result<String> {
        let mut estimated = 256usize;
        estimated += self.elements.len() * 96;
        estimated += self
            .elements
            .iter()
            .map(|e| match e {
                DocumentElement::Paragraph(p) => p.text().map_or(0, |t| t.len()),
                DocumentElement::Heading(h) => h.text().map_or(0, |t| t.len()),
                DocumentElement::Table(_) | DocumentElement::List(_) => 256,
                DocumentElement::Section(xml) => xml.len(),
            })
            .sum::<usize>();

        let mut body = String::with_capacity(estimated);

        if let Some(sequence) = &self.page_sequence {
            body.push_str(&sequence.to_xml_fragment()?);
        }

        if !self.property_forms.is_empty()
            || !self.control_forms.is_empty()
            || !self.interactive_forms.is_empty()
            || !self.selection_forms.is_empty()
            || !self.visual_forms.is_empty()
            || !self.generic_forms.is_empty()
            || !self.password_file_forms.is_empty()
            || !self.image_frame_forms.is_empty()
            || !self.value_range_forms.is_empty()
            || !self.typed_value_forms.is_empty()
            || !self.grid_forms.is_empty()
            || !self.connection_resource_forms.is_empty()
        {
            body.push_str("<office:forms>");
            for form in &self.property_forms {
                body.push_str(&form.to_xml_fragment()?);
            }
            for form in &self.control_forms {
                body.push_str(&form.to_xml_fragment()?);
            }
            for form in &self.interactive_forms {
                body.push_str(&form.to_xml_fragment()?);
            }
            for form in &self.selection_forms {
                body.push_str(&form.to_xml_fragment()?);
            }
            for form in &self.visual_forms {
                body.push_str(&form.to_xml_fragment()?);
            }
            for form in &self.generic_forms {
                body.push_str(&form.to_xml_fragment()?);
            }
            for form in &self.password_file_forms {
                body.push_str(&form.to_xml_fragment()?);
            }
            for form in &self.image_frame_forms {
                body.push_str(&form.to_xml_fragment()?);
            }
            for form in &self.value_range_forms {
                body.push_str(&form.to_xml_fragment()?);
            }
            for form in &self.typed_value_forms {
                body.push_str(&form.to_xml_fragment()?);
            }
            for form in &self.grid_forms {
                body.push_str(&form.to_xml_fragment()?);
            }
            for form in &self.connection_resource_forms {
                body.push_str(&form.to_xml_fragment()?);
            }
            body.push_str("</office:forms>");
        }

        // Add all elements in order they were added
        for element in &self.elements {
            match element {
                DocumentElement::Paragraph(para) => {
                    let elem: crate::elements::element::Element = para.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Heading(heading) => {
                    let elem: crate::elements::element::Element = heading.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Table(table) => {
                    let elem: crate::elements::element::Element = table.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::List(list) => {
                    let elem: crate::elements::element::Element = list.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Section(section) => body.push_str(section),
            }
        }

        for index in &self.text_indexes {
            body.push_str(index);
        }

        if !self.text_index_marks.is_empty() {
            let prefix = r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#;
            let suffix = "</office:text>";
            let mut wrapped = format!("{prefix}{body}{suffix}");
            for (paragraph_index, mark) in &self.text_index_marks {
                wrapped = crate::insert_text_index_mark_xml(&wrapped, *paragraph_index, mark)?;
            }
            body = unwrap_text_body(&wrapped, prefix, suffix)?;
        }

        if !self.reference_marks.is_empty() {
            let prefix = r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#;
            let suffix = "</office:text>";
            let mut wrapped = format!("{prefix}{body}{suffix}");
            for (paragraph_index, mark) in &self.reference_marks {
                wrapped = crate::insert_reference_mark_xml(&wrapped, *paragraph_index, mark)?;
            }
            body = unwrap_text_body(&wrapped, prefix, suffix)?;
        }

        if !self.bookmark_targets.is_empty() {
            let prefix = r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#;
            let suffix = "</office:text>";
            let mut wrapped = format!("{prefix}{body}{suffix}");
            for (paragraph_index, target) in &self.bookmark_targets {
                wrapped = crate::insert_bookmark_xml(&wrapped, *paragraph_index, target)?;
            }
            body = unwrap_text_body(&wrapped, prefix, suffix)?;
        }

        if !self.ruby_annotations.is_empty() {
            let prefix = r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#;
            let suffix = "</office:text>";
            let mut wrapped = format!("{prefix}{body}{suffix}");
            for insertion in &self.ruby_annotations {
                wrapped = match insertion {
                    AnnotationInsertion::Append {
                        paragraph_index,
                        annotation,
                    } => crate::insert_ruby_annotation_xml(&wrapped, *paragraph_index, annotation)?,
                    AnnotationInsertion::Wrap {
                        paragraph_index,
                        range,
                        annotation,
                    } => crate::wrap_ruby_annotation_xml(
                        &wrapped,
                        *paragraph_index,
                        range.clone(),
                        annotation,
                    )?,
                };
            }
            body = unwrap_text_body(&wrapped, prefix, suffix)?;
        }

        if !self.notes.is_empty() {
            let prefix = r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#;
            let suffix = "</office:text>";
            let mut wrapped = format!("{prefix}{body}{suffix}");
            for (paragraph_index, note) in &self.notes {
                wrapped = crate::insert_note_xml(&wrapped, *paragraph_index, note)?;
            }
            body = unwrap_text_body(&wrapped, prefix, suffix)?;
        }

        Ok(body)
    }

    /// Generate the complete content.xml
    pub(super) fn generate_content_xml(&self) -> Result<String> {
        let body = self.generate_content_body()?;

        Ok(format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><office:text>{body}</office:text></office:body></office:document-content>"#
        ))
    }

    /// Generate meta.xml with metadata
    pub(super) fn generate_meta_xml(&self) -> String {
        let mut meta = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator>"#.to_string();

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            push_metadata_element(&mut meta, "dc:title", title);
        }

        if let Some(ref author) = self.metadata.author {
            push_metadata_element(&mut meta, "dc:creator", author);
        }

        if let Some(ref subject) = self.metadata.subject {
            push_metadata_element(&mut meta, "dc:subject", subject);
        }

        if let Some(ref description) = self.metadata.description {
            push_metadata_element(&mut meta, "dc:description", description);
        }

        if let Some(ref keywords) = self.metadata.keywords {
            push_metadata_element(&mut meta, "meta:keyword", keywords);
        }

        if let Some(ref identifier) = self.metadata.identifier {
            push_metadata_element(&mut meta, "dc:identifier", identifier);
        }

        if let Some(ref language) = self.metadata.language {
            push_metadata_element(&mut meta, "dc:language", language);
        }

        if let Some(created) = &self.metadata.created {
            meta.push_str("<meta:creation-date>");
            meta.push_str(&created.to_rfc3339_opts(chrono::SecondsFormat::AutoSi, true));
            meta.push_str("</meta:creation-date>");
        } else if let Some(created) = &self.metadata.created_local {
            meta.push_str("<meta:creation-date>");
            meta.push_str(&created.format("%Y-%m-%dT%H:%M:%S%.f").to_string());
            meta.push_str("</meta:creation-date>");
        }

        if let Some(modified) = &self.metadata.modified {
            meta.push_str("<dc:date>");
            meta.push_str(&modified.to_rfc3339_opts(chrono::SecondsFormat::AutoSi, true));
            meta.push_str("</dc:date>");
        } else if let Some(modified) = &self.metadata.modified_local {
            meta.push_str("<dc:date>");
            meta.push_str(&modified.format("%Y-%m-%dT%H:%M:%S%.f").to_string());
            meta.push_str("</dc:date>");
        }

        meta.push_str("</office:meta>");
        meta.push_str("</office:document-meta>");

        meta
    }

    /// Generate styles.xml with list styles
    pub(super) fn generate_styles_xml(&self) -> Result<String> {
        let mut xml = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" office:version="1.3"><office:font-face-decls/><office:styles><!-- Numbered list style --><text:list-style style:name="L1"><text:list-level-style-number text:level="1" text:style-name="Numbering_20_Symbols" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.27cm" fo:text-indent="-0.635cm" fo:margin-left="1.27cm"/></style:list-level-properties></text:list-level-style-number><text:list-level-style-number text:level="2" text:style-name="Numbering_20_Symbols" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.905cm" fo:text-indent="-0.635cm" fo:margin-left="1.905cm"/></style:list-level-properties></text:list-level-style-number><text:list-level-style-number text:level="3" text:style-name="Numbering_20_Symbols" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="2.54cm" fo:text-indent="-0.635cm" fo:margin-left="2.54cm"/></style:list-level-properties></text:list-level-style-number></text:list-style></office:styles><office:automatic-styles/><office:master-styles/></office:document-styles>"#.to_string();
        if !self.paragraph_tab_styles.is_empty() {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self
                .paragraph_tab_styles
                .iter()
                .map(|style| {
                    self.paragraph_drop_cap_styles
                        .iter()
                        .find(|cap| {
                            crate::style::paragraph::drop_cap::same_style_identity(cap, style)
                        })
                        .map_or_else(
                            || style.to_xml_fragment(),
                            |cap| {
                                crate::style::paragraph::drop_cap::merge_with_tab_style(style, cap)
                            },
                        )
                })
                .collect::<Result<String>>()?;
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_drop_cap_styles.is_empty() {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self
                .paragraph_drop_cap_styles
                .iter()
                .filter(|cap| {
                    !self.paragraph_tab_styles.iter().any(|tabs| {
                        crate::style::paragraph::drop_cap::same_style_identity(cap, tabs)
                    })
                })
                .map(|style| style.to_xml_fragment())
                .collect::<Result<String>>()?;
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_flow_styles.is_empty() {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self
                .paragraph_flow_styles
                .iter()
                .map(|x| x.to_xml_fragment())
                .collect::<Result<String>>()?;
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_margin_styles.is_empty() {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self
                .paragraph_margin_styles
                .iter()
                .map(|x| x.to_xml_fragment())
                .collect::<Result<String>>()?;
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_border_styles.is_empty() {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self
                .paragraph_border_styles
                .iter()
                .map(|x| x.to_xml_fragment())
                .collect::<Result<String>>()?;
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_alignment_styles.is_empty() {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self
                .paragraph_alignment_styles
                .iter()
                .map(|x| x.to_xml_fragment())
                .collect::<Result<String>>()?;
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_break_styles.is_empty() {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self
                .paragraph_break_styles
                .iter()
                .map(|x| x.to_xml_fragment())
                .collect::<Result<String>>()?;
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_writing_mode_styles.is_empty() {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self
                .paragraph_writing_mode_styles
                .iter()
                .map(|x| x.to_xml_fragment())
                .collect::<Result<String>>()?;
            xml.insert_str(insertion, &fragments);
        }
        if !self.table_row_property_styles.is_empty() {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self
                .table_row_property_styles
                .iter()
                .map(|x| x.to_xml_fragment())
                .collect::<Result<String>>()?;
            xml.insert_str(insertion, &fragments);
        }
        if !self.table_property_styles.is_empty() {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self
                .table_property_styles
                .iter()
                .map(|x| x.to_xml_fragment())
                .collect::<Result<String>>()?;
            xml.insert_str(insertion, &fragments);
        }
        if !self.table_column_property_styles.is_empty() {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self
                .table_column_property_styles
                .iter()
                .map(|x| x.to_xml_fragment())
                .collect::<Result<String>>()?;
            xml.insert_str(insertion, &fragments);
        }
        if !self.table_cell_property_styles.is_empty() {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self
                .table_cell_property_styles
                .iter()
                .map(|x| x.to_xml_fragment())
                .collect::<Result<String>>()?;
            xml.insert_str(insertion, &fragments);
        }
        if !self.section_property_styles.is_empty() {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self
                .section_property_styles
                .iter()
                .map(|x| x.to_xml_fragment())
                .collect::<Result<String>>()?;
            xml.insert_str(insertion, &fragments);
        }
        if let Some(configuration) = &self.line_numbering_configuration {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragment = configuration.to_xml()?;
            xml.insert_str(insertion, &fragment);
        }
        if self.notes_configurations.footnote.is_some()
            || self.notes_configurations.endnote.is_some()
        {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self.notes_configurations.to_xml_fragment()?;
            xml.insert_str(insertion, &fragments);
        }
        if !self.ruby_styles.is_empty() {
            let insertion = find_required(&xml, "</office:styles>")?;
            let fragments = self
                .ruby_styles
                .iter()
                .map(|style| style.to_xml_fragment())
                .collect::<Result<String>>()?;
            xml.insert_str(insertion, &fragments);
        }
        if !self.page_layout_columns.is_empty()
            || !self.page_layout_footnote_separators.is_empty()
            || !self.page_layout_header_footer_properties.is_empty()
        {
            let mut fragments = self
                .page_layout_columns
                .iter()
                .map(|(name, columns)| {
                    let mut fragment = columns.to_page_layout_fragment(name)?;
                    if let Some((_, separator)) = self
                        .page_layout_footnote_separators
                        .iter()
                        .find(|(separator_name, _)| separator_name == name)
                    {
                        let insertion =
                            find_required(&fragment, "</style:page-layout-properties>")?;
                        fragment.insert_str(insertion, &separator.to_xml_fragment()?);
                    }
                    for (_, region, properties) in self
                        .page_layout_header_footer_properties
                        .iter()
                        .filter(|(property_name, _, _)| property_name == name)
                    {
                        let insertion = rfind_required(&fragment, "</style:page-layout>")?;
                        fragment.insert_str(insertion, &properties.to_region_fragment(*region)?);
                    }
                    Ok(fragment)
                })
                .collect::<Result<String>>()?;
            for (name, separator) in &self.page_layout_footnote_separators {
                if !self
                    .page_layout_columns
                    .iter()
                    .any(|(column_name, _)| column_name == name)
                {
                    let mut fragment = separator.to_page_layout_fragment(name)?;
                    for (_, region, properties) in self
                        .page_layout_header_footer_properties
                        .iter()
                        .filter(|(property_name, _, _)| property_name == name)
                    {
                        let insertion = rfind_required(&fragment, "</style:page-layout>")?;
                        fragment.insert_str(insertion, &properties.to_region_fragment(*region)?);
                    }
                    fragments.push_str(&fragment);
                }
            }
            for (index, (name, _, _)) in
                self.page_layout_header_footer_properties.iter().enumerate()
            {
                if self.page_layout_columns.iter().any(|(n, _)| n == name)
                    || self
                        .page_layout_footnote_separators
                        .iter()
                        .any(|(n, _)| n == name)
                    || self.page_layout_header_footer_properties[..index]
                        .iter()
                        .any(|(n, _, _)| n == name)
                {
                    continue;
                }
                let mut fragment =
                    format!("<style:page-layout style:name=\"{}\">", escape_xml(name));
                for (_, region, properties) in self
                    .page_layout_header_footer_properties
                    .iter()
                    .filter(|(n, _, _)| n == name)
                {
                    fragment.push_str(&properties.to_region_fragment(*region)?);
                }
                fragment.push_str("</style:page-layout>");
                fragments.push_str(&fragment);
            }
            xml = xml.replacen(
                "<office:automatic-styles/>",
                &format!("<office:automatic-styles>{fragments}</office:automatic-styles>"),
                1,
            );
        }
        for alignment in &self.list_level_label_alignments {
            xml = crate::list_label_alignment::set_xml(&xml, alignment)?;
        }
        Ok(xml)
    }
}

fn push_metadata_element(output: &mut String, name: &str, value: &str) {
    output.push('<');
    output.push_str(name);
    output.push('>');
    output.push_str(&escape_xml(value));
    output.push_str("</");
    output.push_str(name);
    output.push('>');
}

fn find_required(xml: &str, delimiter: &str) -> Result<usize> {
    xml.find(delimiter).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "generated styles XML is missing required delimiter {delimiter}"
        ))
    })
}

fn rfind_required(xml: &str, delimiter: &str) -> Result<usize> {
    xml.rfind(delimiter).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "generated styles XML is missing required delimiter {delimiter}"
        ))
    })
}

fn unwrap_text_body(wrapped: &str, prefix: &str, suffix: &str) -> Result<String> {
    wrapped
        .strip_prefix(prefix)
        .and_then(|value| value.strip_suffix(suffix))
        .map(str::to_owned)
        .ok_or_else(|| Error::InvalidFormat("generated ODT text wrapper is malformed".to_string()))
}
