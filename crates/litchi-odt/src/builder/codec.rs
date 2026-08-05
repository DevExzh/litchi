//! XML generation for the builder's ODF document parts.

use super::model::{AnnotationInsertion, Builder, DocumentElement};
use litchi_core::xml::escape_xml;

impl Builder {
    /// Generate the content.xml body
    pub(super) fn generate_content_body(&self) -> String {
        let mut estimated = 256usize;
        estimated += self.elements.len() * 96;
        estimated += self
            .elements
            .iter()
            .map(|e| match e {
                DocumentElement::Paragraph(p) => p.text().map(|t| t.len()).unwrap_or(0),
                DocumentElement::Heading(h) => h.text().map(|t| t.len()).unwrap_or(0),
                DocumentElement::Table(_) => 256,
                DocumentElement::List(_) => 256,
                DocumentElement::Section(xml) => xml.len(),
            })
            .sum::<usize>();

        let mut body = String::with_capacity(estimated);

        if let Some(sequence) = &self.page_sequence {
            body.push_str(
                &sequence
                    .to_xml_fragment()
                    .expect("validated builder page sequence"),
            );
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
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder property form"),
                );
            }
            for form in &self.control_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder control form"),
                );
            }
            for form in &self.interactive_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder interactive form"),
                );
            }
            for form in &self.selection_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder selection form"),
                );
            }
            for form in &self.visual_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder visual form"),
                );
            }
            for form in &self.generic_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder generic form"),
                );
            }
            for form in &self.password_file_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder password/file form"),
                );
            }
            for form in &self.image_frame_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder image-frame form"),
                );
            }
            for form in &self.value_range_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder value-range form"),
                );
            }
            for form in &self.typed_value_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder typed-value form"),
                );
            }
            for form in &self.grid_forms {
                body.push_str(&form.to_xml_fragment().expect("validated builder grid form"));
            }
            for form in &self.connection_resource_forms {
                body.push_str(
                    &form
                        .to_xml_fragment()
                        .expect("validated builder connection-resource form"),
                );
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
                wrapped = crate::insert_text_index_mark_xml(&wrapped, *paragraph_index, mark)
                    .expect("validated builder index mark");
            }
            body = wrapped[prefix.len()..wrapped.len() - suffix.len()].to_string();
        }

        if !self.reference_marks.is_empty() {
            let prefix = r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#;
            let suffix = "</office:text>";
            let mut wrapped = format!("{prefix}{body}{suffix}");
            for (paragraph_index, mark) in &self.reference_marks {
                wrapped = crate::insert_reference_mark_xml(&wrapped, *paragraph_index, mark)
                    .expect("validated builder reference mark");
            }
            body = wrapped[prefix.len()..wrapped.len() - suffix.len()].to_string();
        }

        if !self.bookmark_targets.is_empty() {
            let prefix = r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#;
            let suffix = "</office:text>";
            let mut wrapped = format!("{prefix}{body}{suffix}");
            for (paragraph_index, target) in &self.bookmark_targets {
                wrapped = crate::insert_bookmark_xml(&wrapped, *paragraph_index, target)
                    .expect("validated builder bookmark target");
            }
            body = wrapped[prefix.len()..wrapped.len() - suffix.len()].to_string();
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
                    } => crate::insert_ruby_annotation_xml(&wrapped, *paragraph_index, annotation)
                        .expect("validated builder ruby annotation"),
                    AnnotationInsertion::Wrap {
                        paragraph_index,
                        range,
                        annotation,
                    } => crate::wrap_ruby_annotation_xml(
                        &wrapped,
                        *paragraph_index,
                        range.clone(),
                        annotation,
                    )
                    .expect("validated builder ruby annotation range"),
                };
            }
            body = wrapped[prefix.len()..wrapped.len() - suffix.len()].to_string();
        }

        if !self.notes.is_empty() {
            let prefix = r#"<office:text xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#;
            let suffix = "</office:text>";
            let mut wrapped = format!("{prefix}{body}{suffix}");
            for (paragraph_index, note) in &self.notes {
                wrapped = crate::insert_note_xml(&wrapped, *paragraph_index, note)
                    .expect("validated builder note");
            }
            body = wrapped[prefix.len()..wrapped.len() - suffix.len()].to_string();
        }

        body
    }

    /// Generate the complete content.xml
    pub(super) fn generate_content_xml(&self) -> String {
        let body = self.generate_content_body();

        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><office:text>{}</office:text></office:body></office:document-content>"#,
            body
        )
    }

    /// Generate meta.xml with metadata
    pub(super) fn generate_meta_xml(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();

        let mut meta = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator><meta:creation-date>{}</meta:creation-date><dc:date>{}</dc:date>"#,
            now, now
        );

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            meta.push_str(&format!("<dc:title>{}</dc:title>", escape_xml(title)));
        }

        if let Some(ref author) = self.metadata.author {
            meta.push_str(&format!("<dc:creator>{}</dc:creator>", escape_xml(author)));
        }

        if let Some(ref subject) = self.metadata.subject {
            meta.push_str(&format!("<dc:subject>{}</dc:subject>", escape_xml(subject)));
        }

        if let Some(ref description) = self.metadata.description {
            meta.push_str(&format!(
                "<dc:description>{}</dc:description>",
                escape_xml(description)
            ));
        }

        if let Some(ref keywords) = self.metadata.keywords {
            meta.push_str(&format!(
                "<meta:keyword>{}</meta:keyword>",
                escape_xml(keywords)
            ));
        }

        meta.push_str("</office:meta>");
        meta.push_str("</office:document-meta>");

        meta
    }

    /// Generate styles.xml with list styles
    pub(super) fn generate_styles_xml(&self) -> String {
        let mut xml = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" office:version="1.3"><office:font-face-decls/><office:styles><!-- Numbered list style --><text:list-style style:name="L1"><text:list-level-style-number text:level="1" text:style-name="Numbering_20_Symbols" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.27cm" fo:text-indent="-0.635cm" fo:margin-left="1.27cm"/></style:list-level-properties></text:list-level-style-number><text:list-level-style-number text:level="2" text:style-name="Numbering_20_Symbols" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.905cm" fo:text-indent="-0.635cm" fo:margin-left="1.905cm"/></style:list-level-properties></text:list-level-style-number><text:list-level-style-number text:level="3" text:style-name="Numbering_20_Symbols" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="2.54cm" fo:text-indent="-0.635cm" fo:margin-left="2.54cm"/></style:list-level-properties></text:list-level-style-number></text:list-style></office:styles><office:automatic-styles/><office:master-styles/></office:document-styles>"#.to_string();
        if !self.paragraph_tab_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
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
                            || {
                                style
                                    .to_xml_fragment()
                                    .expect("validated paragraph tab style")
                            },
                            |cap| {
                                crate::style::paragraph::drop_cap::merge_with_tab_style(style, cap)
                                    .expect("validated merged paragraph style")
                            },
                        )
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_drop_cap_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .paragraph_drop_cap_styles
                .iter()
                .filter(|cap| {
                    !self.paragraph_tab_styles.iter().any(|tabs| {
                        crate::style::paragraph::drop_cap::same_style_identity(cap, tabs)
                    })
                })
                .map(|style| {
                    style
                        .to_xml_fragment()
                        .expect("validated paragraph drop-cap style")
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_flow_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .paragraph_flow_styles
                .iter()
                .map(|x| x.to_xml_fragment().expect("validated paragraph flow style"))
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_margin_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .paragraph_margin_styles
                .iter()
                .map(|x| {
                    x.to_xml_fragment()
                        .expect("validated paragraph margin style")
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_border_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .paragraph_border_styles
                .iter()
                .map(|x| {
                    x.to_xml_fragment()
                        .expect("validated paragraph border style")
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_alignment_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .paragraph_alignment_styles
                .iter()
                .map(|x| {
                    x.to_xml_fragment()
                        .expect("validated paragraph alignment style")
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_break_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .paragraph_break_styles
                .iter()
                .map(|x| {
                    x.to_xml_fragment()
                        .expect("validated paragraph break style")
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.paragraph_writing_mode_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .paragraph_writing_mode_styles
                .iter()
                .map(|x| {
                    x.to_xml_fragment()
                        .expect("validated paragraph writing-mode style")
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.table_row_property_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .table_row_property_styles
                .iter()
                .map(|x| {
                    x.to_xml_fragment()
                        .expect("validated table-row property style")
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.table_property_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .table_property_styles
                .iter()
                .map(|x| x.to_xml_fragment().expect("validated table property style"))
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.table_column_property_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .table_column_property_styles
                .iter()
                .map(|x| {
                    x.to_xml_fragment()
                        .expect("validated table-column property style")
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.table_cell_property_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .table_cell_property_styles
                .iter()
                .map(|x| {
                    x.to_xml_fragment()
                        .expect("validated table-cell property style")
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if !self.section_property_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .section_property_styles
                .iter()
                .map(|x| {
                    x.to_xml_fragment()
                        .expect("validated section property style")
                })
                .collect::<String>();
            xml.insert_str(insertion, &fragments);
        }
        if let Some(configuration) = &self.line_numbering_configuration {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragment = configuration
                .to_xml()
                .expect("validated line-numbering configuration");
            xml.insert_str(insertion, &fragment);
        }
        if self.notes_configurations.footnote.is_some()
            || self.notes_configurations.endnote.is_some()
        {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .notes_configurations
                .to_xml_fragment()
                .expect("validated notes configurations");
            xml.insert_str(insertion, &fragments);
        }
        if !self.ruby_styles.is_empty() {
            let insertion = xml.find("</office:styles>").expect("static styles root");
            let fragments = self
                .ruby_styles
                .iter()
                .map(|style| style.to_xml_fragment().expect("validated ruby style"))
                .collect::<String>();
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
                    let mut fragment = columns
                        .to_page_layout_fragment(name)
                        .expect("validated column page layout");
                    if let Some((_, separator)) = self
                        .page_layout_footnote_separators
                        .iter()
                        .find(|(separator_name, _)| separator_name == name)
                    {
                        let insertion = fragment
                            .find("</style:page-layout-properties>")
                            .expect("static column page layout fragment");
                        fragment.insert_str(
                            insertion,
                            &separator
                                .to_xml_fragment()
                                .expect("validated footnote separator"),
                        );
                    }
                    for (_, region, properties) in self
                        .page_layout_header_footer_properties
                        .iter()
                        .filter(|(property_name, _, _)| property_name == name)
                    {
                        let insertion = fragment
                            .rfind("</style:page-layout>")
                            .expect("page layout fragment");
                        fragment.insert_str(
                            insertion,
                            &properties
                                .to_region_fragment(*region)
                                .expect("validated header/footer properties"),
                        );
                    }
                    fragment
                })
                .collect::<String>();
            for (name, separator) in &self.page_layout_footnote_separators {
                if !self
                    .page_layout_columns
                    .iter()
                    .any(|(column_name, _)| column_name == name)
                {
                    let mut fragment = separator
                        .to_page_layout_fragment(name)
                        .expect("validated footnote separator page layout");
                    for (_, region, properties) in self
                        .page_layout_header_footer_properties
                        .iter()
                        .filter(|(property_name, _, _)| property_name == name)
                    {
                        let insertion = fragment
                            .rfind("</style:page-layout>")
                            .expect("page layout fragment");
                        fragment.insert_str(
                            insertion,
                            &properties
                                .to_region_fragment(*region)
                                .expect("validated header/footer properties"),
                        );
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
                    fragment.push_str(
                        &properties
                            .to_region_fragment(*region)
                            .expect("validated header/footer properties"),
                    );
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
            xml = crate::list_label_alignment::set_xml(&xml, alignment)
                .expect("validated generated list alignment");
        }
        xml
    }
}
