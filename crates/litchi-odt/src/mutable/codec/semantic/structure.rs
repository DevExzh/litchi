//! Page sequence, declaration, and section semantic snapshot edits.

use super::super::super::model::MutableDocument;
use crate::page_sequence::{Sequence, parse_page_sequence, set_page_sequence_xml};
use crate::variable_declaration::{Declarations, Group, Kind, Part, Scope};
use litchi_core::Result;

impl MutableDocument {
    /// Return the explicit page-sequence master-page assignments, if authored.
    ///
    /// The model preserves only the ordered `text:master-page-name` values of
    /// the document's `text:page-sequence` (ODF 1.3 §5.3); litchi does not
    /// paginate or resolve the referenced master pages.
    pub fn page_sequence(&self) -> Result<Option<Sequence>> {
        self.with_content_xml(parse_page_sequence)
    }

    /// Set, replace, or clear the document's `text:page-sequence`.
    ///
    /// A new sequence is written as the first child of `office:text`, matching
    /// the element order of ODF 1.3 §5.1. Passing `None` removes an existing
    /// sequence and is a no-op when none exists. Master-page names are stored
    /// lexically and never resolved against `styles.xml`.
    pub fn set_page_sequence(&mut self, sequence: Option<&Sequence>) -> Result<()> {
        let updated = self.with_content_xml(|xml| set_page_sequence_xml(xml, sequence))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Return validated variable, user-field, sequence, and DDE declarations.
    pub fn variable_declarations(&self) -> Result<Declarations> {
        self.with_content_xml(|content| {
            if let Some(styles) = self.styles_xml.as_deref() {
                crate::variable_declaration::parse_parts(&[
                    (content, Part::Content),
                    (styles, Part::Styles),
                ])
            } else {
                crate::variable_declaration::parse_parts(&[(content, Part::Content)])
            }
        })
    }

    /// Atomically insert or replace one declaration container.
    ///
    /// Content and styles are reparsed together before commit, preserving all
    /// cross-part declaration and field-reference invariants.
    pub fn set_variable_declaration_group(&mut self, group: &Group) -> Result<Option<Group>> {
        let current = self.variable_declarations()?;
        let old = current
            .groups
            .iter()
            .find(|candidate| {
                candidate.part == group.part
                    && candidate.scope == group.scope
                    && candidate.kind == group.kind
            })
            .cloned();
        match group.part {
            Part::Content => {
                let updated =
                    self.with_content_xml(|xml| crate::variable_declaration::set_xml(xml, group))?;
                if let Some(styles) = self.styles_xml.as_deref() {
                    crate::variable_declaration::parse_parts(&[
                        (&updated, Part::Content),
                        (styles, Part::Styles),
                    ])?;
                } else {
                    crate::variable_declaration::parse_parts(&[(&updated, Part::Content)])?;
                }
                self.content_xml = Some(updated);
            },
            Part::Styles => {
                let styles = self.styles_xml.as_deref().ok_or_else(|| {
                    litchi_core::Error::InvalidFormat("styles.xml is absent".to_string())
                })?;
                let updated = crate::variable_declaration::set_xml(styles, group)?;
                self.with_content_xml(|content| {
                    crate::variable_declaration::parse_parts(&[
                        (content, Part::Content),
                        (&updated, Part::Styles),
                    ])
                })?;
                self.styles_xml = Some(updated);
            },
            Part::Flat => {
                return Err(litchi_core::Error::InvalidFormat(
                    "MutableDocument cannot edit flat-document declarations".to_string(),
                ));
            },
        }
        Ok(old)
    }

    /// Atomically remove one declaration container, returning its old value.
    ///
    /// Removal fails without mutation when any remaining field references a
    /// declaration from the removed container.
    pub fn remove_variable_declaration_group(
        &mut self,
        part: Part,
        scope: &Scope,
        kind: Kind,
    ) -> Result<Option<Group>> {
        let current = self.variable_declarations()?;
        let Some(old) = current
            .groups
            .iter()
            .find(|candidate| {
                candidate.part == part && &candidate.scope == scope && candidate.kind == kind
            })
            .cloned()
        else {
            return Ok(None);
        };
        match part {
            Part::Content => {
                let updated = self.with_content_xml(|xml| {
                    crate::variable_declaration::remove_xml(xml, scope, kind)
                })?;
                if let Some(styles) = self.styles_xml.as_deref() {
                    crate::variable_declaration::parse_parts(&[
                        (&updated, Part::Content),
                        (styles, Part::Styles),
                    ])?;
                } else {
                    crate::variable_declaration::parse_parts(&[(&updated, Part::Content)])?;
                }
                self.content_xml = Some(updated);
            },
            Part::Styles => {
                let styles = self.styles_xml.as_deref().ok_or_else(|| {
                    litchi_core::Error::InvalidFormat("styles.xml is absent".to_string())
                })?;
                let updated = crate::variable_declaration::remove_xml(styles, scope, kind)?;
                self.with_content_xml(|content| {
                    crate::variable_declaration::parse_parts(&[
                        (content, Part::Content),
                        (&updated, Part::Styles),
                    ])
                })?;
                self.styles_xml = Some(updated);
            },
            Part::Flat => {
                return Err(litchi_core::Error::InvalidFormat(
                    "MutableDocument cannot edit flat-document declarations".to_string(),
                ));
            },
        }
        Ok(Some(old))
    }

    /// Return typed nested sections from current authoritative content XML.
    pub fn sections(&self) -> Result<Vec<crate::Section>> {
        self.with_content_xml(crate::parser::Parser::parse_sections)
    }

    /// Append a complete typed section without rewriting existing body content.
    pub fn add_section(&mut self, section: &crate::Section) -> Result<()> {
        let updated = self.with_content_xml(|xml| crate::add_section_xml(xml, section))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Atomically update section metadata/source while preserving enclosed XML bytes.
    pub fn update_section(&mut self, name: &str, replacement: &crate::Section) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::update_section_xml(xml, name, replacement))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Delete a named section together with its enclosed content.
    pub fn remove_section(&mut self, name: &str) -> Result<()> {
        let updated = self.with_content_xml(|xml| crate::remove_section_xml(xml, name))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Remove one section wrapper/source while retaining all enclosed content bytes.
    pub fn unwrap_section(&mut self, name: &str) -> Result<()> {
        let updated = self.with_content_xml(|xml| crate::unwrap_section_xml(xml, name))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Remove every section wrapper/source while retaining mixed nested content.
    pub fn clear_sections(&mut self) -> Result<()> {
        let updated = self.with_content_xml(crate::clear_sections_xml)?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Wrap an inclusive stable block range in a typed section.
    pub fn wrap_section(
        &mut self,
        section: &crate::Section,
        start: crate::Block,
        end: crate::Block,
    ) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::wrap_section_xml(xml, section, &start, &end))?;
        self.content_xml = Some(updated);
        Ok(())
    }
}
