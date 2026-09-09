use super::prelude::*;

fn parse_deferred_source<T>(
    source: &std::result::Result<Option<Vec<u8>>, String>,
    parser: impl Fn(&[u8]) -> Result<T>,
) -> std::result::Result<Option<T>, String> {
    match source {
        Ok(Some(bytes)) => parser(bytes).map(Some).map_err(|error| error.to_string()),
        Ok(None) => Ok(None),
        Err(error) => Err(error.clone()),
    }
}

impl Document {
    /// Get all text content from the document.
    ///
    /// This extracts all text from the document, concatenated together.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_doc::Package;
    ///
    /// let mut pkg = Package::open("document.doc")?;
    /// let doc = pkg.document()?;
    /// let text = doc.text()?;
    /// println!("{}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        self.text_extractor.extract_all_text()
    }

    /// Get the number of paragraphs in the document.
    ///
    /// This method counts the same logical paragraphs returned by [`Self::paragraphs`].
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_doc::Package;
    ///
    /// let mut pkg = Package::open("document.doc")?;
    /// let doc = pkg.document()?;
    /// let count = doc.paragraph_count()?;
    /// println!("Paragraphs: {}", count);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn paragraph_count(&self) -> Result<usize> {
        let text = self.text_extractor.text();
        count_paragraphs_in_ranges(
            text,
            self.fib
                .get_all_subdoc_ranges()
                .into_iter()
                .map(|(_, start_cp, end_cp)| (start_cp, end_cp)),
        )
    }

    /// Get the number of tables in the document.
    ///
    /// Counts top-level tables (`table_level` == 1) by scanning paragraph properties
    /// for table markers. Based on Apache POI's table detection algorithm.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_doc::Package;
    ///
    /// let mut pkg = Package::open("document.doc")?;
    /// let doc = pkg.document()?;
    /// let count = doc.table_count()?;
    /// println!("Tables: {}", count);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn table_count(&self) -> Result<usize> {
        // Count tables by iterating through paragraphs and tracking table boundaries
        // A new table starts when we encounter a paragraph with in_table=true and
        // table_level=1 after a paragraph that was not in a table or had a different level
        let paragraphs = self.paragraphs()?;
        let mut table_count = 0;
        let mut in_table_level_1 = false;

        for para in paragraphs {
            let props = para.properties();

            // Check if this paragraph is in a top-level table (level 1)
            if props.in_table && props.table_nesting_level == 1 {
                // If we weren't previously in a level-1 table, this is a new table
                if !in_table_level_1 {
                    table_count += 1;
                    in_table_level_1 = true;
                }
            } else {
                // We've exited the table
                in_table_level_1 = false;
            }
        }

        Ok(table_count)
    }

    /// Get access to the File Information Block.
    ///
    /// This provides lower-level access to document properties and structure.
    #[inline]
    #[must_use]
    pub fn fib(&self) -> &FileInformationBlock {
        &self.fib
    }

    /// Get the parsed Word 97+ stylesheet.
    ///
    /// Null fixed-index slots are retained, and each non-empty style exposes its
    /// exact UPX property payloads for subsequent inheritance and formatting.
    #[must_use]
    pub fn stylesheet(&self) -> Option<&StyleSheet> {
        self.stylesheet.as_ref()
    }

    /// Get the document's fixed associated-string metadata table.
    ///
    /// Template and mail-merge paths are inert strings and are never opened.
    #[must_use]
    pub fn associated_strings(&self) -> Option<&DocumentAssociatedStrings> {
        self.associated_strings.as_ref()
    }

    /// Get the ordered `LISTNUM` list-name metadata table.
    #[must_use]
    pub fn list_names(&self) -> Option<&ListNamesTable> {
        self.list_names.as_ref()
    }

    /// Get list-level template codes parallel to the document's list definitions.
    #[must_use]
    pub fn list_templates(&self) -> Option<&ListTemplateTable> {
        self.list_templates.as_ref()
    }

    /// Strictly access spelling and grammar proofing-state ranges.
    ///
    /// Parsing is deferred so nonconforming producer caches do not prevent the document's
    /// primary text from opening. Any malformed PLCF is reported when this metadata is requested.
    pub fn proofing_tables(&self) -> Result<&ProofingTables> {
        self.proofing_tables
            .as_ref()
            .map_err(|error| PackageError::Corrupted(format!("invalid proofing metadata: {error}")))
    }

    /// Strictly access current and legacy grammar-checker cookie tables.
    ///
    /// Parsing is deferred so nonconforming producer caches do not prevent the document's
    /// primary text from opening. Cookie payloads remain opaque and are never interpreted.
    pub fn grammar_cookie_tables(&self) -> Result<&GrammarCookieTables> {
        self.grammar_cookies.as_ref().map_err(|error| {
            PackageError::Corrupted(format!("invalid grammar cookie metadata: {error}"))
        })
    }

    /// Strictly access the deprecated table-character cache (`PlcfTch`).
    ///
    /// Parsing is deferred because Word itself is instructed to ignore this
    /// producer cache. The cache is exposed as metadata only and is never
    /// acted upon.
    pub fn table_character_cache(&self) -> Result<Option<&TableCharacterCache>> {
        self.table_char_cache
            .as_ref()
            .map(Option::as_ref)
            .map_err(|error| {
                PackageError::Corrupted(format!("invalid table character cache: {error}"))
            })
    }

    /// Strictly access the main and header textbox break tables.
    ///
    /// Parsing is deferred so malformed optional metadata does not prevent the
    /// document's primary text from opening. The version-specific `Tbkd` flag
    /// bits are producer caches and are never interpreted.
    pub fn textbox_break_tables(&self) -> Result<&TextBoxBreakTables> {
        self.textbox_breaks.as_ref().map_err(|error| {
            PackageError::Corrupted(format!("invalid textbox break metadata: {error}"))
        })
    }

    /// Strictly access Text Services Framework records and their GUID table.
    ///
    /// Parsing is deferred so malformed optional metadata does not prevent the
    /// document's primary text from opening. Service-provided payloads remain
    /// opaque and are never interpreted.
    pub fn text_services_tables(&self) -> Result<&TextServicesTables> {
        self.text_services.as_ref().map_err(|error| {
            PackageError::Corrupted(format!("invalid text services metadata: {error}"))
        })
    }

    /// Strictly access the ordered Word 97/2000 save history.
    ///
    /// Parsing is deferred because modern Word versions are instructed to ignore
    /// this legacy cache. Saved paths remain inert and are never opened or resolved.
    pub fn saved_by_table(&self) -> Result<&SavedByTable> {
        self.saved_by_table
            .as_ref()
            .map_err(|error| PackageError::Corrupted(format!("invalid saved-by metadata: {error}")))
    }

    /// Strictly access the inert main-document saved selection (`Selsf`).
    ///
    /// Selection state is a historical Word UI cache. The returned record is
    /// never applied to document navigation or host selection state, and its
    /// optional malformed bytes are reported only when this accessor is used.
    pub fn saved_selection(&self) -> Result<Option<&SavedSelection>> {
        match self.saved_selection.get_or_init(|| {
            let ccp_text = self.fib.get_main_doc_range().1;
            parse_deferred_source(&self.saved_selection_source, |source| {
                let selection = SavedSelection::parse_bytes(source)?;
                if selection.cp_first() > ccp_text {
                    return Err(PackageError::Corrupted(format!(
                        "Selsf cpFirst {} exceeds ccpText {ccp_text}",
                        selection.cp_first()
                    )));
                }
                if selection.cp_lim() > ccp_text {
                    return Err(PackageError::Corrupted(format!(
                        "Selsf cpLim {} exceeds ccpText {ccp_text}",
                        selection.cp_lim()
                    )));
                }
                Ok(selection)
            })
        }) {
            Ok(value) => Ok(value.as_ref()),
            Err(error) => Err(PackageError::Corrupted(format!(
                "invalid saved selection metadata: {error}"
            ))),
        }
    }

    /// Strictly access the inert paragraph-group property array (`PGPArray`).
    ///
    /// PGP entries describe paragraph margins, borders, and HTML-oriented
    /// block types. They are returned as metadata and are never applied to
    /// document paragraphs.
    pub fn paragraph_groups(&self) -> Result<Option<&PgpArray>> {
        match self.paragraph_groups.get_or_init(|| {
            parse_deferred_source(&self.paragraph_groups_source, PgpArray::parse_bytes)
        }) {
            Ok(value) => Ok(value.as_ref()),
            Err(error) => Err(PackageError::Corrupted(format!(
                "invalid paragraph-group metadata: {error}"
            ))),
        }
    }

    /// Strictly access inert frame-set and list-style records from `RgDofr`.
    ///
    /// The records are exposed as bounded metadata. Frame file names are not
    /// opened, frame layout is not rendered, and list styles are not applied
    /// to document paragraphs.
    pub fn dofr_records(&self) -> Result<Option<&DofrArray>> {
        match self.dofr_records.get_or_init(|| {
            parse_deferred_source(&self.dofr_records_source, DofrArray::parse_bytes)
        }) {
            Ok(value) => Ok(value.as_ref()),
            Err(error) => Err(PackageError::Corrupted(format!(
                "invalid RgDofr metadata: {error}"
            ))),
        }
    }

    /// Strictly access inert `PrDrvr`, `PrEnvPort`, and `PrEnvLand` metadata.
    ///
    /// Printer environment blocks are retained as bounded bytes; no printer is
    /// contacted and no print settings are applied.
    pub fn print_environment(&self) -> Result<Option<&DocumentPrintEnvironment>> {
        match self.print_environment.get_or_init(|| {
            let driver =
                parse_deferred_source(&self.print_driver_source, PrintDriver::parse_bytes)?;
            let portrait = parse_deferred_source(
                &self.print_environment_portrait_source,
                PrintEnvironment::parse_bytes,
            )?;
            let landscape = parse_deferred_source(
                &self.print_environment_landscape_source,
                PrintEnvironment::parse_bytes,
            )?;
            Ok(DocumentPrintEnvironment::from_parts(
                driver, portrait, landscape,
            ))
        }) {
            Ok(value) => Ok(value.as_ref()),
            Err(error) => Err(PackageError::Corrupted(format!(
                "invalid print-environment metadata: {error}"
            ))),
        }
    }

    /// Strictly access Word's inert VBA signature variables in `StwUser`.
    ///
    /// Parsing is deferred because this optional table is independent of the
    /// main text. Signature and certificate-store bytes remain opaque; no
    /// certificate trust is established and no VBA project is opened or
    /// executed.
    pub fn vba_signatures(&self) -> Result<Option<&DocumentVbaSignatures>> {
        match self.vba_signatures.get_or_init(|| {
            parse_deferred_source(
                &self.vba_signatures_source,
                DocumentVbaSignatures::parse_bytes,
            )
        }) {
            Ok(value) => Ok(value.as_ref()),
            Err(error) => Err(PackageError::Corrupted(format!(
                "invalid VBA signature metadata: {error}"
            ))),
        }
    }

    /// Strictly access the caption label and `AutoCaption` tables.
    ///
    /// Parsing is deferred so malformed optional metadata does not prevent the
    /// document's primary text from opening. Caption labels remain inert text
    /// and referenced OLE objects are never activated.
    pub fn caption_tables(&self) -> Result<&CaptionTables> {
        self.caption_tables
            .as_ref()
            .map_err(|error| PackageError::Corrupted(format!("invalid caption metadata: {error}")))
    }

    /// Strictly access the repair-bookmark tables recorded when Word repaired
    /// the document's bookmark pairs.
    ///
    /// Parsing is deferred so malformed optional metadata does not prevent the
    /// document's primary text from opening. Repair descriptions remain inert
    /// text; no repair is ever applied or reverted.
    pub fn repair_bookmarks(&self) -> Result<Option<&DocumentRepairBookmarks>> {
        self.repair_bookmarks
            .as_ref()
            .map(Option::as_ref)
            .map_err(|error| {
                PackageError::Corrupted(format!("invalid repair bookmark metadata: {error}"))
            })
    }

    /// Strictly access glossary-only `AutoText` and formatted `AutoCorrect` metadata.
    ///
    /// Ordinary documents return `None`. Parsing is deferred so malformed
    /// optional metadata does not prevent primary text access.
    pub fn glossary_metadata(&self) -> Result<Option<&GlossaryMetadata>> {
        self.glossary_metadata
            .as_ref()
            .map(Option::as_ref)
            .map_err(|error| PackageError::Corrupted(format!("invalid glossary metadata: {error}")))
    }

    /// Get one glossary entry's content without its structural final character.
    ///
    /// Word-compatible producers treat the last CP in each item range as the
    /// entry-ending paragraph mark. The stored text is returned passively; fields,
    /// links, objects, and macros are never evaluated or activated.
    pub fn glossary_item_text(&self, index: usize) -> Result<Option<&str>> {
        let Some(metadata) = self.glossary_metadata()? else {
            return Ok(None);
        };
        let Some(item) = metadata.items().get(index) else {
            return Ok(None);
        };
        Ok(Some(self.text_extractor.text_at_range(
            item.start_cp(),
            item.end_cp().saturating_sub(1),
        )))
    }

    /// Strictly access a template's secondary-FIB attached `AutoText` document.
    ///
    /// The returned content is passive and never evaluates fields, follows
    /// links, activates embedded objects, or resolves or executes macros.
    pub fn attached_glossary(&self) -> Result<Option<&AttachedGlossary>> {
        self.attached_glossary
            .as_ref()
            .map(Option::as_ref)
            .map_err(|error| PackageError::Corrupted(format!("invalid attached glossary: {error}")))
    }
}

fn count_paragraphs_in_ranges(
    text: &str,
    ranges: impl IntoIterator<Item = (u32, u32)>,
) -> Result<usize> {
    ranges
        .into_iter()
        .try_fold(0usize, |count, (start_cp, end_cp)| {
            if start_cp >= end_cp {
                return Ok(count);
            }

            count
                .checked_add(ParagraphExtractor::count_paragraphs_in_range(
                    text,
                    (start_cp, end_cp),
                ))
                .ok_or_else(|| {
                    PackageError::Corrupted(
                        "DOC paragraph count exceeds the addressable range".to_owned(),
                    )
                })
        })
}

#[cfg(test)]
mod tests {
    use super::count_paragraphs_in_ranges;

    #[test]
    fn paragraph_count_skips_equal_reversed_and_wrapped_ranges() {
        let text = "one\rtwo\r";
        let ranges = [(0, 8), (8, 8), (8, 7), (u32::MAX - 1, 2)];

        assert_eq!(count_paragraphs_in_ranges(text, ranges).unwrap(), 2);
    }
}
