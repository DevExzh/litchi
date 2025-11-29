/// DocumentPart - the main document.xml part of a Word document.
use crate::ooxml::docx::paragraph::Paragraph;
use crate::ooxml::docx::table::Table;
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;
use smallvec::SmallVec;
use std::sync::Arc;

/// The main document part of a Word document.
///
/// This corresponds to the `/word/document.xml` part in the package.
/// It contains the main document content including paragraphs, tables,
/// sections, and other block-level elements.
pub struct DocumentPart<'a> {
    /// Reference to the underlying part
    part: &'a dyn Part,
}

impl<'a> DocumentPart<'a> {
    /// Create a DocumentPart from a Part.
    ///
    /// # Arguments
    ///
    /// * `part` - The part containing the document.xml content
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the shared Arc of XML bytes (zero-copy from Part).
    #[inline]
    fn get_xml_arc(&self) -> Arc<Vec<u8>> {
        self.part.blob_arc()
    }

    /// Get the XML bytes of the document.
    #[inline]
    pub fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Extract all paragraph text from the document.
    ///
    /// This performs a quick extraction of all text content by finding
    /// `<w:t>` elements in the XML.
    ///
    /// # Performance
    ///
    /// Uses `quick-xml` for efficient streaming XML parsing with pre-allocated
    /// buffer and unsafe string conversion for optimal performance.
    pub fn extract_text(&self) -> Result<String> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        // Pre-allocate with estimated capacity to reduce reallocations
        let estimated_capacity = self.xml_bytes().len() / 8; // Rough estimate for text content
        let mut result = String::with_capacity(estimated_capacity);
        let mut in_text_element = false;

        // Use read_event() for zero-copy parsing from slice
        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    // Check if this is a w:t element
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = true;
                    }
                },
                Ok(Event::Text(e)) if in_text_element => {
                    // Extract text content - use unsafe conversion for better performance
                    let text = unsafe { std::str::from_utf8_unchecked(e.as_ref()) };
                    result.push_str(text);
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(result)
    }

    /// Count the number of paragraphs in the document.
    ///
    /// Counts `<w:p>` elements in the document body.
    pub fn paragraph_count(&self) -> Result<usize> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut count = 0;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"p" {
                        count += 1;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(count)
    }

    /// Count the number of tables in the document.
    ///
    /// Counts `<w:tbl>` elements in the document body.
    pub fn table_count(&self) -> Result<usize> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut count = 0;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"tbl" {
                        count += 1;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(count)
    }

    /// Get all paragraphs in the document.
    ///
    /// Extracts all `<w:p>` elements from the document body.
    ///
    /// # Performance
    ///
    /// Uses streaming XML parsing with pre-allocated SmallVec for efficient
    /// storage of typically small paragraph collections. Optimized to minimize
    /// allocations via pre-sized reserves.
    pub fn paragraphs(&self) -> Result<SmallVec<[Paragraph; 32]>> {
        let xml_bytes = self.xml_bytes();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        // Estimate paragraph count
        let estimated = (xml_bytes.len() / 400).max(8);
        let mut paragraphs = SmallVec::with_capacity(estimated);
        let mut current_para_xml = Vec::with_capacity(4096);
        let mut in_para = false;
        let mut depth = 0u32;

        // Use read_event() for zero-copy parsing from slice
        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"p" && !in_para {
                        in_para = true;
                        depth = 1;
                        current_para_xml.clear();
                        write_start_tag(&mut current_para_xml, &e);
                    } else if in_para {
                        depth += 1;
                        write_start_tag_dynamic(&mut current_para_xml, &e);
                    }
                },
                Ok(Event::End(e)) => {
                    if in_para {
                        write_end_tag(&mut current_para_xml, e.name().as_ref());
                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"p" {
                            // Clone bytes and clear buffer (preserves capacity for next element)
                            let para_xml = current_para_xml.clone();
                            current_para_xml.clear();
                            paragraphs.push(Paragraph::new(para_xml));
                            in_para = false;
                        }
                    }
                },
                Ok(Event::Text(e)) if in_para => {
                    current_para_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::Empty(e)) if in_para => {
                    write_empty_tag(&mut current_para_xml, &e);
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(paragraphs)
    }

    /// Get all tables in the document.
    ///
    /// Extracts all `<w:tbl>` elements from the document body.
    ///
    /// # Performance
    ///
    /// Uses SmallVec for efficient storage of typically small table collections.
    /// Optimized to minimize allocations via pre-sized reserves.
    pub fn tables(&self) -> Result<SmallVec<[Table; 8]>> {
        let xml_bytes = self.xml_bytes();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        let mut tables = SmallVec::new();
        let mut current_table_xml = Vec::with_capacity(8192);
        let mut in_table = false;
        let mut depth = 0u32;

        // Use read_event() for zero-copy parsing from slice
        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"tbl" && !in_table {
                        in_table = true;
                        depth = 1;
                        current_table_xml.clear();
                        write_start_tag(&mut current_table_xml, &e);
                    } else if in_table {
                        depth += 1;
                        write_start_tag_dynamic(&mut current_table_xml, &e);
                    }
                },
                Ok(Event::End(e)) => {
                    if in_table {
                        write_end_tag(&mut current_table_xml, e.name().as_ref());
                        depth -= 1;
                        if depth == 0 && e.local_name().as_ref() == b"tbl" {
                            // Clone bytes and clear buffer (preserves capacity for next element)
                            let table_xml = current_table_xml.clone();
                            current_table_xml.clear();
                            tables.push(Table::new(table_xml));
                            in_table = false;
                        }
                    }
                },
                Ok(Event::Text(e)) if in_table => {
                    current_table_xml.extend_from_slice(e.as_ref());
                },
                Ok(Event::Empty(e)) if in_table => {
                    write_empty_tag(&mut current_table_xml, &e);
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(tables)
    }

    /// Get all document elements (paragraphs and tables) in document order.
    ///
    /// This method parses the XML once and extracts both paragraphs and tables,
    /// returning an ordered vector that preserves the document structure.
    /// This is more efficient than calling `paragraphs()` and `tables()` separately,
    /// and it maintains the correct order of elements for sequential processing.
    ///
    /// # Performance
    ///
    /// Uses a single-pass XML parser that extracts both `<w:p>` and `<w:tbl>` elements
    /// in document order, which is significantly faster than parsing the XML twice.
    ///
    /// # Performance
    ///
    /// Uses TRUE zero-copy parsing with NO Vec::push calls.
    /// All allocations happen upfront via pre-counting.
    pub fn elements(&self) -> Result<Vec<crate::document::DocumentElement>> {
        use crate::document::DocumentElement;

        // Get cached Arc (created once, shared across calls)
        let source_arc = self.get_xml_arc();
        let len = source_arc.len();

        // PASS 1: Count elements (no allocations)
        let count = count_elements(&source_arc);
        if count == 0 {
            return Ok(Vec::new());
        }

        // Single allocation: result vector with exact capacity
        let mut elements: Vec<DocumentElement> = Vec::with_capacity(count);

        // OPTIMIZATION: Pre-increment Arc refcount to avoid per-element atomic operations.
        // Convert Arc to raw pointer - this "forgets" the Arc without decrementing refcount.
        let arc_ptr = Arc::into_raw(source_arc);
        // Reborrow from the raw pointer for reading (valid for the Arc's lifetime)
        let xml_bytes: &[u8] = unsafe { &*arc_ptr };
        // Pre-increment refcount by (count - 1) since into_raw gave us one ownership.
        // SAFETY: arc_ptr came from Arc::into_raw, so it's valid.
        unsafe {
            for _ in 0..(count - 1) {
                Arc::increment_strong_count(arc_ptr);
            }
        }

        // PASS 2: Fill elements directly using unsafe set_len to avoid push
        let mut write_idx = 0usize;
        let mut i = 0usize;

        // Get raw pointer for direct writes
        let elements_ptr = elements.as_mut_ptr();

        while i < len && write_idx < count {
            let Some(tag_start) = memchr::memchr(b'<', &xml_bytes[i..]) else {
                break;
            };
            let tag_start = i + tag_start;

            // Check for <w:p
            if tag_start + 5 < len && &xml_bytes[tag_start..tag_start + 4] == b"<w:p" {
                let next_char = xml_bytes[tag_start + 4];
                if (next_char == b'>' || next_char == b' ' || next_char == b'/')
                    && let Some(end) = find_paragraph_end(&xml_bytes[tag_start..])
                {
                    let end_pos = tag_start + end;
                    let elem_len = (end_pos - tag_start) as u32;

                    // SAFETY: arc_ptr is valid; each from_raw consumes one refcount we pre-incremented
                    let arc_clone = unsafe { Arc::from_raw(arc_ptr) };

                    // Write directly to pre-allocated slot (no push)
                    unsafe {
                        std::ptr::write(
                            elements_ptr.add(write_idx),
                            DocumentElement::Paragraph(Box::new(crate::document::Paragraph::Docx(
                                Paragraph::from_arc_range(arc_clone, tag_start as u32, elem_len),
                            ))),
                        );
                    }
                    write_idx += 1;
                    i = end_pos;
                    continue;
                }
            }
            // Check for <w:tbl
            else if tag_start + 7 < len && &xml_bytes[tag_start..tag_start + 6] == b"<w:tbl" {
                let next_char = xml_bytes[tag_start + 6];
                if (next_char == b'>' || next_char == b' ')
                    && let Some(end) = find_table_end(&xml_bytes[tag_start..])
                {
                    let end_pos = tag_start + end;
                    let elem_len = (end_pos - tag_start) as u32;

                    // SAFETY: arc_ptr is valid; each from_raw consumes one refcount we pre-incremented
                    let arc_clone = unsafe { Arc::from_raw(arc_ptr) };

                    // Write directly to pre-allocated slot (no push)
                    unsafe {
                        std::ptr::write(
                            elements_ptr.add(write_idx),
                            DocumentElement::Table(Box::new(crate::document::Table::Docx(
                                Box::new(Table::from_arc_range(
                                    arc_clone,
                                    tag_start as u32,
                                    elem_len,
                                )),
                            ))),
                        );
                    }
                    write_idx += 1;
                    i = end_pos;
                    continue;
                }
            }

            i = tag_start + 1;
        }

        // Handle refcount mismatch: if we found fewer elements than counted,
        // we need to decrement the unused pre-incremented refcounts.
        // SAFETY: arc_ptr is valid, and we're decrementing refs we incremented but didn't use.
        if write_idx < count {
            unsafe {
                for _ in 0..(count - write_idx) {
                    Arc::decrement_strong_count(arc_ptr);
                }
            }
        }

        // Set final length (elements were written directly)
        unsafe {
            elements.set_len(write_idx);
        }

        Ok(elements)
    }
}

/// Write a start tag with raw bytes from BytesStart, ending with ">".
/// Optimized to minimize capacity checks by calculating total size upfront.
#[inline(always)]
fn write_start_tag(out: &mut Vec<u8>, e: &quick_xml::events::BytesStart<'_>) {
    let raw = e.as_ref();
    let needed = raw.len() + 2; // '<' + raw + '>'
    out.reserve(needed);
    // SAFETY: We just reserved enough space, so these writes won't reallocate.
    // Using extend_from_slice in a single batch is faster than push + extend + push.
    let old_len = out.len();
    // Write all bytes at once to avoid multiple capacity checks
    unsafe {
        let ptr = out.as_mut_ptr().add(old_len);
        *ptr = b'<';
        std::ptr::copy_nonoverlapping(raw.as_ptr(), ptr.add(1), raw.len());
        *ptr.add(1 + raw.len()) = b'>';
        out.set_len(old_len + needed);
    }
}

/// Write a start tag - alias for write_start_tag (they were identical).
#[inline(always)]
fn write_start_tag_dynamic(out: &mut Vec<u8>, e: &quick_xml::events::BytesStart<'_>) {
    write_start_tag(out, e);
}

/// Write an empty tag "<name attrs/>".
/// Optimized to minimize capacity checks by calculating total size upfront.
#[inline(always)]
fn write_empty_tag(out: &mut Vec<u8>, e: &quick_xml::events::BytesStart<'_>) {
    let raw = e.as_ref();
    let needed = raw.len() + 3; // '<' + raw + '/>'
    out.reserve(needed);
    let old_len = out.len();
    // Write all bytes at once to avoid multiple capacity checks
    unsafe {
        let ptr = out.as_mut_ptr().add(old_len);
        *ptr = b'<';
        std::ptr::copy_nonoverlapping(raw.as_ptr(), ptr.add(1), raw.len());
        *ptr.add(1 + raw.len()) = b'/';
        *ptr.add(2 + raw.len()) = b'>';
        out.set_len(old_len + needed);
    }
}

/// Write a closing tag "</name>" efficiently.
#[inline(always)]
fn write_end_tag(out: &mut Vec<u8>, name: &[u8]) {
    let needed = name.len() + 3; // '</' + name + '>'
    out.reserve(needed);
    let old_len = out.len();
    unsafe {
        let ptr = out.as_mut_ptr().add(old_len);
        *ptr = b'<';
        *ptr.add(1) = b'/';
        std::ptr::copy_nonoverlapping(name.as_ptr(), ptr.add(2), name.len());
        *ptr.add(2 + name.len()) = b'>';
        out.set_len(old_len + needed);
    }
}

/// Count the number of top-level paragraphs and tables in the XML.
/// This is a fast pre-scan to determine exact allocation size.
#[inline]
fn count_elements(xml_bytes: &[u8]) -> usize {
    let mut count = 0usize;
    let mut i = 0usize;
    let len = xml_bytes.len();

    while i < len {
        let Some(tag_start) = memchr::memchr(b'<', &xml_bytes[i..]) else {
            break;
        };
        let tag_start = i + tag_start;

        // Check for <w:p
        if tag_start + 5 < len && &xml_bytes[tag_start..tag_start + 4] == b"<w:p" {
            let next_char = xml_bytes[tag_start + 4];
            if (next_char == b'>' || next_char == b' ' || next_char == b'/')
                && let Some(end) = find_paragraph_end(&xml_bytes[tag_start..])
            {
                count += 1;
                i = tag_start + end;
                continue;
            }
        }
        // Check for <w:tbl
        else if tag_start + 7 < len && &xml_bytes[tag_start..tag_start + 6] == b"<w:tbl" {
            let next_char = xml_bytes[tag_start + 6];
            if (next_char == b'>' || next_char == b' ')
                && let Some(end) = find_table_end(&xml_bytes[tag_start..])
            {
                count += 1;
                i = tag_start + end;
                continue;
            }
        }

        i = tag_start + 1;
    }

    count
}

/// Find the end of a `<w:p>` element. Handles nested paragraphs (though rare).
/// Returns byte offset AFTER the closing `</w:p>`.
#[inline]
fn find_paragraph_end(xml: &[u8]) -> Option<usize> {
    // Check for self-closing tag first: <w:p.../>
    let first_gt = memchr::memchr(b'>', xml)?;
    if first_gt > 0 && xml[first_gt - 1] == b'/' {
        return Some(first_gt + 1);
    }

    let mut depth = 1i32;
    let mut pos = first_gt + 1;

    while pos < xml.len() && depth > 0 {
        let Some(next_lt) = memchr::memchr(b'<', &xml[pos..]) else {
            break;
        };
        pos += next_lt;

        // Check for </w:p>
        if pos + 6 <= xml.len() && &xml[pos..pos + 5] == b"</w:p" {
            let c = xml[pos + 5];
            if c == b'>' {
                depth -= 1;
                if depth == 0 {
                    return Some(pos + 6);
                }
            }
        }
        // Check for <w:p> or <w:p ...> (nested paragraph - rare but possible)
        else if pos + 5 <= xml.len() && &xml[pos..pos + 4] == b"<w:p" {
            let c = xml[pos + 4];
            if c == b'>' || c == b' ' {
                let gt = memchr::memchr(b'>', &xml[pos..])?;
                if gt > 0 && xml[pos + gt - 1] != b'/' {
                    depth += 1;
                }
            }
        }
        pos += 1;
    }
    None
}

/// Find the end of a `<w:tbl>` element. Handles nested tables.
/// Returns byte offset AFTER the closing `</w:tbl>`.
#[inline]
fn find_table_end(xml: &[u8]) -> Option<usize> {
    // Check for self-closing tag first: <w:tbl.../>
    let first_gt = memchr::memchr(b'>', xml)?;
    if first_gt > 0 && xml[first_gt - 1] == b'/' {
        return Some(first_gt + 1);
    }

    let mut depth = 1i32;
    let mut pos = first_gt + 1;

    while pos < xml.len() && depth > 0 {
        let Some(next_lt) = memchr::memchr(b'<', &xml[pos..]) else {
            break;
        };
        pos += next_lt;

        // Check for </w:tbl>
        if pos + 8 <= xml.len() && &xml[pos..pos + 7] == b"</w:tbl" {
            let c = xml[pos + 7];
            if c == b'>' {
                depth -= 1;
                if depth == 0 {
                    return Some(pos + 8);
                }
            }
        }
        // Check for <w:tbl> or <w:tbl ...> (nested table)
        else if pos + 7 <= xml.len() && &xml[pos..pos + 6] == b"<w:tbl" {
            let c = xml[pos + 6];
            if c == b'>' || c == b' ' {
                let gt = memchr::memchr(b'>', &xml[pos..])?;
                if gt > 0 && xml[pos + gt - 1] != b'/' {
                    depth += 1;
                }
            }
        }
        pos += 1;
    }
    None
}

#[cfg(test)]
mod tests {
    // Tests will be added as we have a way to construct test XmlParts
}
