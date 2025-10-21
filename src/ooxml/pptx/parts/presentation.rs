/// Presentation part - the main part in a .pptx package.
///
/// Corresponds to `/ppt/presentation.xml` in the package.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;

/// The main presentation part.
///
/// This part contains the presentation-level properties and references to slides,
/// slide masters, and other presentation resources.
///
/// # Example
///
/// ```rust,ignore
/// let pres_part = PresentationPart::from_part(opc_part)?;
/// let slide_count = pres_part.slide_count()?;
/// ```
pub struct PresentationPart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
}

impl<'a> PresentationPart<'a> {
    /// Create a PresentationPart from an OPC Part.
    ///
    /// # Arguments
    ///
    /// * `part` - The underlying OPC part
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let pres_part = PresentationPart::from_part(opc_part)?;
    /// ```
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        Ok(Self { part })
    }

    /// Get the XML bytes of the presentation.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.part.blob()
    }

    /// Get the number of slides in the presentation.
    ///
    /// This counts the `<p:sldId>` elements in the presentation.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let count = pres_part.slide_count()?;
    /// println!("Presentation has {} slides", count);
    /// ```
    pub fn slide_count(&self) -> Result<usize> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut count = 0;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"sldId" {
                        count += 1;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        Ok(count)
    }

    /// Get the slide width in EMUs (English Metric Units).
    ///
    /// Returns None if the slide size is not defined.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// if let Some(width) = pres_part.slide_width()? {
    ///     println!("Slide width: {} EMUs", width);
    /// }
    /// ```
    pub fn slide_width(&self) -> Result<Option<i64>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"sldSz" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"cx" {
                                let value = std::str::from_utf8(&attr.value)
                                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                                return value.parse::<i64>().map(Some).map_err(|e| {
                                    OoxmlError::Xml(format!("Invalid slide width: {}", e))
                                });
                            }
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        Ok(None)
    }

    /// Get the slide height in EMUs (English Metric Units).
    ///
    /// Returns None if the slide size is not defined.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// if let Some(height) = pres_part.slide_height()? {
    ///     println!("Slide height: {} EMUs", height);
    /// }
    /// ```
    pub fn slide_height(&self) -> Result<Option<i64>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"sldSz" {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"cy" {
                                let value = std::str::from_utf8(&attr.value)
                                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                                return value.parse::<i64>().map(Some).map_err(|e| {
                                    OoxmlError::Xml(format!("Invalid slide height: {}", e))
                                });
                            }
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        Ok(None)
    }

    /// Get the relationship IDs of all slides in presentation order.
    ///
    /// Returns a vector of relationship IDs that can be used to access
    /// the actual slide parts.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let slide_rids = pres_part.slide_rids()?;
    /// for rid in slide_rids {
    ///     // Use rid to get slide part
    /// }
    /// ```
    pub fn slide_rids(&self) -> Result<Vec<String>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut rids = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"sldId" {
                        for attr in e.attributes().flatten() {
                            // Look for r:id attribute (can be r:id or just id with relationships namespace)
                            let key = attr.key.as_ref();
                            // Check if this is the relationship ID attribute
                            if key == b"r:id"
                                || (key.starts_with(b"r:")
                                    && attr.key.local_name().as_ref() == b"id")
                                || attr.key.local_name().as_ref() == b"id"
                            {
                                let rid = std::str::from_utf8(&attr.value)
                                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                                // Only push if it looks like a relationship ID (starts with "rId")
                                if rid.starts_with("rId") {
                                    rids.push(rid.to_string());
                                    break;
                                }
                            }
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        Ok(rids)
    }

    /// Get the relationship IDs of all slide masters.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let master_rids = pres_part.slide_master_rids()?;
    /// ```
    pub fn slide_master_rids(&self) -> Result<Vec<String>> {
        let mut reader = Reader::from_reader(self.xml_bytes());
        reader.config_mut().trim_text(true);

        let mut rids = Vec::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() == b"sldMasterId" {
                        for attr in e.attributes().flatten() {
                            // Look for r:id attribute (can be r:id or just id with relationships namespace)
                            let key = attr.key.as_ref();
                            // Check if this is the relationship ID attribute
                            if key == b"r:id"
                                || (key.starts_with(b"r:")
                                    && attr.key.local_name().as_ref() == b"id")
                                || attr.key.local_name().as_ref() == b"id"
                            {
                                let rid = std::str::from_utf8(&attr.value)
                                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                                // Only push if it looks like a relationship ID (starts with "rId")
                                if rid.starts_with("rId") {
                                    rids.push(rid.to_string());
                                    break;
                                }
                            }
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
        }

        Ok(rids)
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}
