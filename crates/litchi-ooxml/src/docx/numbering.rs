/// Numbering support for reading numbering definitions from Word documents.
///
/// This module provides types and methods for accessing numbering (lists) in Word documents.
/// Numbering defines how lists and outline numbering are formatted.
use crate::docx::namespace::{is_wordprocessing_namespace, word_attribute_value};
use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

/// Numbering definitions in a Word document.
///
/// Contains abstract numbering definitions and numbering instances.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// if let Some(numbering) = doc.numbering()? {
///     println!("Found {} numbering definitions", numbering.num_count());
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct Numbering {
    /// Abstract numbering definitions (templates)
    abstract_nums: Vec<AbstractNum>,
    /// Numbering instances (concrete uses)
    nums: Vec<Num>,
}

/// An abstract numbering definition (template).
#[derive(Debug, Clone)]
pub struct AbstractNum {
    /// Abstract numbering ID
    id: u32,
    /// Multilevel type (`singleLevel`, `multilevel`, or `hybridMultilevel`).
    num_type: Option<String>,
    /// Paragraph style whose numbering definition this abstract definition inherits.
    num_style_link: Option<String>,
}

/// A numbering instance (concrete use of an abstract numbering).
#[derive(Debug, Clone)]
pub struct Num {
    /// Numbering ID
    id: u32,
    /// Reference to abstract numbering ID
    abstract_num_id: u32,
}

struct PendingAbstract {
    depth: usize,
    id: u32,
    num_type: Option<String>,
    num_style_link: Option<String>,
}

struct PendingNum {
    depth: usize,
    id: u32,
    abstract_num_id: Option<u32>,
}

impl Numbering {
    /// Create a new empty Numbering.
    pub fn new() -> Self {
        Self {
            abstract_nums: Vec::new(),
            nums: Vec::new(),
        }
    }

    /// Get all abstract numbering definitions.
    #[inline]
    pub fn abstract_nums(&self) -> &[AbstractNum] {
        &self.abstract_nums
    }

    /// Get all numbering instances.
    #[inline]
    pub fn nums(&self) -> &[Num] {
        &self.nums
    }

    /// Get the count of abstract numbering definitions.
    #[inline]
    pub fn abstract_num_count(&self) -> usize {
        self.abstract_nums.len()
    }

    /// Get the count of numbering instances.
    #[inline]
    pub fn num_count(&self) -> usize {
        self.nums.len()
    }

    /// Get an abstract numbering definition by ID.
    pub fn get_abstract_num(&self, id: u32) -> Option<&AbstractNum> {
        self.abstract_nums.iter().find(|a| a.id == id)
    }

    /// Get a numbering instance by ID.
    pub fn get_num(&self, id: u32) -> Option<&Num> {
        self.nums.iter().find(|n| n.id == id)
    }

    /// Extract numbering from a numbering.xml part.
    ///
    /// # Arguments
    ///
    /// * `part` - The numbering part
    ///
    /// # Returns
    ///
    /// A Numbering object
    pub(crate) fn extract_from_part(part: &dyn Part) -> Result<Self> {
        let mut reader = NsReader::from_reader(part.blob());
        let mut abstract_nums = Vec::new();
        let mut nums = Vec::new();
        let mut pending_abstract: Option<PendingAbstract> = None;
        let mut pending_num: Option<PendingNum> = None;
        let mut depth = 0usize;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);

            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("numbering XML nesting is too deep".to_string())
                    })?;
                    if is_wordprocessing_namespace(&namespace)
                        && element.local_name().as_ref() == b"abstractNum"
                    {
                        if pending_abstract.is_some() || pending_num.is_some() {
                            return Err(OoxmlError::InvalidFormat(
                                "nested numbering definitions are invalid".to_string(),
                            ));
                        }
                        pending_abstract = Some(PendingAbstract {
                            depth,
                            id: required_u32_attribute(
                                &element,
                                b"abstractNumId",
                                decoder,
                                &resolver,
                            )?,
                            num_type: None,
                            num_style_link: None,
                        });
                    } else if is_wordprocessing_namespace(&namespace)
                        && element.local_name().as_ref() == b"num"
                    {
                        if pending_abstract.is_some() || pending_num.is_some() {
                            return Err(OoxmlError::InvalidFormat(
                                "nested numbering definitions are invalid".to_string(),
                            ));
                        }
                        pending_num = Some(PendingNum {
                            depth,
                            id: required_u32_attribute(&element, b"numId", decoder, &resolver)?,
                            abstract_num_id: None,
                        });
                    } else {
                        parse_numbering_child(
                            &namespace,
                            &element,
                            decoder,
                            &resolver,
                            depth,
                            pending_abstract.as_mut(),
                            pending_num.as_mut(),
                        )?;
                    }
                },
                Event::Empty(element) => {
                    if is_wordprocessing_namespace(&namespace)
                        && element.local_name().as_ref() == b"abstractNum"
                    {
                        if pending_abstract.is_some() || pending_num.is_some() {
                            return Err(OoxmlError::InvalidFormat(
                                "nested numbering definitions are invalid".to_string(),
                            ));
                        }
                        let id =
                            required_u32_attribute(&element, b"abstractNumId", decoder, &resolver)?;
                        push_abstract_num(
                            &mut abstract_nums,
                            AbstractNum {
                                id,
                                num_type: None,
                                num_style_link: None,
                            },
                        )?;
                    } else if is_wordprocessing_namespace(&namespace)
                        && element.local_name().as_ref() == b"num"
                    {
                        return Err(OoxmlError::InvalidFormat(
                            "numbering instance is missing abstractNumId".to_string(),
                        ));
                    } else {
                        let child_depth = depth.checked_add(1).ok_or_else(|| {
                            OoxmlError::InvalidFormat(
                                "numbering XML nesting is too deep".to_string(),
                            )
                        })?;
                        parse_numbering_child(
                            &namespace,
                            &element,
                            decoder,
                            &resolver,
                            child_depth,
                            pending_abstract.as_mut(),
                            pending_num.as_mut(),
                        )?;
                    }
                },
                Event::End(element) => {
                    if is_wordprocessing_namespace(&namespace)
                        && element.local_name().as_ref() == b"abstractNum"
                        && pending_abstract
                            .as_ref()
                            .is_some_and(|pending| pending.depth == depth)
                    {
                        let pending = pending_abstract.take().ok_or_else(|| {
                            OoxmlError::InvalidFormat(
                                "missing abstract numbering definition".to_string(),
                            )
                        })?;
                        push_abstract_num(
                            &mut abstract_nums,
                            AbstractNum {
                                id: pending.id,
                                num_type: pending.num_type,
                                num_style_link: pending.num_style_link,
                            },
                        )?;
                    } else if is_wordprocessing_namespace(&namespace)
                        && element.local_name().as_ref() == b"num"
                        && pending_num
                            .as_ref()
                            .is_some_and(|pending| pending.depth == depth)
                    {
                        let pending = pending_num.take().ok_or_else(|| {
                            OoxmlError::InvalidFormat("missing numbering instance".to_string())
                        })?;
                        let abstract_num_id = pending.abstract_num_id.ok_or_else(|| {
                            OoxmlError::InvalidFormat(format!(
                                "numbering instance {} is missing abstractNumId",
                                pending.id
                            ))
                        })?;
                        if nums.iter().any(|number: &Num| number.id == pending.id) {
                            return Err(OoxmlError::InvalidFormat(format!(
                                "duplicate numbering instance ID {}",
                                pending.id
                            )));
                        }
                        nums.push(Num {
                            id: pending.id,
                            abstract_num_id,
                        });
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid numbering XML nesting".to_string())
                    })?;
                },
                Event::Eof if depth != 0 || pending_abstract.is_some() || pending_num.is_some() => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated numbering XML".to_string(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        for number in &nums {
            if !abstract_nums
                .iter()
                .any(|abstract_num| abstract_num.id == number.abstract_num_id)
            {
                return Err(OoxmlError::InvalidFormat(format!(
                    "numbering instance {} references missing abstractNum {}",
                    number.id, number.abstract_num_id
                )));
            }
        }

        Ok(Self {
            abstract_nums,
            nums,
        })
    }
}

fn parse_numbering_child(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    depth: usize,
    pending_abstract: Option<&mut PendingAbstract>,
    pending_num: Option<&mut PendingNum>,
) -> Result<()> {
    if !is_wordprocessing_namespace(namespace) {
        return Ok(());
    }
    if let Some(pending) = pending_abstract
        && depth == pending.depth + 1
    {
        match element.local_name().as_ref() {
            b"multiLevelType" => {
                if pending.num_type.is_some() {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "abstractNum {} has duplicate multiLevelType",
                        pending.id
                    )));
                }
                let value = required_string_attribute(element, b"val", decoder, resolver)?;
                if !matches!(
                    value.as_str(),
                    "singleLevel" | "multilevel" | "hybridMultilevel"
                ) {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "invalid multiLevelType '{value}'"
                    )));
                }
                pending.num_type = Some(value);
            },
            b"numStyleLink" => {
                if pending.num_style_link.is_some() {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "abstractNum {} has duplicate numStyleLink",
                        pending.id
                    )));
                }
                pending.num_style_link = Some(required_string_attribute(
                    element, b"val", decoder, resolver,
                )?);
            },
            _ => {},
        }
    }
    if let Some(pending) = pending_num
        && depth == pending.depth + 1
        && element.local_name().as_ref() == b"abstractNumId"
    {
        if pending.abstract_num_id.is_some() {
            return Err(OoxmlError::InvalidFormat(format!(
                "numbering instance {} has duplicate abstractNumId",
                pending.id
            )));
        }
        pending.abstract_num_id = Some(required_u32_attribute(element, b"val", decoder, resolver)?);
    }
    Ok(())
}

fn required_string_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<String> {
    word_attribute_value(element, name, decoder, resolver)?.ok_or_else(|| {
        OoxmlError::InvalidFormat(format!(
            "Word numbering element is missing required '{}' attribute",
            String::from_utf8_lossy(name)
        ))
    })
}

fn required_u32_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<u32> {
    let value = required_string_attribute(element, name, decoder, resolver)?;
    value
        .parse::<u32>()
        .map_err(|_| OoxmlError::InvalidFormat(format!("invalid Word numbering integer '{value}'")))
}

fn push_abstract_num(abstract_nums: &mut Vec<AbstractNum>, value: AbstractNum) -> Result<()> {
    if abstract_nums
        .iter()
        .any(|abstract_num| abstract_num.id == value.id)
    {
        return Err(OoxmlError::InvalidFormat(format!(
            "duplicate abstract numbering ID {}",
            value.id
        )));
    }
    abstract_nums.push(value);
    Ok(())
}

impl Default for Numbering {
    fn default() -> Self {
        Self::new()
    }
}

impl AbstractNum {
    /// Get the abstract numbering ID.
    #[inline]
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get the numbering type.
    #[inline]
    pub fn num_type(&self) -> Option<&str> {
        self.num_type.as_deref()
    }

    /// Get the linked numbering style, if present.
    #[inline]
    pub fn num_style_link(&self) -> Option<&str> {
        self.num_style_link.as_deref()
    }
}

impl Num {
    /// Get the numbering ID.
    #[inline]
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get the abstract numbering ID this references.
    #[inline]
    pub fn abstract_num_id(&self) -> u32 {
        self.abstract_num_id
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackURI;
    use litchi_opc::part::BlobPart;

    fn parse_numbering(xml: &[u8]) -> Result<Numbering> {
        let part = BlobPart::new(
            PackURI::new("/word/numbering.xml").unwrap(),
            "application/xml".to_string(),
            xml.to_vec(),
        );
        Numbering::extract_from_part(&part)
    }

    #[test]
    fn test_numbering_creation() {
        let numbering = Numbering::new();
        assert_eq!(numbering.abstract_num_count(), 0);
        assert_eq!(numbering.num_count(), 0);
    }

    #[test]
    fn test_numbering_default() {
        let numbering: Numbering = Default::default();
        assert_eq!(numbering.abstract_num_count(), 0);
        assert_eq!(numbering.num_count(), 0);
    }

    #[test]
    fn test_numbering_empty_accessors() {
        let numbering = Numbering::new();
        assert!(numbering.abstract_nums().is_empty());
        assert!(numbering.nums().is_empty());
        assert!(numbering.get_abstract_num(0).is_none());
        assert!(numbering.get_num(0).is_none());
    }

    #[test]
    fn test_numbering_with_abstract_nums() {
        let mut numbering = Numbering::new();
        numbering.abstract_nums.push(AbstractNum {
            id: 1,
            num_type: Some("hybridMultilevel".to_string()),
            num_style_link: None,
        });
        numbering.abstract_nums.push(AbstractNum {
            id: 2,
            num_type: Some("arabicPeriod".to_string()),
            num_style_link: None,
        });

        assert_eq!(numbering.abstract_num_count(), 2);
        assert_eq!(numbering.get_abstract_num(1).unwrap().id(), 1);
        assert_eq!(numbering.get_abstract_num(2).unwrap().id(), 2);
        assert!(numbering.get_abstract_num(3).is_none());
    }

    #[test]
    fn test_numbering_with_nums() {
        let mut numbering = Numbering::new();
        numbering.nums.push(Num {
            id: 10,
            abstract_num_id: 1,
        });
        numbering.nums.push(Num {
            id: 11,
            abstract_num_id: 2,
        });

        assert_eq!(numbering.num_count(), 2);
        assert_eq!(numbering.get_num(10).unwrap().abstract_num_id(), 1);
        assert_eq!(numbering.get_num(11).unwrap().abstract_num_id(), 2);
        assert!(numbering.get_num(99).is_none());
    }

    #[test]
    fn test_abstract_num_accessors() {
        let abstract_num = AbstractNum {
            id: 5,
            num_type: Some("bullet".to_string()),
            num_style_link: None,
        };

        assert_eq!(abstract_num.id(), 5);
        assert_eq!(abstract_num.num_type(), Some("bullet"));
    }

    #[test]
    fn test_abstract_num_no_type() {
        let abstract_num = AbstractNum {
            id: 3,
            num_type: None,
            num_style_link: None,
        };

        assert_eq!(abstract_num.id(), 3);
        assert_eq!(abstract_num.num_type(), None);
    }

    #[test]
    fn test_abstract_num_clone() {
        let abstract_num = AbstractNum {
            id: 7,
            num_type: Some("roman".to_string()),
            num_style_link: None,
        };
        let cloned = abstract_num.clone();

        assert_eq!(cloned.id(), abstract_num.id());
        assert_eq!(cloned.num_type(), abstract_num.num_type());
    }

    #[test]
    fn test_abstract_num_debug() {
        let abstract_num = AbstractNum {
            id: 1,
            num_type: Some("test".to_string()),
            num_style_link: None,
        };
        let debug_str = format!("{:?}", abstract_num);
        assert!(debug_str.contains("AbstractNum"));
        assert!(debug_str.contains("1"));
    }

    #[test]
    fn test_num_accessors() {
        let num = Num {
            id: 15,
            abstract_num_id: 3,
        };

        assert_eq!(num.id(), 15);
        assert_eq!(num.abstract_num_id(), 3);
    }

    #[test]
    fn test_num_clone() {
        let num = Num {
            id: 20,
            abstract_num_id: 5,
        };
        let cloned = num.clone();

        assert_eq!(cloned.id(), num.id());
        assert_eq!(cloned.abstract_num_id(), num.abstract_num_id());
    }

    #[test]
    fn test_num_debug() {
        let num = Num {
            id: 1,
            abstract_num_id: 2,
        };
        let debug_str = format!("{:?}", num);
        assert!(debug_str.contains("Num"));
        assert!(debug_str.contains("1"));
        assert!(debug_str.contains("2"));
    }

    #[test]
    fn test_numbering_clone() {
        let mut numbering = Numbering::new();
        numbering.abstract_nums.push(AbstractNum {
            id: 1,
            num_type: Some("type1".to_string()),
            num_style_link: None,
        });
        numbering.nums.push(Num {
            id: 10,
            abstract_num_id: 1,
        });

        let cloned = numbering.clone();
        assert_eq!(cloned.abstract_num_count(), 1);
        assert_eq!(cloned.num_count(), 1);
    }

    #[test]
    fn test_numbering_debug() {
        let numbering = Numbering::new();
        let debug_str = format!("{:?}", numbering);
        assert!(debug_str.contains("Numbering"));
    }

    #[test]
    fn parses_aliased_numbering_types_links_and_instances() {
        let numbering = parse_numbering(
            br#"<q:numbering xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:false="urn:not-wordprocessingml">
                <false:abstractNum false:abstractNumId="99"/>
                <q:abstractNum q:abstractNumId="1"><q:multiLevelType q:val="hybridMultilevel"/><q:numStyleLink q:val="List &amp; Style"/></q:abstractNum>
                <q:abstractNum q:abstractNumId="2"><q:multiLevelType q:val="singleLevel"/></q:abstractNum>
                <q:num q:numId="5"><q:abstractNumId q:val="1"/></q:num>
                <q:num q:numId="6"><q:abstractNumId q:val="2"/></q:num>
            </q:numbering>"#,
        )
        .unwrap();

        assert_eq!(numbering.abstract_num_count(), 2);
        assert_eq!(numbering.num_count(), 2);
        let abstract_num = numbering.get_abstract_num(1).unwrap();
        assert_eq!(abstract_num.num_type(), Some("hybridMultilevel"));
        assert_eq!(abstract_num.num_style_link(), Some("List & Style"));
        assert_eq!(numbering.get_num(5).unwrap().abstract_num_id(), 1);
    }

    #[test]
    fn parses_strict_and_empty_abstract_numbering() {
        let numbering = parse_numbering(
            br#"<s:numbering xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:abstractNum s:abstractNumId="3"/><s:num s:numId="7"><s:abstractNumId s:val="3"/></s:num></s:numbering>"#,
        )
        .unwrap();
        assert_eq!(numbering.get_abstract_num(3).unwrap().num_type(), None);
        assert_eq!(numbering.get_num(7).unwrap().abstract_num_id(), 3);
    }

    #[test]
    fn rejects_invalid_or_incomplete_numbering_definitions() {
        let wrapper = |content: &str| {
            format!(
                r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{content}</w:numbering>"#
            )
        };
        for invalid in [
            r#"<w:abstractNum/>"#,
            r#"<w:abstractNum w:abstractNumId="x"/>"#,
            r#"<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="invalid"/></w:abstractNum>"#,
            r#"<w:abstractNum w:abstractNumId="1"/><w:abstractNum w:abstractNumId="1"/>"#,
            r#"<w:num w:numId="1"/>"#,
            r#"<w:abstractNum w:abstractNumId="1"/><w:num w:numId="1"><w:abstractNumId w:val="2"/></w:num>"#,
        ] {
            assert!(parse_numbering(wrapper(invalid).as_bytes()).is_err());
        }
        assert!(parse_numbering(
            br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1">"#
        )
        .is_err());
    }
}
