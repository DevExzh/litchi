/// Numbering support for reading numbering definitions from Word documents.
///
/// This module provides types and methods for accessing numbering (lists) in Word documents.
/// Numbering defines how lists and outline numbering are formatted.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;

/// Numbering definitions in a Word document.
///
/// Contains abstract numbering definitions and numbering instances.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
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
    /// Numbering type (e.g., "hybridMultilevel", "arabicPeriod")
    num_type: Option<String>,
}

/// A numbering instance (concrete use of an abstract numbering).
#[derive(Debug, Clone)]
pub struct Num {
    /// Numbering ID
    id: u32,
    /// Reference to abstract numbering ID
    abstract_num_id: u32,
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
        let xml_bytes = part.blob();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        let mut abstract_nums = Vec::new();
        let mut nums = Vec::new();
        let mut in_abstract_num = false;
        let mut in_num = false;
        let mut current_abstract_id: Option<u32> = None;
        let mut current_abstract_type: Option<String> = None;
        let mut current_num_id: Option<u32> = None;
        let mut current_abstract_num_id: Option<u32> = None;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    match e.local_name().as_ref() {
                        b"abstractNum" => {
                            in_abstract_num = true;
                            current_abstract_id = None;
                            current_abstract_type = None;

                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"abstractNumId" {
                                    let id_str = String::from_utf8_lossy(&attr.value);
                                    current_abstract_id =
                                        atoi_simd::parse::<u32, false, false>(id_str.as_bytes())
                                            .ok();
                                }
                            }
                        },
                        b"numStyleLink" if in_abstract_num => {
                            // Link to a style
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    current_abstract_type =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        },
                        b"num" if !in_abstract_num => {
                            in_num = true;
                            current_num_id = None;
                            current_abstract_num_id = None;

                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"numId" {
                                    let id_str = String::from_utf8_lossy(&attr.value);
                                    current_num_id =
                                        atoi_simd::parse::<u32, false, false>(id_str.as_bytes())
                                            .ok();
                                }
                            }
                        },
                        b"abstractNumId" if in_num => {
                            // Reference to abstract numbering
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    let id_str = String::from_utf8_lossy(&attr.value);
                                    current_abstract_num_id =
                                        atoi_simd::parse::<u32, false, false>(id_str.as_bytes())
                                            .ok();
                                }
                            }
                        },
                        _ => {},
                    }
                },
                Ok(Event::End(e)) => match e.local_name().as_ref() {
                    b"abstractNum" => {
                        if let Some(id) = current_abstract_id {
                            abstract_nums.push(AbstractNum {
                                id,
                                num_type: current_abstract_type.clone(),
                            });
                        }
                        in_abstract_num = false;
                    },
                    b"num" => {
                        if let (Some(id), Some(abstract_id)) =
                            (current_num_id, current_abstract_num_id)
                        {
                            nums.push(Num {
                                id,
                                abstract_num_id: abstract_id,
                            });
                        }
                        in_num = false;
                    },
                    _ => {},
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(Self {
            abstract_nums,
            nums,
        })
    }
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
        });
        numbering.abstract_nums.push(AbstractNum {
            id: 2,
            num_type: Some("arabicPeriod".to_string()),
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
        };

        assert_eq!(abstract_num.id(), 5);
        assert_eq!(abstract_num.num_type(), Some("bullet"));
    }

    #[test]
    fn test_abstract_num_no_type() {
        let abstract_num = AbstractNum {
            id: 3,
            num_type: None,
        };

        assert_eq!(abstract_num.id(), 3);
        assert_eq!(abstract_num.num_type(), None);
    }

    #[test]
    fn test_abstract_num_clone() {
        let abstract_num = AbstractNum {
            id: 7,
            num_type: Some("roman".to_string()),
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
}
