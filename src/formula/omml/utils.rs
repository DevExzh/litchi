use crate::formula::ast::MathNode;
use crate::formula::omml::elements::{ElementContext, ElementType};
use std::borrow::Cow as StdCow;

/// High-performance string interning using bumpalo arena
///
/// This function interns strings in the arena to avoid allocations
/// and improve memory locality.
pub fn intern_string<'arena>(arena: &'arena bumpalo::Bump, s: &str) -> &'arena str {
    arena.alloc_str(s)
}

/// Fast text content extraction and processing
///
/// Extracts text content from MathNodes, handling different node types efficiently.
#[allow(dead_code)] // Utility function reserved for text extraction features
pub fn extract_text_content(nodes: &[MathNode]) -> String {
    let mut result = String::new();
    for node in nodes {
        match node {
            MathNode::Text(text) => result.push_str(text.as_ref()),
            MathNode::Fenced { content, .. } => {
                result.push_str(&extract_text_content(content));
            },
            MathNode::Function { name, argument } => {
                result.push_str(name.as_ref());
                result.push('(');
                result.push_str(&extract_text_content(argument));
                result.push(')');
            },
            // Add more cases as needed
            _ => {}, // Skip non-text nodes
        }
    }
    result
}

/// Fast element type lookup using perfect hashing
///
/// Pre-computed hash table for element name to type mapping.
/// This provides O(1) lookup instead of string matching.
#[allow(dead_code)] // Alternative element type lookup, reserved for optimization
pub fn get_element_type_fast(name: &[u8]) -> ElementType {
    match name {
        b"m:oMath" | b"oMath" => ElementType::Math,
        b"m:r" | b"r" => ElementType::Run,
        b"m:t" | b"t" => ElementType::Text,
        b"m:f" | b"f" => ElementType::Fraction,
        b"m:num" | b"num" => ElementType::Numerator,
        b"m:den" | b"den" => ElementType::Denominator,
        b"m:rad" | b"rad" => ElementType::Radical,
        b"m:deg" | b"deg" => ElementType::Degree,
        b"m:e" | b"e" => ElementType::Base,
        b"m:sSup" | b"sSup" => ElementType::Superscript,
        b"m:sSub" | b"sSub" => ElementType::Subscript,
        b"m:sSubSup" | b"sSubSup" => ElementType::SubSup,
        b"m:sup" | b"sup" => ElementType::SuperscriptElement,
        b"m:sub" | b"sub" => ElementType::SubscriptElement,
        b"m:d" | b"d" => ElementType::Delimiter,
        b"m:nary" | b"nary" => ElementType::Nary,
        b"m:func" | b"func" => ElementType::Function,
        b"m:fName" | b"fName" => ElementType::FunctionName,
        b"m:m" | b"m" => ElementType::Matrix,
        b"m:mr" | b"mr" => ElementType::MatrixRow,
        b"m:mPr" | b"mPr" => ElementType::Properties,
        b"m:acc" | b"acc" => ElementType::Accent,
        b"m:accPr" | b"accPr" => ElementType::AccentProperties,
        b"m:bar" | b"bar" => ElementType::Bar,
        b"m:box" | b"box" => ElementType::Box,
        b"m:phant" | b"phant" => ElementType::Phantom,
        b"m:groupChr" | b"groupChr" => ElementType::GroupChar,
        b"m:borderBox" | b"borderBox" => ElementType::BorderBox,
        b"m:eqArr" | b"eqArr" => ElementType::EqArr,
        b"m:eqArrPr" | b"eqArrPr" => ElementType::EqArrPr,
        b"m:rPr" | b"rPr" => ElementType::Properties,
        b"m:fPr" | b"fPr" => ElementType::Properties,
        b"m:radPr" | b"radPr" => ElementType::Properties,
        b"m:sSupPr" | b"sSupPr" => ElementType::Properties,
        b"m:sSubPr" | b"sSubPr" => ElementType::Properties,
        b"m:dPr" | b"dPr" => ElementType::Properties,
        b"m:naryPr" | b"naryPr" => ElementType::Properties,
        b"m:funcPr" | b"funcPr" => ElementType::Properties,
        b"m:groupChrPr" | b"groupChrPr" => ElementType::Properties,
        b"m:chr" | b"chr" => ElementType::Text, // Character element
        b"m:sPre" | b"sPre" => ElementType::Run, // Pre-script
        b"m:sPost" | b"sPost" => ElementType::Run, // Post-script
        b"m:lim" | b"lim" => ElementType::Nary, // Limit
        b"m:limLow" | b"limLow" => ElementType::SubscriptElement, // Lower limit
        b"m:limUpp" | b"limUpp" => ElementType::SuperscriptElement, // Upper limit
        _ => ElementType::Unknown,
    }
}

/// Fast attribute lookup using SIMD-accelerated search
///
/// Uses memchr for fast byte searching in attribute data.
#[allow(dead_code)] // Alternative attribute lookup, reserved for optimization
pub fn find_attribute_fast<'a>(
    attrs: &'a [quick_xml::events::attributes::Attribute<'a>],
    key: &str,
) -> Option<&'a quick_xml::events::attributes::Attribute<'a>> {
    for attr in attrs {
        if let Ok(attr_key) = std::str::from_utf8(attr.key.as_ref())
            && (attr_key == key || attr_key == format!("m:{}", key))
        {
            return Some(attr);
        }
    }
    None
}

/// Batch processing of element contexts
///
/// Reuses element contexts to reduce allocations.
pub struct ContextPool<'arena> {
    pool: Vec<ElementContext<'arena>>,
    available: Vec<usize>,
}

impl<'arena> ContextPool<'arena> {
    pub fn new(capacity: usize) -> Self {
        Self {
            pool: Vec::with_capacity(capacity),
            available: Vec::new(),
        }
    }

    pub fn get(&mut self, element_type: ElementType) -> ElementContext<'arena> {
        if let Some(index) = self.available.pop() {
            let mut context = self.pool.swap_remove(index);
            context.element_type = element_type;
            context.clear();
            context
        } else {
            ElementContext::new(element_type)
        }
    }

    pub fn put(&mut self, mut context: ElementContext<'arena>) {
        if self.pool.len() < self.pool.capacity() {
            context.clear();
            self.pool.push(context);
        }
        // If pool is full, context is dropped
    }
}

/// Zero-copy text processing
///
/// Processes text content without unnecessary allocations.
pub fn process_text_zero_copy<'a>(text: &'a str) -> StdCow<'a, str> {
    // Remove leading/trailing whitespace without allocation if possible
    let trimmed = text.trim();
    if trimmed.len() == text.len() {
        StdCow::Borrowed(text)
    } else {
        StdCow::Owned(trimmed.to_string())
    }
}

/// Fast numeric parsing for OMML attributes
///
/// Uses fast parsing libraries for performance.
#[allow(dead_code)] // Utility function for numeric attribute parsing
pub fn parse_numeric_attr(attr: Option<&str>) -> Option<f32> {
    attr.and_then(|s| fast_float2::parse(s).ok())
}

/// Memory-efficient vector operations
///
/// Extends vectors without unnecessary reallocations.
pub fn extend_vec_efficient<T>(vec: &mut Vec<T>, items: impl IntoIterator<Item = T>) {
    vec.extend(items);
}

/// Fast element stacking
///
/// Custom stack implementation optimized for OMML parsing.
/// Pre-allocates capacity and provides fast access patterns.
pub struct ElementStack<'arena> {
    stack: Vec<ElementContext<'arena>>,
}

impl<'arena> ElementStack<'arena> {
    /// Create a new stack with pre-allocated capacity for performance
    #[allow(dead_code)]
    pub fn new() -> Self {
        Self {
            stack: Vec::with_capacity(64), // Typical OMML depth is much less than this
        }
    }

    /// Create a new stack with specified capacity
    pub fn with_capacity(capacity: usize) -> Self {
        Self {
            stack: Vec::with_capacity(capacity),
        }
    }

    /// Push a context onto the stack
    #[inline(always)]
    pub fn push(&mut self, context: ElementContext<'arena>) {
        self.stack.push(context);
    }

    /// Pop a context from the stack
    #[inline(always)]
    pub fn pop(&mut self) -> Option<ElementContext<'arena>> {
        self.stack.pop()
    }

    /// Get reference to the top context
    #[inline(always)]
    pub fn last(&self) -> Option<&ElementContext<'arena>> {
        self.stack.last()
    }

    /// Get mutable reference to the top context
    #[inline(always)]
    pub fn last_mut(&mut self) -> Option<&mut ElementContext<'arena>> {
        self.stack.last_mut()
    }

    /// Get reference to the context at the specified depth from the top
    /// (0 = top, 1 = parent of top, etc.)
    #[inline(always)]
    #[allow(dead_code)] // Reserved for advanced stack operations
    pub fn peek(&self, depth: usize) -> Option<&ElementContext<'arena>> {
        let len = self.stack.len();
        if depth < len {
            Some(&self.stack[len - 1 - depth])
        } else {
            None
        }
    }

    /// Get mutable reference to the context at the specified depth from the top
    #[inline(always)]
    #[allow(dead_code)] // Reserved for advanced stack operations
    pub fn peek_mut(&mut self, depth: usize) -> Option<&mut ElementContext<'arena>> {
        let len = self.stack.len();
        if depth < len {
            let idx = len - 1 - depth;
            Some(&mut self.stack[idx])
        } else {
            None
        }
    }

    /// Check if stack is empty
    #[inline(always)]
    pub fn is_empty(&self) -> bool {
        self.stack.is_empty()
    }

    /// Get current stack depth
    #[inline(always)]
    #[allow(dead_code)]
    pub fn len(&self) -> usize {
        self.stack.len()
    }

    /// Clear all elements from the stack
    #[allow(dead_code)]
    pub fn clear(&mut self) {
        self.stack.clear();
    }

    /// Get the capacity of the underlying vector
    #[inline(always)]
    #[allow(dead_code)]
    pub fn capacity(&self) -> usize {
        self.stack.capacity()
    }

    /// Reserve additional capacity
    #[allow(dead_code)] // Reserved for stack optimization
    pub fn reserve(&mut self, additional: usize) {
        self.stack.reserve(additional);
    }

    /// Shrink capacity to fit current length
    #[allow(dead_code)] // Reserved for memory optimization
    pub fn shrink_to_fit(&mut self) {
        self.stack.shrink_to_fit();
    }
}

/// Fast XML namespace handling
///
/// Strips XML namespaces efficiently.
#[allow(dead_code)] // Utility function for namespace handling
pub fn strip_namespace(name: &[u8]) -> &[u8] {
    if let Some(colon_pos) = memchr::memchr(b':', name) {
        &name[colon_pos + 1..]
    } else {
        name
    }
}

/// Error handling utilities
///
/// Fast error path handling.
#[allow(dead_code)] // Utility function for error handling
pub fn handle_parse_error<T>(result: Result<T, impl std::fmt::Display>) -> Result<T, String> {
    result.map_err(|e| e.to_string())
}

/// Validation utilities for OMML parsing
///
/// Validates OMML element and attribute names.
#[allow(dead_code)] // Utility function for validation
pub fn is_valid_omml_element_name(name: &str) -> bool {
    // Basic validation - element names should be alphanumeric with possible namespace prefix
    !name.is_empty()
        && name
            .chars()
            .all(|c| c.is_alphanumeric() || c == ':' || c == '_' || c == '-')
}

/// Validates OMML attribute values for basic sanity checks
#[allow(dead_code)] // Utility function for validation
pub fn validate_omml_attribute_value(value: &str) -> bool {
    // Basic validation - no null bytes, reasonable length
    !value.is_empty() && value.len() < 10000 && !value.contains('\0')
}

/// Memory-efficient string deduplication
///
/// Uses a simple interning mechanism for frequently used strings.
#[allow(dead_code)] // Reserved for string interning optimization
pub struct StringInterner {
    strings: std::collections::HashSet<String>,
}

#[allow(dead_code)] // Reserved for string interning optimization
impl StringInterner {
    pub fn new() -> Self {
        Self {
            strings: std::collections::HashSet::new(),
        }
    }

    pub fn intern(&mut self, s: &str) -> &str {
        if self.strings.contains(s) {
            // Return reference to existing string
            self.strings.get(s).unwrap().as_str()
        } else {
            // Insert new string and return reference
            self.strings.insert(s.to_string());
            self.strings.get(s).unwrap().as_str()
        }
    }
}

/// Fast attribute value extraction with SIMD
///
/// Optimized version using SIMD for common patterns.
#[allow(dead_code)] // Alternative SIMD-accelerated attribute lookup
pub fn extract_attribute_value_simd<'a>(
    attrs: &'a [quick_xml::events::attributes::Attribute<'a>],
    key: &str,
) -> Option<&'a [u8]> {
    for attr in attrs {
        if let Ok(attr_key) = std::str::from_utf8(attr.key.as_ref())
            && (attr_key == key || attr_key == format!("m:{}", key))
        {
            return Some(&attr.value);
        }
    }
    None
}

/// XML content normalization
///
/// Normalizes whitespace and entities in XML text content.
#[allow(dead_code)] // Utility function for XML text normalization
pub fn normalize_xml_text(text: &str) -> String {
    // Basic normalization - collapse whitespace, unescape common entities
    text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&#x20;", " ")
        .replace("&#160;", "\u{00A0}") // Non-breaking space
}

/// OMML document validation
///
/// Validates the structure and content of parsed OMML.
pub fn validate_omml_structure(nodes: &[super::MathNode]) -> Result<(), super::OmmlError> {
    // Empty OMML documents are not allowed
    if nodes.is_empty() {
        return Err(super::OmmlError::InvalidStructure(
            "Empty OMML document".to_string(),
        ));
    }

    // Check for required root math element
    let has_math_root = nodes
        .iter()
        .any(|node| matches!(node, super::MathNode::Row(_)));
    if !has_math_root && !nodes.is_empty() {
        // Allow documents that don't start with explicit math element
        // as long as they contain valid mathematical content
        validate_math_nodes(nodes)?;
    }

    Ok(())
}

/// Validate mathematical nodes for structural correctness
pub fn validate_math_nodes(nodes: &[super::MathNode]) -> Result<(), super::OmmlError> {
    for node in nodes {
        match node {
            super::MathNode::Frac {
                numerator,
                denominator,
                ..
            } => {
                if numerator.is_empty() {
                    return Err(super::OmmlError::MissingRequiredElement(
                        "Fraction numerator is empty".to_string(),
                    ));
                }
                if denominator.is_empty() {
                    return Err(super::OmmlError::MissingRequiredElement(
                        "Fraction denominator is empty".to_string(),
                    ));
                }
            },
            super::MathNode::Root { base, .. } => {
                if base.is_empty() {
                    return Err(super::OmmlError::MissingRequiredElement(
                        "Root base is empty".to_string(),
                    ));
                }
            },
            super::MathNode::Power { base, exponent } => {
                if base.is_empty() {
                    return Err(super::OmmlError::MissingRequiredElement(
                        "Power base is empty".to_string(),
                    ));
                }
                if exponent.is_empty() {
                    return Err(super::OmmlError::MissingRequiredElement(
                        "Power exponent is empty".to_string(),
                    ));
                }
            },
            super::MathNode::Sub { base, subscript } => {
                if base.is_empty() {
                    return Err(super::OmmlError::MissingRequiredElement(
                        "Subscript base is empty".to_string(),
                    ));
                }
                if subscript.is_empty() {
                    return Err(super::OmmlError::MissingRequiredElement(
                        "Subscript is empty".to_string(),
                    ));
                }
            },
            super::MathNode::Function { name, argument } => {
                if name.is_empty() {
                    return Err(super::OmmlError::MissingRequiredElement(
                        "Function name is empty".to_string(),
                    ));
                }
                if argument.is_empty() {
                    return Err(super::OmmlError::MissingRequiredElement(
                        "Function argument is empty".to_string(),
                    ));
                }
            },
            super::MathNode::Fenced { content, .. } => {
                if content.is_empty() {
                    return Err(super::OmmlError::ValidationError(
                        "Fenced content is empty".to_string(),
                    ));
                }
            },
            super::MathNode::Matrix { rows, .. } => {
                if rows.is_empty() {
                    return Err(super::OmmlError::ValidationError(
                        "Matrix has no rows".to_string(),
                    ));
                }
                for (i, row) in rows.iter().enumerate() {
                    if row.is_empty() {
                        return Err(super::OmmlError::ValidationError(format!(
                            "Matrix row {} is empty",
                            i
                        )));
                    }
                }
            },
            _ => {}, // Other nodes don't have specific validation requirements
        }
    }
    Ok(())
}

/// Validate OMML element nesting
///
/// Checks that elements are properly nested according to OMML specification.
pub fn validate_element_nesting(
    element_type: &ElementType,
    parent_type: Option<&ElementType>,
) -> Result<(), super::OmmlError> {
    match element_type {
        ElementType::Math => {
            // Math element should be root or not have a parent
            if parent_type.is_some() {
                return Err(super::OmmlError::InvalidStructure(
                    "Math element should be root".to_string(),
                ));
            }
        },
        ElementType::Numerator | ElementType::Denominator => {
            if !matches!(parent_type, Some(ElementType::Fraction)) {
                return Err(super::OmmlError::InvalidStructure(
                    "Numerator/denominator must be inside fraction".to_string(),
                ));
            }
        },
        ElementType::Degree => {
            if !matches!(parent_type, Some(ElementType::Radical)) {
                return Err(super::OmmlError::InvalidStructure(
                    "Degree must be inside radical".to_string(),
                ));
            }
        },
        ElementType::Base => {
            match parent_type {
                Some(
                    ElementType::Superscript
                    | ElementType::Subscript
                    | ElementType::SubSup
                    | ElementType::Radical
                    | ElementType::Accent
                    | ElementType::Bar
                    | ElementType::GroupChar,
                ) => {},
                _ => {
                    // Allow base elements in other contexts too - they might be generic containers
                },
            }
        },
        ElementType::SuperscriptElement => match parent_type {
            Some(
                ElementType::Superscript
                | ElementType::SubSup
                | ElementType::Nary
                | ElementType::Integrand,
            ) => {},
            _ => {
                return Err(super::OmmlError::InvalidStructure(
                    "Superscript element in invalid context".to_string(),
                ));
            },
        },
        ElementType::SubscriptElement => match parent_type {
            Some(
                ElementType::Subscript
                | ElementType::SubSup
                | ElementType::Nary
                | ElementType::Integrand,
            ) => {},
            _ => {
                return Err(super::OmmlError::InvalidStructure(
                    "Subscript element in invalid context".to_string(),
                ));
            },
        },
        _ => {}, // Other elements have more flexible nesting rules
    }
    Ok(())
}
