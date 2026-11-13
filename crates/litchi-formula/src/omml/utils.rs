use crate::omml::elements::{ElementContext, ElementType};
use std::borrow::Cow as StdCow;

/// High-performance string interning using bumpalo arena
///
/// This function interns strings in the arena to avoid allocations
/// and improve memory locality.
pub fn intern_string<'arena>(arena: &'arena bumpalo::Bump, s: &str) -> &'arena str {
    arena.alloc_str(s)
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

    /// Check if stack is empty
    #[inline(always)]
    pub fn is_empty(&self) -> bool {
        self.stack.is_empty()
    }
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
            super::MathNode::Root { base, .. } if base.is_empty() => {
                return Err(super::OmmlError::MissingRequiredElement(
                    "Root base is empty".to_string(),
                ));
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
            super::MathNode::Fenced { content, .. } if content.is_empty() => {
                return Err(super::OmmlError::ValidationError(
                    "Fenced content is empty".to_string(),
                ));
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
        ElementType::Math
            // Math element should be root or not have a parent
            if parent_type.is_some() => {
                return Err(super::OmmlError::InvalidStructure(
                    "Math element should be root".to_string(),
                ));
            },
        ElementType::Numerator | ElementType::Denominator
            if !matches!(parent_type, Some(ElementType::Fraction)) => {
                return Err(super::OmmlError::InvalidStructure(
                    "Numerator/denominator must be inside fraction".to_string(),
                ));
            },
        ElementType::Degree
            if !matches!(parent_type, Some(ElementType::Radical)) => {
                return Err(super::OmmlError::InvalidStructure(
                    "Degree must be inside radical".to_string(),
                ));
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
                | ElementType::Integrand
                | ElementType::PreScript,
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
                | ElementType::Integrand
                | ElementType::PreScript,
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
