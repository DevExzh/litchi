// OMML Context Management
//
// This module provides utilities for managing parsing context and element stacks.

use crate::formula::ast::{MathNode, LargeOperator, AccentType, Fence, MatrixFence};
use crate::formula::omml::elements::*;
use crate::formula::omml::utils::extend_vec_efficient;

/// High-performance text buffer for collecting text content
#[derive(Debug)]
pub struct TextBuffer {
    buffer: String,
    has_content: bool,
}

impl TextBuffer {
    pub fn new() -> Self {
        Self {
            buffer: String::new(),
            has_content: false,
        }
    }

    pub fn push_str(&mut self, s: &str) {
        if !s.is_empty() {
            self.buffer.push_str(s);
            self.has_content = true;
        }
    }

    pub fn clear(&mut self) {
        self.buffer.clear();
        self.has_content = false;
    }

    pub fn is_empty(&self) -> bool {
        !self.has_content
    }

    pub fn as_str(&self) -> &str {
        &self.buffer
    }

    pub fn into_string(self) -> String {
        self.buffer
    }
}

impl Default for TextBuffer {
    fn default() -> Self {
        Self::new()
    }
}

/// Context for tracking element state during parsing
#[derive(Debug)]
pub struct ElementContext<'arena> {
    pub element_type: ElementType,
    pub children: Vec<MathNode<'arena>>,
    pub text: TextBuffer,
    pub properties: ElementProperties,

    // Core mathematical components
    pub base: Option<Vec<MathNode<'arena>>>,
    pub numerator: Option<Vec<MathNode<'arena>>>,
    pub denominator: Option<Vec<MathNode<'arena>>>,
    pub subscript: Option<Vec<MathNode<'arena>>>,
    pub superscript: Option<Vec<MathNode<'arena>>>,
    pub degree: Option<Vec<MathNode<'arena>>>,
    pub lower_limit: Option<Vec<MathNode<'arena>>>,
    pub upper_limit: Option<Vec<MathNode<'arena>>>,
    pub integrand: Option<Vec<MathNode<'arena>>>,

    // Matrix and array structures
    pub matrix_rows: Vec<Vec<Vec<MathNode<'arena>>>>,
    pub matrix_cells: Vec<MathNode<'arena>>,
    pub eq_array_rows: Vec<Vec<MathNode<'arena>>>,

    // Function and operator data
    pub function_name: Option<String>,
    pub operator: Option<LargeOperator>,
    pub accent_type: Option<AccentType>,

    // Fencing and delimiters
    pub fence_open: Option<Fence>,
    pub fence_close: Option<Fence>,
    pub matrix_fence: Option<MatrixFence>,

    // Scripts and limits
    pub pre_scripts: Vec<MathNode<'arena>>,
    pub post_scripts: Vec<MathNode<'arena>>,

    // Spacing and layout
    pub spacing_nodes: Vec<MathNode<'arena>>,

    // Group and character data
    pub character_data: Option<String>,
}

impl<'arena> ElementContext<'arena> {
    pub fn new(element_type: ElementType) -> Self {
        Self {
            element_type,
            children: Vec::new(),
            text: TextBuffer::new(),
            properties: ElementProperties::default(),
            base: None,
            numerator: None,
            denominator: None,
            subscript: None,
            superscript: None,
            degree: None,
            lower_limit: None,
            upper_limit: None,
            integrand: None,
            matrix_rows: Vec::new(),
            matrix_cells: Vec::new(),
            eq_array_rows: Vec::new(),
            function_name: None,
            operator: None,
            accent_type: None,
            fence_open: None,
            fence_close: None,
            matrix_fence: None,
            pre_scripts: Vec::new(),
            post_scripts: Vec::new(),
            spacing_nodes: Vec::new(),
            character_data: None,
        }
    }

    /// Clear the context for reuse
    pub fn clear(&mut self) {
        self.children.clear();
        self.text.clear();
        self.properties = ElementProperties::default();
        self.base = None;
        self.numerator = None;
        self.denominator = None;
        self.subscript = None;
        self.superscript = None;
        self.degree = None;
        self.lower_limit = None;
        self.upper_limit = None;
        self.integrand = None;
        self.matrix_rows.clear();
        self.matrix_cells.clear();
        self.eq_array_rows.clear();
        self.function_name = None;
        self.fence_open = None;
        self.fence_close = None;
        self.operator = None;
        self.accent_type = None;
        self.matrix_fence = None;
        self.pre_scripts.clear();
        self.post_scripts.clear();
        self.spacing_nodes.clear();
        self.character_data = None;
    }

    /// Check if the context has any content
    #[inline]
    pub fn has_content(&self) -> bool {
        !self.children.is_empty()
            || !self.text.is_empty()
            || self.base.is_some()
            || self.numerator.is_some()
            || self.denominator.is_some()
            || self.subscript.is_some()
            || self.superscript.is_some()
            || self.degree.is_some()
            || self.lower_limit.is_some()
            || self.upper_limit.is_some()
            || self.integrand.is_some()
            || !self.matrix_rows.is_empty()
            || !self.matrix_cells.is_empty()
            || !self.eq_array_rows.is_empty()
            || self.function_name.is_some()
            || self.character_data.is_some()
    }

    /// Get the total number of child nodes across all collections
    #[inline]
    pub fn total_child_count(&self) -> usize {
        self.children.len()
            + self.pre_scripts.len()
            + self.post_scripts.len()
            + self.spacing_nodes.len()
            + self.matrix_cells.len()
            + self.eq_array_rows.iter().map(|row| row.len()).sum::<usize>()
            + self.matrix_rows.iter().map(|row| row.iter().map(|cell| cell.len()).sum::<usize>()).sum::<usize>()
    }

    /// Reserve capacity for expected number of children
    pub fn reserve_children(&mut self, capacity: usize) {
        self.children.reserve(capacity);
    }

    /// Reserve capacity for matrix rows
    pub fn reserve_matrix_rows(&mut self, capacity: usize) {
        self.matrix_rows.reserve(capacity);
    }

    /// Reserve capacity for equation array rows
    pub fn reserve_eq_array_rows(&mut self, capacity: usize) {
        self.eq_array_rows.reserve(capacity);
    }

    /// Check if this is a structural element (contains other elements)
    #[inline]
    pub fn is_structural(&self) -> bool {
        matches!(
            self.element_type,
            ElementType::Math
                | ElementType::Fraction
                | ElementType::Radical
                | ElementType::Superscript
                | ElementType::Subscript
                | ElementType::SubSup
                | ElementType::Delimiter
                | ElementType::Nary
                | ElementType::Function
                | ElementType::Matrix
                | ElementType::Accent
                | ElementType::Bar
                | ElementType::Box
                | ElementType::Phantom
                | ElementType::GroupChar
                | ElementType::BorderBox
                | ElementType::EqArr
        )
    }

    /// Check if this is a leaf element (contains only text/symbols)
    #[inline]
    pub fn is_leaf(&self) -> bool {
        matches!(
            self.element_type,
            ElementType::Text | ElementType::Character | ElementType::Run
        )
    }
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

/// Fast element stacking
///
/// Custom stack implementation optimized for OMML parsing.
/// Pre-allocates capacity and provides fast access patterns.
pub struct ElementStack<'arena> {
    stack: Vec<ElementContext<'arena>>,
}

impl<'arena> ElementStack<'arena> {
    /// Create a new stack with pre-allocated capacity for performance
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
    pub fn len(&self) -> usize {
        self.stack.len()
    }

    /// Clear all elements from the stack
    pub fn clear(&mut self) {
        self.stack.clear();
    }

    /// Get the capacity of the underlying vector
    #[inline(always)]
    pub fn capacity(&self) -> usize {
        self.stack.capacity()
    }

    /// Reserve additional capacity
    pub fn reserve(&mut self, additional: usize) {
        self.stack.reserve(additional);
    }

    /// Shrink capacity to fit current length
    pub fn shrink_to_fit(&mut self) {
        self.stack.shrink_to_fit();
    }
}
