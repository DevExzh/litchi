//! Allocation-conscious rendering arena for postfix formula expressions.
//!
//! Formula archive readers own decoding, references, and format-specific error
//! types. This module owns the format-independent expression arena used after
//! those adapters have converted native events into renderable pieces.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The arena keeps its public budget contract beside its storage operations."
)]

use std::ops::Range;
use std::sync::atomic::{AtomicU32, Ordering};

/// Budget adapter used by the shared formula renderer.
///
/// The arena is shared by format crates with different error enums. Callers
/// implement this trait to keep output-limit, allocation, and malformed-handle
/// errors in their own error model while the arena retains one bounded code
/// path.
pub trait FormulaRenderBudget {
    /// Error returned by a budget operation.
    type Error;

    /// Construct an output-limit error for an observed byte count.
    fn output_limit(&self, observed: usize) -> Self::Error;

    /// Construct an allocation error for a bounded request.
    fn allocation(&self, resource: &'static str, amount: usize) -> Self::Error;

    /// Construct an error for an invalid or foreign expression handle.
    fn invalid(&self, message: &'static str) -> Self::Error;

    /// Check a retained or rendered output size without charging it.
    fn check(&self, amount: usize) -> Result<(), Self::Error>;

    /// Check the expression arena and traversal structure before reserving it.
    ///
    /// `nodes` is the resulting retained node count. `parts` includes the
    /// resulting retained arena parts and any temporary pending or cumulative
    /// traversal references that the renderer is about to reserve or visit.
    /// Implementations should treat this as a ceiling check only; structural
    /// work is charged separately by the format-specific decoder.
    fn check_structure(&self, nodes: usize, parts: usize) -> Result<(), Self::Error>;

    /// Charge a rendered output size before publishing it.
    fn charge(&mut self, amount: usize) -> Result<(), Self::Error>;
}

/// Opaque expression handle owned by one [`FormulaRenderer`] arena.
///
/// Handles are copyable so event visitors can keep postfix expressions on a
/// small stack. The private packed representation carries an arena generation
/// and a bounded node index, allowing a renderer to reject handles copied from
/// a different arena without adding metadata to every stored node.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FormulaExpr(u64);

impl FormulaExpr {
    const INDEX_MASK: u64 = u32::MAX as u64;

    fn new(arena_id: u32, index: usize) -> Option<Self> {
        if arena_id == 0 {
            return None;
        }
        let index = u32::try_from(index).ok()?;
        Some(Self((u64::from(arena_id) << 32) | u64::from(index)))
    }

    const fn arena_id(self) -> u32 {
        (self.0 >> 32) as u32
    }

    const fn index(self) -> usize {
        (self.0 & Self::INDEX_MASK) as usize
    }
}

static NEXT_ARENA_ID: AtomicU32 = AtomicU32::new(1);

fn next_arena_id() -> u32 {
    allocate_arena_id(&NEXT_ARENA_ID)
}

fn allocate_arena_id(counter: &AtomicU32) -> u32 {
    let mut current = counter.load(Ordering::Relaxed);
    loop {
        if current == 0 {
            // Zero is a permanent exhausted marker. Returning it makes every
            // later arena operation fail through the caller's typed budget
            // error instead of ever reusing a generation.
            return 0;
        }
        let next = current.wrapping_add(1);
        match counter.compare_exchange_weak(current, next, Ordering::Relaxed, Ordering::Relaxed) {
            Ok(_) => return current,
            Err(observed) => current = observed,
        }
    }
}

#[derive(Debug)]
enum FormulaPart {
    Static(&'static str),
    Owned(String),
    Expr(FormulaExpr),
}

#[derive(Debug)]
struct FormulaNode {
    parts: Range<usize>,
    rendered_len: usize,
}

/// Bounded expression arena and iterative renderer.
///
/// The arena stores each expression's parts once, measures output before
/// retaining new owned text, and renders with an explicit stack. It therefore
/// avoids recursive traversal and keeps every fallible allocation behind the
/// caller's budget adapter.
#[derive(Debug)]
pub struct FormulaRenderer {
    nodes: Vec<FormulaNode>,
    parts: Vec<FormulaPart>,
    owned_bytes: usize,
    arena_id: u32,
}

impl Default for FormulaRenderer {
    fn default() -> Self {
        Self {
            nodes: Vec::new(),
            parts: Vec::new(),
            owned_bytes: 0,
            arena_id: next_arena_id(),
        }
    }
}

impl FormulaRenderer {
    /// Number of expression nodes retained by the arena.
    #[must_use]
    pub const fn node_count(&self) -> usize {
        self.nodes.len()
    }

    /// Number of expression parts retained by the arena.
    #[must_use]
    pub const fn part_count(&self) -> usize {
        self.parts.len()
    }

    /// Check whether additional owned text can remain in the arena.
    pub fn check_additional_owned<B: FormulaRenderBudget>(
        &self,
        additional: usize,
        budget: &B,
    ) -> Result<(), B::Error> {
        let retained = self
            .owned_bytes
            .checked_add(additional)
            .and_then(|bytes| bytes.checked_add(1))
            .ok_or_else(|| budget.output_limit(usize::MAX))?;
        budget.check(retained)
    }

    /// Add a static expression part.
    pub fn static_expr<B: FormulaRenderBudget>(
        &mut self,
        value: &'static str,
        budget: &B,
    ) -> Result<FormulaExpr, B::Error> {
        self.fixed([FormulaPart::Static(value)], budget)
    }

    /// Add an owned expression part.
    pub fn owned_expr<B: FormulaRenderBudget>(
        &mut self,
        value: String,
        budget: &B,
    ) -> Result<FormulaExpr, B::Error> {
        self.fixed([FormulaPart::Owned(value)], budget)
    }

    /// Build a binary expression, optionally wrapped in parentheses.
    pub fn binary<B: FormulaRenderBudget>(
        &mut self,
        left: FormulaExpr,
        operator: &'static str,
        right: FormulaExpr,
        wrapped: bool,
        budget: &B,
    ) -> Result<FormulaExpr, B::Error> {
        if wrapped {
            self.fixed(
                [
                    FormulaPart::Static("("),
                    FormulaPart::Expr(left),
                    FormulaPart::Static(operator),
                    FormulaPart::Expr(right),
                    FormulaPart::Static(")"),
                ],
                budget,
            )
        } else {
            self.fixed(
                [
                    FormulaPart::Expr(left),
                    FormulaPart::Static(operator),
                    FormulaPart::Expr(right),
                ],
                budget,
            )
        }
    }

    /// Build a unary expression from static prefix and suffix parts.
    pub fn unary<B: FormulaRenderBudget>(
        &mut self,
        prefix: &'static str,
        expression: FormulaExpr,
        suffix: &'static str,
        budget: &B,
    ) -> Result<FormulaExpr, B::Error> {
        self.fixed(
            [
                FormulaPart::Static(prefix),
                FormulaPart::Expr(expression),
                FormulaPart::Static(suffix),
            ],
            budget,
        )
    }

    /// Build a comma-separated function call.
    pub fn comma_joined<B: FormulaRenderBudget>(
        &mut self,
        function_prefix: Option<String>,
        arguments: Vec<FormulaExpr>,
        open: &'static str,
        close: &'static str,
        budget: &B,
    ) -> Result<FormulaExpr, B::Error> {
        let part_count = arguments
            .len()
            .checked_mul(2)
            .and_then(|count| count.checked_add(3))
            .ok_or_else(|| budget.output_limit(usize::MAX))?;
        self.check_structure_for_addition(part_count, budget)?;
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(part_count)
            .map_err(|_error| budget.allocation("Numbers formula render parts", part_count))?;
        if let Some(label) = function_prefix {
            parts.push(FormulaPart::Owned(label));
        }
        parts.push(FormulaPart::Static(open));
        for (index, argument) in arguments.into_iter().enumerate() {
            if index != 0 {
                parts.push(FormulaPart::Static(","));
            }
            parts.push(FormulaPart::Expr(argument));
        }
        parts.push(FormulaPart::Static(close));
        self.dynamic(parts, budget)
    }

    /// Build a row and column delimited array expression.
    pub fn array<B: FormulaRenderBudget>(
        &mut self,
        values: Vec<FormulaExpr>,
        columns: usize,
        budget: &B,
    ) -> Result<FormulaExpr, B::Error> {
        let part_count = values
            .len()
            .checked_mul(2)
            .and_then(|count| count.checked_add(2))
            .ok_or_else(|| budget.output_limit(usize::MAX))?;
        self.check_structure_for_addition(part_count, budget)?;
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(part_count)
            .map_err(|_error| budget.allocation("Numbers formula array parts", part_count))?;
        parts.push(FormulaPart::Static("{"));
        for (index, value) in values.into_iter().enumerate() {
            if index != 0 {
                parts.push(FormulaPart::Static(
                    if columns != 0 && index % columns == 0 {
                        ";"
                    } else {
                        ","
                    },
                ));
            }
            parts.push(FormulaPart::Expr(value));
        }
        parts.push(FormulaPart::Static("}"));
        self.dynamic(parts, budget)
    }

    fn check_structure_for_addition<B: FormulaRenderBudget>(
        &self,
        part_count: usize,
        budget: &B,
    ) -> Result<(), B::Error> {
        let nodes = self
            .nodes
            .len()
            .checked_add(1)
            .ok_or_else(|| budget.output_limit(usize::MAX))?;
        let parts = self
            .parts
            .len()
            .checked_add(part_count)
            .ok_or_else(|| budget.output_limit(usize::MAX))?;
        budget.check_structure(nodes, parts)
    }

    fn fixed<const N: usize, B: FormulaRenderBudget>(
        &mut self,
        parts: [FormulaPart; N],
        budget: &B,
    ) -> Result<FormulaExpr, B::Error> {
        let (rendered_len, owned_bytes) = self.measure(&parts, budget)?;
        self.reserve_node(N, budget)?;
        let start = self.parts.len();
        self.parts.extend(parts);
        self.push_node(start, rendered_len, owned_bytes, budget)
    }

    fn dynamic<B: FormulaRenderBudget>(
        &mut self,
        parts: Vec<FormulaPart>,
        budget: &B,
    ) -> Result<FormulaExpr, B::Error> {
        let (rendered_len, owned_bytes) = self.measure(&parts, budget)?;
        self.reserve_node(parts.len(), budget)?;
        let start = self.parts.len();
        self.parts.extend(parts);
        self.push_node(start, rendered_len, owned_bytes, budget)
    }

    fn measure<B: FormulaRenderBudget>(
        &self,
        parts: &[FormulaPart],
        budget: &B,
    ) -> Result<(usize, usize), B::Error> {
        let mut rendered_len = 0usize;
        let mut owned_bytes = 0usize;
        for part in parts {
            let part_len = match part {
                FormulaPart::Static(value) => value.len(),
                FormulaPart::Owned(value) => {
                    owned_bytes = owned_bytes
                        .checked_add(value.len())
                        .ok_or_else(|| budget.output_limit(usize::MAX))?;
                    value.len()
                },
                FormulaPart::Expr(expression) => {
                    self.node(
                        *expression,
                        budget,
                        "Numbers formula renderer contains an invalid expression",
                    )?
                    .rendered_len
                },
            };
            rendered_len = rendered_len
                .checked_add(part_len)
                .ok_or_else(|| budget.output_limit(usize::MAX))?;
        }
        let retained_owned = self
            .owned_bytes
            .checked_add(owned_bytes)
            .and_then(|bytes| bytes.checked_add(1))
            .ok_or_else(|| budget.output_limit(usize::MAX))?;
        budget.check(retained_owned)?;
        let output_len = rendered_len
            .checked_add(1)
            .ok_or_else(|| budget.output_limit(usize::MAX))?;
        budget.check(output_len)?;
        Ok((rendered_len, owned_bytes))
    }

    fn reserve_node<B: FormulaRenderBudget>(
        &mut self,
        part_count: usize,
        budget: &B,
    ) -> Result<(), B::Error> {
        if self.arena_id == 0 {
            return Err(budget.invalid("Numbers formula renderer arena generation exhausted"));
        }
        if self.nodes.len() > u32::MAX as usize {
            return Err(budget.allocation(
                "Numbers formula render nodes",
                self.nodes.len().saturating_add(1),
            ));
        }
        let nodes = self
            .nodes
            .len()
            .checked_add(1)
            .ok_or_else(|| budget.output_limit(usize::MAX))?;
        let parts = self
            .parts
            .len()
            .checked_add(part_count)
            .ok_or_else(|| budget.output_limit(usize::MAX))?;
        budget.check_structure(nodes, parts)?;
        self.nodes.try_reserve_exact(1).map_err(|_error| {
            budget.allocation(
                "Numbers formula render nodes",
                self.nodes.len().saturating_add(1),
            )
        })?;
        self.parts.try_reserve_exact(part_count).map_err(|_error| {
            budget.allocation(
                "Numbers formula render parts",
                self.parts.len().saturating_add(part_count),
            )
        })?;
        Ok(())
    }

    fn push_node<B: FormulaRenderBudget>(
        &mut self,
        start: usize,
        rendered_len: usize,
        owned_bytes: usize,
        budget: &B,
    ) -> Result<FormulaExpr, B::Error> {
        self.owned_bytes = self
            .owned_bytes
            .checked_add(owned_bytes)
            .ok_or_else(|| budget.allocation("Numbers formula owned text", usize::MAX))?;
        let end = self.parts.len();
        let expression = FormulaExpr::new(self.arena_id, self.nodes.len())
            .ok_or_else(|| budget.allocation("Numbers formula render nodes", usize::MAX))?;
        self.nodes.push(FormulaNode {
            parts: start..end,
            rendered_len,
        });
        Ok(expression)
    }

    /// Render one expression into a newly allocated string.
    pub fn render<B: FormulaRenderBudget>(
        &self,
        expression: FormulaExpr,
        budget: &mut B,
    ) -> Result<String, B::Error> {
        let node = self.node(
            expression,
            budget,
            "Numbers formula has no renderable expression",
        )?;
        let output_len = node
            .rendered_len
            .checked_add(1)
            .ok_or_else(|| budget.output_limit(usize::MAX))?;
        budget.charge(output_len)?;

        let mut output = String::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|_error| budget.allocation("Numbers rendered formula", output_len))?;
        output.push('=');

        let mut pending = Vec::new();
        let mut traversed_parts = 0usize;
        self.push_parts_reversed(
            &mut pending,
            node.parts.clone(),
            &mut traversed_parts,
            budget,
        )?;
        while let Some(part) = pending.pop() {
            match part {
                FormulaPart::Static(value) => output.push_str(value),
                FormulaPart::Owned(value) => output.push_str(value),
                FormulaPart::Expr(child) => {
                    let child_node = self.node(
                        *child,
                        budget,
                        "Numbers formula renderer contains an invalid child",
                    )?;
                    // An empty child contributes no bytes to the measured
                    // output. Skipping its parts also prevents a hostile
                    // zero-width DAG from expanding exponentially while
                    // producing the same rendered string.
                    if child_node.rendered_len != 0 {
                        self.push_parts_reversed(
                            &mut pending,
                            child_node.parts.clone(),
                            &mut traversed_parts,
                            budget,
                        )?;
                    }
                },
            }
        }
        debug_assert_eq!(output.len(), output_len);
        Ok(output)
    }

    fn node<'a, B: FormulaRenderBudget>(
        &'a self,
        expression: FormulaExpr,
        budget: &B,
        message: &'static str,
    ) -> Result<&'a FormulaNode, B::Error> {
        if expression.arena_id() != self.arena_id {
            return Err(budget.invalid(message));
        }
        self.nodes
            .get(expression.index())
            .ok_or_else(|| budget.invalid(message))
    }

    fn push_parts_reversed<'a, B: FormulaRenderBudget>(
        &'a self,
        pending: &mut Vec<&'a FormulaPart>,
        range: Range<usize>,
        traversed_parts: &mut usize,
        budget: &B,
    ) -> Result<(), B::Error> {
        let count = range.len();
        let cumulative_traversed = traversed_parts
            .checked_add(count)
            .ok_or_else(|| budget.output_limit(usize::MAX))?;
        let pending_after = pending
            .len()
            .checked_add(count)
            .ok_or_else(|| budget.output_limit(usize::MAX))?;
        let traversal_parts = cumulative_traversed.max(pending_after);
        // Count retained arena parts and temporary traversal references
        // together. The caller's part ceiling covers both persistent RPN
        // storage and this transient work; cumulative scheduling keeps a
        // shared DAG from hiding repeated CPU work behind a small stack.
        let structure_parts = self
            .parts
            .len()
            .checked_add(traversal_parts)
            .ok_or_else(|| budget.output_limit(usize::MAX))?;
        budget.check_structure(self.nodes.len(), structure_parts)?;
        pending.try_reserve_exact(count).map_err(|_error| {
            budget.allocation(
                "Numbers formula render stack",
                pending.len().saturating_add(count),
            )
        })?;
        let parts = self.parts.get(range).ok_or_else(|| {
            budget.invalid("Numbers formula renderer contains an invalid part range")
        })?;
        *traversed_parts = cumulative_traversed;
        pending.extend(parts.iter().rev());
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::allocate_arena_id;
    use std::sync::atomic::AtomicU32;

    #[test]
    fn arena_generation_never_reuses_after_exhaustion() {
        let counter = AtomicU32::new(u32::MAX);
        assert_eq!(allocate_arena_id(&counter), u32::MAX);
        assert_eq!(allocate_arena_id(&counter), 0);
        assert_eq!(allocate_arena_id(&counter), 0);
    }
}
