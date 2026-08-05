//! Table row/column structure state used by the semantic builders.

use super::{Error, MAX_TABLE_STRUCTURE_DEPTH, Result, TableGroup, TableRange, TableStructure};

struct StructureContext {
    display: Option<bool>,
    children: Vec<TableStructure>,
    header_start: Option<usize>,
}

impl StructureContext {
    fn root() -> Self {
        Self {
            display: None,
            children: Vec::new(),
            header_start: None,
        }
    }
}

pub(super) struct StructureStack {
    contexts: Vec<StructureContext>,
}

impl StructureStack {
    fn new() -> Self {
        Self {
            contexts: vec![StructureContext::root()],
        }
    }

    fn begin_group(&mut self, display: bool) -> Result<()> {
        if self.contexts.len() > MAX_TABLE_STRUCTURE_DEPTH {
            return Err(Error::InvalidFormat(format!(
                "table structure exceeds the {MAX_TABLE_STRUCTURE_DEPTH} level nesting safety limit"
            )));
        }
        if self
            .contexts
            .last()
            .is_some_and(|context| context.header_start.is_some())
        {
            return Err(Error::InvalidFormat(
                "table groups cannot be nested inside a header container".to_string(),
            ));
        }
        self.contexts.push(StructureContext {
            display: Some(display),
            children: Vec::new(),
            header_start: None,
        });
        Ok(())
    }

    fn end_group(&mut self) -> Result<()> {
        if self.contexts.len() <= 1 {
            return Err(Error::InvalidFormat(
                "table group end has no matching start".to_string(),
            ));
        }
        let context = self.contexts.pop().expect("non-root context was checked");
        if context.header_start.is_some() {
            return Err(Error::InvalidFormat(
                "table header container is not closed before its group".to_string(),
            ));
        }
        if context.children.is_empty() {
            return Err(Error::InvalidFormat(
                "table groups must contain at least one row or column".to_string(),
            ));
        }
        self.contexts
            .last_mut()
            .expect("root context is retained")
            .children
            .push(TableStructure::Group(TableGroup {
                display: context.display.expect("group contexts have display state"),
                children: context.children,
            }));
        Ok(())
    }

    fn begin_header(&mut self, position: usize) -> Result<()> {
        let context = self.contexts.last_mut().expect("root context is retained");
        if context.header_start.replace(position).is_some() {
            return Err(Error::InvalidFormat(
                "table header containers cannot be nested".to_string(),
            ));
        }
        Ok(())
    }

    fn end_header(&mut self, position: usize) -> Result<()> {
        let context = self.contexts.last_mut().expect("root context is retained");
        let start = context.header_start.take().ok_or_else(|| {
            Error::InvalidFormat("table header end has no matching start".to_string())
        })?;
        if position <= start {
            return Err(Error::InvalidFormat(
                "table header containers must not be empty".to_string(),
            ));
        }
        Ok(())
    }

    fn add_range(&mut self, start: usize, end: usize) -> Result<()> {
        let range = TableRange::new(start, end)?;
        let context = self.contexts.last_mut().expect("root context is retained");
        let entry = if context.header_start.is_some() {
            TableStructure::Header(range)
        } else {
            TableStructure::Range(range)
        };
        if let Some(previous) = context.children.last_mut() {
            match (previous, &entry) {
                (TableStructure::Range(previous), TableStructure::Range(next))
                | (TableStructure::Header(previous), TableStructure::Header(next))
                    if previous.end == next.start =>
                {
                    previous.end = next.end;
                    return Ok(());
                },
                _ => {},
            }
        }
        context.children.push(entry);
        Ok(())
    }

    fn finish(mut self) -> Result<Vec<TableStructure>> {
        if self.contexts.len() != 1 {
            return Err(Error::InvalidFormat(
                "table group is not closed before the table ends".to_string(),
            ));
        }
        let root = self.contexts.pop().expect("one root context was checked");
        if root.header_start.is_some() {
            return Err(Error::InvalidFormat(
                "table header container is not closed before the table ends".to_string(),
            ));
        }
        Ok(root.children)
    }
}
