use super::super::codec::*;
use super::super::validation::*;
use super::super::*;
use super::*;
/// A checked pane selector. Add-in IDs are the semantic primary key; numeric
/// positions are available for ordered document workflows.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Selector<'a> {
    Id(&'a str),
    Index(usize),
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Id(value)
    }
}

impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct Panes {
    pub(in crate::web) panes: Vec<Pane>,
}

impl Panes {
    #[must_use]
    pub const fn new() -> Self {
        Self { panes: Vec::new() }
    }

    #[must_use]
    pub fn len(&self) -> usize {
        self.panes.len()
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.panes.is_empty()
    }

    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Pane> {
        self.panes.iter()
    }

    /// Look up a pane by semantic add-in ID or checked numeric position.
    #[must_use]
    pub fn get<'a, 'key>(&'a self, selector: impl Into<Selector<'key>>) -> Option<&'a Pane> {
        match selector.into() {
            Selector::Id(id) => self.panes.iter().find(|pane| pane.add_in.id == id),
            Selector::Index(index) => self.panes.get(index),
        }
    }

    /// Edit one pane transactionally while preserving collection-wide invariants.
    ///
    /// Image payloads remain shared through `Arc`; a failed edit leaves the
    /// original pane untouched. Returns `false` when the selector is absent.
    pub fn edit<'key>(
        &mut self,
        selector: impl Into<Selector<'key>>,
        edit: impl FnOnce(&mut Pane) -> Result<()>,
    ) -> Result<bool> {
        let index = match selector.into() {
            Selector::Id(id) => self.panes.iter().position(|pane| pane.add_in.id == id),
            Selector::Index(index) => (index < self.panes.len()).then_some(index),
        };
        let Some(index) = index else {
            return Ok(false);
        };

        let mut candidate = self.panes[index].clone();
        edit(&mut candidate)?;
        if self
            .panes
            .iter()
            .enumerate()
            .any(|(other, pane)| other != index && pane.add_in.id == candidate.add_in.id)
        {
            return invalid(format!("duplicate add-in id '{}'", candidate.add_in.id));
        }
        if self.panes.iter().enumerate().any(|(other, pane)| {
            other != index && pane.relationship_id == candidate.relationship_id
        }) {
            return invalid(format!(
                "duplicate task-pane relationship ID '{}'",
                candidate.relationship_id
            ));
        }
        canonicalize_pane_snapshot_resources(&mut candidate, &self.panes, Some(index))?;
        validate_task_pane(&candidate)?;
        self.panes[index] = candidate;
        Ok(true)
    }

    /// Remove a pane by semantic add-in ID or checked numeric position.
    pub fn remove<'key>(&mut self, selector: impl Into<Selector<'key>>) -> Option<Pane> {
        let index = match selector.into() {
            Selector::Id(id) => self.panes.iter().position(|pane| pane.add_in.id == id)?,
            Selector::Index(index) if index < self.panes.len() => index,
            Selector::Index(_) => return None,
        };
        Some(self.panes.remove(index))
    }

    pub fn push(&mut self, mut pane: Pane) -> Result<&mut Self> {
        if self.panes.len() >= MAX_WEB_EXTENSION_ITEMS {
            return limit(
                "web extension panes",
                MAX_WEB_EXTENSION_ITEMS,
                self.panes.len().saturating_add(1),
            );
        }
        if self
            .panes
            .iter()
            .any(|value| value.add_in.id == pane.add_in.id)
        {
            return invalid(format!("duplicate add-in id '{}'", pane.add_in.id));
        }
        canonicalize_pane_snapshot_resources(&mut pane, &self.panes, None)?;
        if pane.relationship_id.is_empty()
            || self
                .panes
                .iter()
                .any(|value| value.relationship_id == pane.relationship_id)
        {
            pane.relationship_id = self.next_relationship_id()?;
        }
        validate_task_pane(&pane)?;
        self.panes.push(pane);
        Ok(self)
    }

    pub(in crate::web) fn next_relationship_id(&self) -> Result<String> {
        let attempts = self.panes.len().checked_add(1).ok_or(Error::Limit {
            resource: "web extension pane relationship IDs",
            max: usize::MAX,
            actual: usize::MAX,
        })?;
        for index in 1..=attempts {
            let candidate = format!("rIdAddIn{index}");
            if self
                .panes
                .iter()
                .all(|pane| pane.relationship_id != candidate)
            {
                return Ok(candidate);
            }
        }
        Err(Error::Relationship(
            "unable to allocate an add-in relationship ID".into(),
        ))
    }
}
