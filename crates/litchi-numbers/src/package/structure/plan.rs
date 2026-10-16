//! The removal plan the ported CalculationEngine phase ([`super::formula`])
//! reads and extends, in the shapes of the Pages body-table deletion owner.

use std::collections::HashMap;
use std::num::NonZeroU64;

use super::{Fault, FaultResult, allocation};
use crate::package::Package;

/// Where one source object lives.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(in crate::package) struct DeletionObjectLocation {
    pub(in crate::package) component_index: usize,
    pub(in crate::package) object_index: usize,
}

/// The source-wide object index.
#[derive(Debug, Default)]
pub(in crate::package) struct DeletionObjectIndex {
    by_identifier: HashMap<NonZeroU64, DeletionObjectLocation>,
}

impl DeletionObjectIndex {
    /// Index every object of `package`; an identifier stored twice is
    /// refused.
    pub(in crate::package) fn new(package: &Package) -> FaultResult<Self> {
        let catalog = package.state.components.catalog();
        let mut by_identifier = HashMap::new();
        for (component_index, component) in catalog.iter().enumerate() {
            by_identifier
                .try_reserve(component.archive().objects.len())
                .map_err(|_error| allocation(component.archive().objects.len()))?;
            for (object_index, object) in component.archive().objects.iter().enumerate() {
                let identifier = object
                    .archive_info
                    .identifier
                    .and_then(NonZeroU64::new)
                    .ok_or(Fault::Invalid)?;
                if by_identifier
                    .insert(
                        identifier,
                        DeletionObjectLocation {
                            component_index,
                            object_index,
                        },
                    )
                    .is_some()
                {
                    return Err(Fault::Unsupported);
                }
            }
        }
        Ok(Self { by_identifier })
    }

    pub(in crate::package) fn get(
        &self,
        identifier: NonZeroU64,
    ) -> Option<&DeletionObjectLocation> {
        self.by_identifier.get(&identifier)
    }
}

/// One payload replacement and the header references it prunes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(in crate::package) struct MessageEdit {
    pub(in crate::package) component_index: usize,
    pub(in crate::package) object_index: usize,
    pub(in crate::package) message_index: usize,
    pub(in crate::package) object_identifier: NonZeroU64,
    pub(in crate::package) message_type: u32,
    pub(in crate::package) data: Vec<u8>,
    pub(in crate::package) remove_object_references: Vec<NonZeroU64>,
    pub(in crate::package) remove_data_references: Vec<NonZeroU64>,
}

/// One object selected for removal.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(in crate::package) struct ObjectRemoval {
    pub(in crate::package) component_index: usize,
    pub(in crate::package) object_index: usize,
    pub(in crate::package) identifier: NonZeroU64,
}

/// The message edits and object removals one phase proved.
#[derive(Debug, Clone, Default)]
pub(in crate::package) struct RemovalPlan {
    pub(in crate::package) message_edits: Vec<MessageEdit>,
    pub(in crate::package) object_removals: Vec<ObjectRemoval>,
    pub(in crate::package) formula_contexts: Vec<NonZeroU64>,
}

/// The graph phase's output the formula phase reads.
#[derive(Debug)]
pub(in crate::package) struct GraphPlan {
    pub(in crate::package) index: DeletionObjectIndex,
    pub(in crate::package) removals: RemovalPlan,
}

/// The formula phase's output.
#[derive(Debug)]
pub(in crate::package) struct FormulaPlan {
    pub(in crate::package) removals: RemovalPlan,
}
