//! Immutable hyperlink edits of an XLSB [`Package`] (unified-ops record
//! 0322).
//!
//! A worksheet stores each hyperlink as one `BrtHLink` record; an external
//! target lives in a worksheet relationship the record names. [`Edit`]
//! lists, adds, replaces and removes those records one at a time: every
//! other record keeps its bytes, an added external link gets a new
//! relationship, and a relationship no remaining link names is removed with
//! its last link. [`Edit::commit`] returns a [`Commit`] with the new
//! package, the reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_hyperlinks`] replays a patch onto an equal package.
//!
//! ```
//! use litchi_xlsb::Package;
//! use litchi_xlsb::hyperlinks::Hyperlink;
//!
//! # fn main() -> Result<(), litchi_xlsb::package::PackageError> {
//! let package = Package::new()?;
//! let mut edit = package.edit_hyperlinks()?;
//! let link = Hyperlink::new_external(0, 0, 0, 0, "https://example.test/".to_owned());
//! assert_eq!(edit.add(0, link)?, 0);
//! let after = edit.commit()?.into_package();
//! let links = after.edit_hyperlinks()?.hyperlinks(0)?;
//! assert_eq!(links[0].target.as_deref(), Some("https://example.test/"));
//! # Ok(())
//! # }
//! ```

use litchi_opc::constants::relationship_type;
use litchi_opc::{PackURI, TargetMode};

use super::Hyperlink;
use crate::package::Package;
use crate::package::error::Result;
use crate::part_edit::{
    Rank, SheetLayout, Staged, family, invalid, record_with_limits, snapshot_options, splice,
};
use crate::raw::{Records, kind};

family!("hyperlink");
snapshot_options!("hyperlink");

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Hyperlinks of one worksheet ([`MAX_HYPERLINKS`]).
    Hyperlinks,
}

/// The most hyperlinks one worksheet holds through this edit.
pub const MAX_HYPERLINKS: usize = 65_530;

/// A staged hyperlink edit. Dropping it discards it.
/// Each operation validates a detached candidate; a failed operation leaves
/// the worksheet, relationships and all earlier staged edits unchanged.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    options: EditOptions,
}

impl Package {
    /// Start a hyperlink edit.
    ///
    /// # Errors
    ///
    /// Returns the workbook materialization errors of this package.
    pub fn edit_hyperlinks(&self) -> Result<Edit> {
        self.edit_hyperlinks_with_options(EditOptions::default())
    }

    /// Start a hyperlink edit with an explicit retained-payload ceiling.
    pub fn edit_hyperlinks_with_options(&self, options: EditOptions) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::of_package(self)?,
            options,
        })
    }

    /// Apply a hyperlink patch to a new package; this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_hyperlinks(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

/// Read the hyperlinks of a worksheet stream, resolving each external
/// target through `relationships`.
pub(crate) fn read_layout(
    worksheet: &[u8],
    layout: &SheetLayout,
    relationships: &litchi_opc::Relationships,
    raw_limits: crate::raw::Limits,
) -> Result<Vec<Hyperlink>> {
    let mut links = Vec::new();
    for unit in layout.units_of(kind::H_LINK) {
        let item = Records::new(&worksheet[unit.span.clone()], raw_limits)
            .next()
            .transpose()?
            .ok_or_else(|| invalid("an empty hyperlink record"))?;
        let mut link = Hyperlink::parse(item.payload(), raw_limits)?;
        if !link.r_id.is_empty() {
            let relationship = relationships
                .get(&link.r_id)
                .ok_or(Error::InvalidRelationship(
                    "a hyperlink names a missing relationship",
                ))?;
            if relationship.is_external() {
                link.target = Some(relationship.target_ref().to_owned());
            }
        }
        links.push(link);
    }
    Ok(links)
}

fn check(link: &Hyperlink) -> Result<()> {
    if link.row_first > link.row_last
        || link.row_last > crate::page_breaks::LAST_ROW
        || link.col_first > link.col_last
        || link.col_last > crate::page_breaks::LAST_COLUMN
    {
        return Err(invalid("the hyperlink range is outside the worksheet grid"));
    }
    if link.target.is_none() && link.location.is_none() {
        return Err(invalid("a hyperlink needs a target or a location"));
    }
    Ok(())
}

impl Edit {
    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    fn sheet(&self, sheet: usize) -> Result<(PackURI, Vec<Hyperlink>, SheetLayout)> {
        let worksheet = self.staged.worksheet(sheet)?;
        let part = self.staged.opc().get_part(&worksheet)?;
        let layout = SheetLayout::scan_with_limits(part.blob(), self.staged.raw_limits())?;
        let links = read_layout(part.blob(), &layout, part.rels(), self.staged.raw_limits())?;
        Ok((worksheet, links, layout))
    }

    /// The staged hyperlinks of the worksheet at complete-catalog position
    /// `sheet`, in stream order, with each external target resolved.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab, a malformed
    /// stream or a link naming a missing relationship.
    pub fn hyperlinks(&self, sheet: usize) -> Result<Vec<Hyperlink>> {
        Ok(self.sheet(sheet)?.1)
    }

    /// Encode `link` for the worksheet: an external target gets a new
    /// relationship.
    fn encode(&mut self, worksheet: &PackURI, link: &Hyperlink) -> Result<Vec<u8>> {
        check(link)?;
        let raw_limits = self.staged.raw_limits();
        let mut stored = link.clone();
        stored.r_id = String::new();
        if let Some(target) = &link.target {
            let package = self.staged.opc_mut()?;
            let relationships = package.get_part_mut(worksheet)?.rels_mut();
            let mut number = relationships.len() + 1;
            let id = loop {
                let id = format!("rId{number}");
                if relationships.get(&id).is_none() {
                    break id;
                }
                number += 1;
            };
            relationships.try_add_relationship(
                relationship_type::HYPERLINK.to_owned(),
                target.clone(),
                id.clone(),
                TargetMode::External,
            )?;
            stored.r_id = id;
        }
        record_with_limits(kind::H_LINK, &stored.serialize(raw_limits)?, raw_limits)
    }

    /// Drop the relationship of a removed or replaced link when no other
    /// link names it.
    fn release(
        &mut self,
        worksheet: &PackURI,
        old: &Hyperlink,
        others: &[Hyperlink],
    ) -> Result<()> {
        if old.r_id.is_empty() || others.iter().any(|link| link.r_id == old.r_id) {
            return Ok(());
        }
        let package = self.staged.opc_mut()?;
        let relationships = package.get_part_mut(worksheet)?.rels_mut();
        if relationships
            .get(&old.r_id)
            .is_some_and(|relationship| relationship.reltype() == relationship_type::HYPERLINK)
        {
            relationships.remove(&old.r_id);
        }
        Ok(())
    }

    /// Append a hyperlink after the worksheet's last one; returns its
    /// position.
    ///
    /// # Errors
    ///
    /// Returns an error for a range outside the grid, a link with neither a
    /// target nor a location, more than [`MAX_HYPERLINKS`] links, or a
    /// changed edit of a signed workbook.
    pub fn add(&mut self, sheet: usize, link: Hyperlink) -> Result<usize> {
        let mut candidate = self.clone();
        let index = candidate.add_inner(sheet, link)?;
        *self = candidate;
        Ok(index)
    }

    fn add_inner(&mut self, sheet: usize, link: Hyperlink) -> Result<usize> {
        let (worksheet, links, layout) = self.sheet(sheet)?;
        if links.len() >= MAX_HYPERLINKS {
            return Err(invalid("the worksheet holds the most hyperlinks it can"));
        }
        let at = layout
            .units_of(kind::H_LINK)
            .last()
            .map_or_else(|| layout.insertion(Rank::Hyperlink), |unit| unit.span.end);
        let encoded = self.encode(&worksheet, &link)?;
        let bytes = splice(self.staged.blob(&worksheet)?, at..at, &encoded);
        self.staged.replace_worksheet(&worksheet, bytes)?;
        self.staged.touch(&worksheet);
        Ok(links.len())
    }

    /// Replace the hyperlink at `index`.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::add`], and an error for a missing position.
    pub fn set(&mut self, sheet: usize, index: usize, link: Hyperlink) -> Result<bool> {
        let mut candidate = self.clone();
        let changed = candidate.set_inner(sheet, index, link)?;
        *self = candidate;
        Ok(changed)
    }

    fn set_inner(&mut self, sheet: usize, index: usize, link: Hyperlink) -> Result<bool> {
        let (worksheet, links, layout) = self.sheet(sheet)?;
        let old = links
            .get(index)
            .cloned()
            .ok_or_else(|| invalid("the hyperlink is missing"))?;
        let mut unchanged = old.clone();
        unchanged.r_id = link.r_id.clone();
        if unchanged == link {
            return Ok(false);
        }
        let span = layout
            .units_of(kind::H_LINK)
            .nth(index)
            .map(|unit| unit.span.clone())
            .ok_or_else(|| invalid("the hyperlink is missing"))?;
        let others: Vec<_> = links
            .iter()
            .enumerate()
            .filter(|(position, _)| *position != index)
            .map(|(_, link)| link.clone())
            .collect();
        self.release(&worksheet, &old, &others)?;
        let encoded = self.encode(&worksheet, &link)?;
        let bytes = splice(self.staged.blob(&worksheet)?, span, &encoded);
        self.staged.replace_worksheet(&worksheet, bytes)?;
        self.staged.touch(&worksheet);
        Ok(true)
    }

    /// Remove the hyperlink at `index`.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing position, a malformed stream or a
    /// changed edit of a signed workbook.
    pub fn remove(&mut self, sheet: usize, index: usize) -> Result<Hyperlink> {
        let mut candidate = self.clone();
        let removed = candidate.remove_inner(sheet, index)?;
        *self = candidate;
        Ok(removed)
    }

    fn remove_inner(&mut self, sheet: usize, index: usize) -> Result<Hyperlink> {
        let (worksheet, mut links, layout) = self.sheet(sheet)?;
        let span = layout
            .units_of(kind::H_LINK)
            .nth(index)
            .map(|unit| unit.span.clone())
            .ok_or_else(|| invalid("the hyperlink is missing"))?;
        let old = links.remove(index);
        self.release(&worksheet, &old, &links)?;
        let bytes = splice(self.staged.blob(&worksheet)?, span, &[]);
        self.staged.replace_worksheet(&worksheet, bytes)?;
        self.staged.touch(&worksheet);
        Ok(old)
    }

    /// Validate the staged package and publish it.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit_with_snapshot_limit(
            self.options.max_snapshot_bytes_value(),
        )?))
    }
}

#[cfg(test)]
mod tests {
    use crate::Package;
    use crate::hyperlinks::Hyperlink;

    #[test]
    fn add_set_remove_replay() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_hyperlinks().unwrap();
        let external = Hyperlink::new_external(0, 0, 0, 0, "https://example.test/".to_owned())
            .with_tooltip("tip".to_owned());
        let internal = Hyperlink::new_internal(1, 1, 1, 1, "Sheet1!C3".to_owned());
        assert_eq!(edit.add(0, external.clone()).unwrap(), 0);
        assert_eq!(edit.add(0, internal.clone()).unwrap(), 1);
        let (after, patch) = edit.commit().unwrap().into_parts();
        let links = after.edit_hyperlinks().unwrap().hyperlinks(0).unwrap();
        assert_eq!(links.len(), 2);
        assert_eq!(links[0].target.as_deref(), Some("https://example.test/"));
        assert_eq!(links[0].tooltip.as_deref(), Some("tip"));
        assert_eq!(links[1].location.as_deref(), Some("Sheet1!C3"));
        assert_eq!(
            package
                .apply_hyperlinks(&patch)
                .map(|commit| commit.into_package())
                .unwrap()
                .to_bytes()
                .unwrap(),
            after.to_bytes().unwrap()
        );
        // Grid reads see the same links.
        let workbook = after.workbook().unwrap();
        let grid = workbook.grid(0usize).unwrap().unwrap();
        assert_eq!(grid.worksheet().hyperlinks().len(), 2);

        let mut edit = after.edit_hyperlinks().unwrap();
        let retarget = Hyperlink::new_external(0, 0, 0, 0, "https://other.test/".to_owned());
        assert!(edit.set(0, 0, retarget).unwrap());
        let set = edit.commit().unwrap().into_package();
        let links = set.edit_hyperlinks().unwrap().hyperlinks(0).unwrap();
        assert_eq!(links[0].target.as_deref(), Some("https://other.test/"));

        let mut edit = set.edit_hyperlinks().unwrap();
        edit.remove(0, 1).unwrap();
        edit.remove(0, 0).unwrap();
        let removed = edit.commit().unwrap().into_package();
        assert_eq!(removed.to_bytes().unwrap(), package.to_bytes().unwrap());
    }

    #[test]
    fn fixture_links_keep_their_bytes() {
        let bytes = std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsb/hyperlink.xlsb"
        ))
        .unwrap();
        let package = Package::from_bytes(&bytes).unwrap();
        let edit = package.edit_hyperlinks().unwrap();
        let links = edit.hyperlinks(0).unwrap();
        assert_eq!(links.len(), 1);
        assert_eq!(links[0].target.as_deref(), Some("http://tika.apache.org/"));
        let mut edit = package.edit_hyperlinks().unwrap();
        edit.add(0, Hyperlink::new_internal(3, 3, 0, 0, "A1".to_owned()))
            .unwrap();
        let after = edit.commit().unwrap().into_package();
        let mut edit = after.edit_hyperlinks().unwrap();
        edit.remove(0, 1).unwrap();
        let back = edit.commit().unwrap().into_package();
        let sheet = |package: &Package| {
            package
                .opc_package()
                .get_part(&litchi_opc::PackURI::new("/xl/worksheets/sheet1.bin").unwrap())
                .unwrap()
                .blob()
                .to_vec()
        };
        assert_eq!(sheet(&back), sheet(&package));
    }
}
