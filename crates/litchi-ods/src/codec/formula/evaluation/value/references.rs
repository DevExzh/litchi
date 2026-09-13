//! Reference construction and non-lazy reference operators for the value VM.
//!
//! This child module owns endpoint expansion and the `:`, `!`, and `~`
//! operators. It deliberately uses the parent VM's private runtime types so
//! references retain borrowed parsed names and the parent storage/cancellation
//! policy.

use super::{
    EvaluationFailure, EvaluationResult, Rect, Resolver, RuntimeArea, RuntimeAreaSet,
    RuntimeReference, RuntimeValue, ScalarError, SheetRef, ValueEvaluator, WorkingValue,
    ensure_capacity, is_reference_error,
};
use crate::codec::formula::reference::{
    Address, Endpoint, EndpointValue, Reference, SheetSelector,
};
use litchi_core::Resource;

fn record_plane_end(
    record: &RuntimeReference<'_>,
    start: usize,
    total: usize,
) -> EvaluationResult<usize> {
    let mut planes = 0usize;
    for public in &record.areas {
        planes =
            planes
                .checked_add(public.extent()[0])
                .ok_or(EvaluationFailure::InvalidExpression(
                    "reference record plane count overflows",
                ))?;
    }
    let end = start
        .checked_add(planes)
        .ok_or(EvaluationFailure::InvalidExpression(
            "reference record area offset overflows",
        ))?;
    if end > total {
        return Err(EvaluationFailure::InvalidExpression(
            "reference record areas are not contiguous",
        ));
    }
    Ok(end)
}

impl<'expr, 'scalar, 'exec, 'position, R> ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>
where
    R: Resolver + ?Sized,
{
    pub(super) fn reference_value(
        &mut self,
        reference: &'expr Reference,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        match reference {
            Reference::Source { .. } => Err(EvaluationFailure::Unsupported(
                super::super::UnsupportedKind::Reference,
            )),
            Reference::Error => Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::Reference,
            ))),
            Reference::Local(_) => match self.reference_areas(reference)? {
                Some(areas) => Ok(RuntimeValue::Areas(areas)),
                None => Ok(RuntimeValue::Scalar(WorkingValue::Error(
                    ScalarError::Reference,
                ))),
            },
        }
    }

    /// Project a demanded single-cell reference without retaining the
    /// transient area/record vectors used by first-class reference values.
    ///
    /// This path is selected only by scalar-demand value frames.  Every
    /// other reference, including a one-cell reference used by a reference
    /// operator or a generic function argument, remains on
    /// [`reference_value`] so its area and lexical metadata stay available.
    pub(super) fn reference_scalar_value(
        &mut self,
        reference: &'expr Reference,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        let endpoint = match reference {
            Reference::Local(Address::Cell(endpoint))
                if matches!(&endpoint.value, EndpointValue::Cell(_)) =>
            {
                endpoint
            },
            _ => return self.reference_value(reference),
        };
        let Some(area) = self.endpoint_area(endpoint, None)? else {
            return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::Reference,
            )));
        };
        self.check_reference_cells(area.rect.count()?)?;
        // The normal path enforces this through the record, flat-area, and
        // public-area reservations.  Keep the logical admission check when
        // those transient vectors are deliberately omitted.
        self.check_reference_areas(1)?;

        // Defer the current-sheet probe and provider read until the scalar
        // consumer projects this token.  The old area path evaluates both
        // infix operands before projecting either one; deferral preserves
        // that resolver/read ordering while still avoiding all three fresh
        // metadata vectors.
        Ok(RuntimeValue::ScalarCell(area))
    }

    pub(super) fn reference_areas(
        &mut self,
        reference: &'expr Reference,
    ) -> EvaluationResult<Option<RuntimeAreaSet<'expr>>> {
        let address = match reference {
            Reference::Local(address) => address,
            Reference::Source { .. } | Reference::Error => return Ok(None),
        };
        match address {
            Address::Cell(endpoint) => {
                let Some(area) = self.endpoint_area(endpoint, None)? else {
                    return Ok(None);
                };
                self.check_reference_cells(area.rect.count()?)?;
                let mut areas = RuntimeAreaSet::direct(reference, self)?;
                areas.push(area, self)?;
                Ok(Some(areas))
            },
            Address::Cells(first, second)
            | Address::Columns(first, second)
            | Address::Rows(first, second) => {
                let first_sheet = self.endpoint_sheet(&first.sheet)?;
                let second_sheet =
                    self.endpoint_sheet_with_inherited(&second.sheet, first_sheet)?;
                let Some(first_index) = self.resolve_sheet(first_sheet)? else {
                    return Ok(None);
                };
                let Some(second_index) = self.resolve_sheet(second_sheet)? else {
                    return Ok(None);
                };
                let first_rect = self.endpoint_rect(first, first_sheet)?;
                let second_rect = self.endpoint_rect(second, second_sheet)?;
                let rect = first_rect.bounding(second_rect)?;
                let (start, end) = if first_index <= second_index {
                    (first_index, second_index)
                } else {
                    (second_index, first_index)
                };
                let sheet_count = self.resolver.sheet_count(self.execution)?;
                if end >= sheet_count {
                    return Ok(None);
                }
                let per_sheet = rect.count()?;
                let sheet_span = end
                    .checked_sub(start)
                    .and_then(|span| span.checked_add(1))
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "reference sheet span overflows",
                    ))?;
                let total = per_sheet
                    .checked_mul(sheet_span)
                    .ok_or_else(|| self.reference_limit_error(usize::MAX))?;
                self.check_reference_cells(total)?;
                self.check_reference_areas(sheet_span)?;

                // Validate every physical sheet before retaining any area. A
                // provider may expose sheets with different finite extents,
                // and an explicit endpoint must not defer that refusal to a
                // later read.
                for index in start..=end {
                    self.scalar.charge_work(1)?;
                    let sheet = if index == first_index {
                        first_sheet
                    } else if index == second_index {
                        second_sheet
                    } else {
                        let name = self.resolver.sheet_name_at(index, self.execution)?.ok_or(
                            EvaluationFailure::InvalidExpression("sheet order index has no name"),
                        )?;
                        SheetRef::Named(name)
                    };
                    if !self.rect_within_extent(sheet, rect)? {
                        return Ok(None);
                    }
                }

                let mut areas = RuntimeAreaSet::direct(reference, self)?;
                for index in start..=end {
                    self.scalar.charge_work(1)?;
                    let sheet = if index == first_index {
                        first_sheet
                    } else if index == second_index {
                        second_sheet
                    } else {
                        let name = self.resolver.sheet_name_at(index, self.execution)?.ok_or(
                            EvaluationFailure::InvalidExpression("sheet order index has no name"),
                        )?;
                        SheetRef::Named(name)
                    };
                    areas.push(
                        RuntimeArea {
                            sheet,
                            sheet_index: index,
                            rect,
                        },
                        self,
                    )?;
                }
                Ok(Some(areas))
            },
        }
    }

    pub(super) fn endpoint_sheet(
        &self,
        selector: &'expr SheetSelector,
    ) -> EvaluationResult<SheetRef<'expr>> {
        match selector {
            SheetSelector::Current => Ok(SheetRef::Current),
            SheetSelector::Explicit(locator) if locator.subtables.is_empty() => {
                Ok(SheetRef::Named(locator.sheet.name.as_str()))
            },
            SheetSelector::Explicit(_) => Err(EvaluationFailure::Unsupported(
                super::super::UnsupportedKind::Reference,
            )),
            SheetSelector::Inherited => Err(EvaluationFailure::InvalidExpression(
                "inherited endpoint has no first sheet",
            )),
        }
    }

    pub(super) fn endpoint_sheet_with_inherited(
        &self,
        selector: &'expr SheetSelector,
        first: SheetRef<'expr>,
    ) -> EvaluationResult<SheetRef<'expr>> {
        if matches!(selector, SheetSelector::Inherited) {
            Ok(first)
        } else {
            self.endpoint_sheet(selector)
        }
    }

    pub(super) fn endpoint_area(
        &self,
        endpoint: &'expr Endpoint,
        inherited: Option<SheetRef<'expr>>,
    ) -> EvaluationResult<Option<RuntimeArea<'expr>>> {
        let sheet = match (&endpoint.sheet, inherited) {
            (SheetSelector::Inherited, Some(sheet)) => sheet,
            (SheetSelector::Inherited, None) => {
                return Err(EvaluationFailure::InvalidExpression(
                    "inherited endpoint has no first sheet",
                ));
            },
            (selector, _) => self.endpoint_sheet(selector)?,
        };
        let Some(sheet_index) = self.resolve_sheet(sheet)? else {
            return Ok(None);
        };
        let rect = self.endpoint_rect(endpoint, sheet)?;
        if !self.rect_within_extent(sheet, rect)? {
            return Ok(None);
        }
        Ok(Some(RuntimeArea {
            sheet,
            sheet_index,
            rect,
        }))
    }

    fn rect_within_extent(&self, sheet: SheetRef<'expr>, rect: Rect) -> EvaluationResult<bool> {
        let Some(extent) = self
            .resolver
            .sheet_extent(self.sheet_name(sheet), self.execution)?
        else {
            return Ok(false);
        };
        Ok(rect.row_end <= extent.rows() && rect.column_end <= extent.columns())
    }

    pub(super) fn endpoint_rect(
        &self,
        endpoint: &Endpoint,
        sheet: SheetRef<'expr>,
    ) -> EvaluationResult<Rect> {
        match &endpoint.value {
            EndpointValue::Cell(cell) => {
                let column = super::column_number(&cell.column.label)?;
                let row = usize::try_from(cell.row.number.saturating_sub(1)).map_err(|_| {
                    EvaluationFailure::InvalidExpression("reference row exceeds usize")
                })?;
                Rect::cell(row, column)
            },
            EndpointValue::Column(column) => {
                let start = super::column_number(&column.label)?;
                let extent = self
                    .resolver
                    .sheet_extent(self.sheet_name(sheet), self.execution)?
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "whole-column reference names a missing sheet",
                    ))?;
                let end = start
                    .checked_add(1)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "reference column range overflow",
                    ))?;
                if end > extent.columns() {
                    return Err(EvaluationFailure::InvalidExpression(
                        "whole-column reference exceeds sheet extent",
                    ));
                }
                Rect::new(0, extent.rows(), start, end)
            },
            EndpointValue::Row(row) => {
                let start = usize::try_from(row.number.saturating_sub(1)).map_err(|_| {
                    EvaluationFailure::InvalidExpression("reference row exceeds usize")
                })?;
                let extent = self
                    .resolver
                    .sheet_extent(self.sheet_name(sheet), self.execution)?
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "whole-row reference names a missing sheet",
                    ))?;
                let end = start
                    .checked_add(1)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "reference row range overflow",
                    ))?;
                if end > extent.rows() {
                    return Err(EvaluationFailure::InvalidExpression(
                        "whole-row reference exceeds sheet extent",
                    ));
                }
                Rect::new(start, end, 0, extent.columns())
            },
        }
    }

    pub(super) fn combine_range(
        &mut self,
        left: RuntimeValue<'expr>,
        right: RuntimeValue<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        if is_reference_error(&left) || is_reference_error(&right) {
            return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::Reference,
            )));
        }
        let left = self.coerce_reference_areas(left)?;
        let right = self.coerce_reference_areas(right)?;
        let Some(first_left) = left.areas.first().copied() else {
            return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::Reference,
            )));
        };
        let Some(first_right) = right.areas.first().copied() else {
            return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::Reference,
            )));
        };

        let mut min_sheet = first_left.sheet_index.min(first_right.sheet_index);
        let mut max_sheet = first_left.sheet_index.max(first_right.sheet_index);
        let mut rect = first_left.rect;
        for area in left
            .areas
            .iter()
            .copied()
            .chain(right.areas.iter().copied())
        {
            self.scalar.charge_work(1)?;
            min_sheet = min_sheet.min(area.sheet_index);
            max_sheet = max_sheet.max(area.sheet_index);
            rect = rect.bounding(area.rect)?;
        }
        let sheet_count = self.resolver.sheet_count(self.execution)?;
        if max_sheet >= sheet_count {
            return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::Reference,
            )));
        }
        let sheet_span = max_sheet
            .checked_sub(min_sheet)
            .and_then(|span| span.checked_add(1))
            .ok_or(EvaluationFailure::InvalidExpression(
                "range sheet span overflows",
            ))?;
        let total = rect
            .count()?
            .checked_mul(sheet_span)
            .ok_or_else(|| self.reference_limit_error(usize::MAX))?;
        self.check_reference_cells(total)?;
        self.check_reference_areas(sheet_span)?;

        // Resolve and validate the complete cuboid before allocating its
        // retained record or areas. This keeps an out-of-extent intermediate
        // sheet from leaving a partially-built reference behind.
        for index in min_sheet..=max_sheet {
            self.scalar.charge_work(1)?;
            let name = self.resolver.sheet_name_at(index, self.execution)?.ok_or(
                EvaluationFailure::InvalidExpression("range sheet order has no name"),
            )?;
            let sheet = SheetRef::Named(name);
            if !self.rect_within_extent(sheet, rect)? {
                return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                    ScalarError::Reference,
                )));
            }
        }

        let mut result = RuntimeAreaSet::derived(self)?;
        result.is_list = false;
        for index in min_sheet..=max_sheet {
            self.scalar.charge_work(1)?;
            let name = self.resolver.sheet_name_at(index, self.execution)?.ok_or(
                EvaluationFailure::InvalidExpression("range sheet order has no name"),
            )?;
            let sheet = SheetRef::Named(name);
            result.push(
                RuntimeArea {
                    sheet,
                    sheet_index: index,
                    rect,
                },
                self,
            )?;
        }
        Ok(RuntimeValue::Areas(result))
    }

    pub(super) fn intersect_ranges(
        &mut self,
        left: RuntimeValue<'expr>,
        right: RuntimeValue<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        if is_reference_error(&left) || is_reference_error(&right) {
            return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::Reference,
            )));
        }
        let left = self.coerce_reference_areas(left)?;
        let right = self.coerce_reference_areas(right)?;
        if (left.records.is_empty() && !left.areas.is_empty())
            || (right.records.is_empty() && !right.areas.is_empty())
        {
            return Err(EvaluationFailure::InvalidExpression(
                "reference records do not own contiguous areas",
            ));
        }
        let is_list = left.is_list || right.is_list;
        let mut surviving_records = 0usize;
        let mut output_areas = 0usize;
        let mut planned_cells = 0usize;
        let mut left_start = 0usize;

        // Public Area values may span several sheets, while the internal
        // RuntimeArea vector stores one physical plane per sheet. Runtime
        // records are appended in the same order as those planes, so derive
        // exact contiguous slices from each record's sheet extents. Matching
        // public bounds would conflate duplicate/overlapping records.
        for left_record in &left.records {
            self.scalar
                .charge_work(u64::try_from(left_record.areas.len()).unwrap_or(u64::MAX))?;
            let left_end = record_plane_end(left_record, left_start, left.areas.len())?;
            let left_areas = &left.areas[left_start..left_end];
            let mut right_start = 0usize;
            for right_record in &right.records {
                self.scalar
                    .charge_work(u64::try_from(right_record.areas.len()).unwrap_or(u64::MAX))?;
                let right_end = record_plane_end(right_record, right_start, right.areas.len())?;
                let right_areas = &right.areas[right_start..right_end];
                let mut pair_areas = 0usize;
                let mut pair_cells = 0usize;
                for left_area in left_areas {
                    for right_area in right_areas {
                        self.scalar.charge_work(1)?;
                        if left_area.sheet_index != right_area.sheet_index {
                            continue;
                        }
                        let Some(rect) = left_area.rect.intersection(right_area.rect) else {
                            continue;
                        };
                        pair_areas = pair_areas
                            .checked_add(1)
                            .ok_or_else(|| self.reference_limit_error(usize::MAX))?;
                        pair_cells = pair_cells
                            .checked_add(rect.count()?)
                            .ok_or_else(|| self.reference_limit_error(usize::MAX))?;
                    }
                }
                if pair_areas != 0 {
                    if is_list {
                        surviving_records = surviving_records
                            .checked_add(1)
                            .ok_or_else(|| self.reference_limit_error(usize::MAX))?;
                    }
                    output_areas = output_areas
                        .checked_add(pair_areas)
                        .ok_or_else(|| self.reference_limit_error(usize::MAX))?;
                    planned_cells = planned_cells
                        .checked_add(pair_cells)
                        .ok_or_else(|| self.reference_limit_error(usize::MAX))?;
                }
                right_start = right_end;
            }
            if right_start != right.areas.len() {
                return Err(EvaluationFailure::InvalidExpression(
                    "reference records do not cover contiguous areas",
                ));
            }
            left_start = left_end;
        }
        if left_start != left.areas.len() {
            return Err(EvaluationFailure::InvalidExpression(
                "reference records do not cover contiguous areas",
            ));
        }
        if output_areas == 0 {
            // Keep the historical formula-level null result while omitting
            // the old phantom empty record from the retained reference set.
            return Ok(RuntimeValue::Scalar(WorkingValue::Error(ScalarError::Null)));
        }
        self.check_reference_cells(planned_cells)?;
        self.check_reference_areas(output_areas)?;
        if is_list {
            self.check_reference_areas(surviving_records)?;
        }

        let mut result = RuntimeAreaSet::empty();
        result.is_list = is_list;
        ensure_capacity(
            &mut result.areas,
            &mut result.area_reservation,
            output_areas,
            self.limits.max_reference_areas,
            self.execution,
            &self.storage_budget,
            "formula reference areas",
        )?;
        if is_list {
            ensure_capacity(
                &mut result.records,
                &mut result.record_reservation,
                surviving_records,
                self.limits.max_reference_areas,
                self.execution,
                &self.storage_budget,
                "formula reference records",
            )?;
        } else {
            result.ensure_record_capacity(1, self)?;
            result.records.push(RuntimeReference {
                reference: None,
                areas: Vec::new(),
                reservation: None,
            });
        }

        let mut left_start = 0usize;
        for left_record in &left.records {
            self.scalar
                .charge_work(u64::try_from(left_record.areas.len()).unwrap_or(u64::MAX))?;
            let left_end = record_plane_end(left_record, left_start, left.areas.len())?;
            let left_areas = &left.areas[left_start..left_end];
            let mut right_start = 0usize;
            for right_record in &right.records {
                self.scalar
                    .charge_work(u64::try_from(right_record.areas.len()).unwrap_or(u64::MAX))?;
                let right_end = record_plane_end(right_record, right_start, right.areas.len())?;
                let right_areas = &right.areas[right_start..right_end];
                let mut pair_started = false;
                for left_area in left_areas {
                    for right_area in right_areas {
                        self.scalar.charge_work(1)?;
                        if left_area.sheet_index != right_area.sheet_index {
                            continue;
                        }
                        let Some(rect) = left_area.rect.intersection(right_area.rect) else {
                            continue;
                        };
                        let area = RuntimeArea {
                            sheet: left_area.sheet,
                            sheet_index: left_area.sheet_index,
                            rect,
                        };
                        if is_list {
                            if !pair_started {
                                result.records.push(RuntimeReference {
                                    reference: None,
                                    areas: Vec::new(),
                                    reservation: None,
                                });
                                pair_started = true;
                            }
                            result.push_raw(area, self)?;
                            result.append_record_area(&area, self)?;
                        } else {
                            result.push(area, self)?;
                        }
                    }
                }
                right_start = right_end;
            }
            if right_start != right.areas.len() {
                return Err(EvaluationFailure::InvalidExpression(
                    "reference records do not cover contiguous areas",
                ));
            }
            left_start = left_end;
        }
        Ok(RuntimeValue::Areas(result))
    }

    pub(super) fn union_ranges(
        &mut self,
        left: RuntimeValue<'expr>,
        right: RuntimeValue<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        if is_reference_error(&left) || is_reference_error(&right) {
            return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::Reference,
            )));
        }
        let mut left = self.coerce_reference_areas(left)?;
        let right = self.coerce_reference_areas(right)?;
        let area_count = left
            .areas
            .len()
            .checked_add(right.areas.len())
            .ok_or_else(|| self.reference_limit_error(usize::MAX))?;
        let record_count = left
            .records
            .len()
            .checked_add(right.records.len())
            .ok_or_else(|| self.reference_limit_error(usize::MAX))?;
        self.check_reference_areas(area_count)?;
        self.check_reference_areas(record_count)?;
        let planned_cells = left
            .cell_count
            .checked_add(right.cell_count)
            .ok_or_else(|| self.reference_limit_error(usize::MAX))?;
        self.check_reference_cells(planned_cells)?;
        left.append_set(right, self)?;
        Ok(RuntimeValue::Areas(left))
    }

    pub(super) fn coerce_reference_areas(
        &mut self,
        value: RuntimeValue<'expr>,
    ) -> EvaluationResult<RuntimeAreaSet<'expr>> {
        match value {
            RuntimeValue::Areas(areas) => Ok(areas),
            RuntimeValue::Scalar(WorkingValue::Error(ScalarError::Reference)) => {
                Ok(RuntimeAreaSet {
                    areas: Vec::new(),
                    records: Vec::new(),
                    cell_count: 0,
                    area_reservation: None,
                    record_reservation: None,
                    is_list: false,
                })
            },
            RuntimeValue::Empty
            | RuntimeValue::Missing
            | RuntimeValue::Scalar(_)
            | RuntimeValue::ScalarCell(_)
            | RuntimeValue::Array(_) => Err(EvaluationFailure::Unsupported(
                super::super::UnsupportedKind::ReferenceOperator,
            )),
        }
    }

    fn reference_limit_error(&self, observed: usize) -> EvaluationFailure {
        EvaluationFailure::ResourceLimit(self.local_limit(
            Resource::Objects,
            u64::try_from(observed).unwrap_or(u64::MAX),
            self.limits.max_reference_cells,
        ))
    }

    fn check_reference_cells(&self, cells: usize) -> EvaluationResult<()> {
        if cells > self.limits.max_reference_cells {
            return Err(self.reference_limit_error(cells));
        }
        Ok(())
    }

    fn check_reference_areas(&self, areas: usize) -> EvaluationResult<()> {
        if areas > self.limits.max_reference_areas {
            return Err(EvaluationFailure::ResourceLimit(self.local_limit(
                Resource::Objects,
                u64::try_from(areas).unwrap_or(u64::MAX),
                self.limits.max_reference_areas,
            )));
        }
        Ok(())
    }
}
