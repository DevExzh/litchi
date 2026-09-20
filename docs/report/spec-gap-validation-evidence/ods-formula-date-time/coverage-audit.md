# ODS date/time coverage audit (source locked; refreeze pending)

Status: **AUDIT COMPLETE / CURRENT CUSTODY RECEIPT PENDING**. This file records
the locked source identifiers, the historical diagnostic receipts, and the
remaining concrete contract gaps. The previous gate bundle was invalidated
because its corpus archive hash was wrong, so this audit does not promote the
coverage manifest or any historical gate result to PASS. The oracle vector
corpus is distinct from its executed replay receipt.

Contract SHA256: `cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f`.
Oracle vector corpus SHA256: `b33011089974b18b1fcc0b984adab839600acc98f09dd850e02522348316259f` (132 rows).

## Current byte identities

| artifact | SHA256 |
| --- | --- |
| `contract.md` | `cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f` |
| `coverage-scope.json` | `28a754ded86ad7524d0d94c5856757f775f52dee2f52dd603f8f66353fdbd114` |
| `coverage-requirements.json` | `18283eb96f31befb2c4e970196df550f329103642de87dd791704e47111ab46d` |
| `date/time production source` | `1da16aae011dd8c94fafc39fdf81cc5d9f8019024ab3860bdc00b4d28962e159` (`value/date_time.rs`) |
| `semantic source` | `d9644432fd5dcac9272dcd3f942e4b669865db22fb6fe866a706ebd05bf04ea1` |
| `resource source` | `fbf6ba638d7b80ebcda36d1091641ef45e42be07f9bc20a987583a3987fed411` |
| `oracle replay source` | `6ab4f5845c58aa88b5890abc766229a4970409cfe581b924228bb4b595e97a05` |
| `oracle vectors` | `b33011089974b18b1fcc0b984adab839600acc98f09dd850e02522348316259f` |
| `independent oracle checker` | `4884b8a0ff7fc7dc2a0259b4e7b4e24c4be169a166cd92982b728b2ed77eb64b` |
| historical `gates/history-provenance-typo/freeze.json` (invalidated) | `fe639f5fc5c72997711d5f37c69972aad95909e8f950a0bb8f5c72f09319ecd2` |
| historical `gates/history-provenance-typo/results.json` (invalidated) | `0a011cd5c6e6ee18e61c8327a3d27f1d41f9f3f6836d55c6148ebf56164ba2d3` |
| historical `gates/history-provenance-typo/ods-tests.log` (invalidated) | `c58cbb73c8939a3c81adda9f40111ed001f59a53a3153f28d7ab6e4cd62cb5bc` |
| historical `gates/history-harness-checksum/freeze.json` (diagnostic; archived) | `bf51f22e49c12427b2595e765f8e13c3bb84305b3d02958682ca4e9d8138b802` |
| historical `gates/history-harness-checksum/results.json` (diagnostic; archived) | `85fee359e7d3c70eba49d504c10238862a5bc5033d2b42d37add8414009a8996` |
| historical `gates/history-harness-checksum/ods-tests.log` (diagnostic; archived) | `2ce061ff2358f8f113d15d3d0afb6ae21741365d37edbf00ba8aa134f4f38a70` |
| historical `gates/history-harness-checksum/oracle-execution.json` (diagnostic; archived) | `e89122244a594b7797776b2f8a19955ea889aa6b10856119e49093a853d1339d` |

The gate rows are historical observations only. They must not be used as the
current source/gate receipt. A replacement freeze must recompute the source,
corpus, lock, gate, and replay identities together.

## Focused Rust identifiers

Semantic target (`crates/litchi-ods/tests/ods_formula_date_time_evaluation.rs`):

- `all_twenty_four_functions_have_scalar_value_differential_vectors`
- `exact_arity_and_domain_errors_cover_every_date_time_name`
- `formula_errors_propagate_through_each_argument_taking_date_function`
- `every_function_rejects_arguments_outside_its_declared_arity`
- `required_missing_and_optional_empty_slots_keep_distinct_profiles`
- `datedif_formats_trim_case_and_apply_signed_month_remainders`
- `date_bounds_fractional_floor_and_zero_offsets_follow_the_profile`
- `component_and_month_shift_edges_use_checked_profile_bounds`
- `days_converts_logicals_and_single_cell_references`
- `invalid_holiday_dates_and_leap_second_text_remain_typed_formula_values`
- `time_parameters_accept_finite_values_outside_the_date_domain`
- `retained_holiday_formula_error_supersedes_generated_scalar_conversion_error`
- `fixed_parsers_accept_grouped_numeric_fallback_and_mixed_fractions`
- `numeric_fallback_covers_exponent_percent_currency_and_fraction_forms`
- `yearfrac_and_iso_week_edges_follow_profile_procedures`
- `yearfrac_defaults_floors_dates_and_handles_thirty_first_rules`
- `timestamp_api_is_explicit_and_volatile_functions_are_deterministic`
- `matrix_date_arguments_preserve_shape_and_elementwise_results`
- `projected_datevalue_reference_preserves_each_date_and_reads_once_per_cell`
- `complete_holiday_and_workweek_sequences_survive_projected_if`
- `nested_munit_date_selector_remains_position_sensitive`
- `value_profile_does_not_widen_existing_value_conversion`

Resource target (`crates/litchi-ods/tests/ods_formula_date_time_limits.rs`):

- `reference_cell_and_array_limits_fail_before_reads_and_release_memory`
- `reference_list_sequence_refusal_is_read_free_but_invalid_reference_length_is_eager`
- `formula_errors_are_retained_while_later_sequence_failures_supersede_them`
- `cancellation_and_source_fences_publish_no_partial_date_result`
- `borrowed_date_text_respects_text_budget_and_typed_failures_are_not_catchable`
- `long_date_intervals_are_bounded_by_work_even_without_reference_reads`
- `workday_consumes_holiday_and_workweek_sequences_before_zero_and_all_off_results`
- `matrix_workday_typed_failure_publishes_no_partial_result`
- `long_grouped_parser_matrix_and_holiday_storage_release_resources`
- `explicit_timestamp_contexts_change_results_and_remain_fenced`
- `workday_cancellation_follows_complete_sequences_and_work_is_bounded`

Oracle replay target (`crates/litchi-ods/tests/ods_formula_date_time_oracle.rs`):

- `date_time_oracle_replays_all_contract_vectors` (replays the corpus once per row and checks all 24 function names).

Production unit target (`crates/litchi-ods/src/codec/formula/evaluation/value/date_time.rs`):

- `projected_timestamp_functions_cache_payloads_only_with_a_snapshot` (checks
  same-invocation cache reuse for NOW/TODAY/EASTERSUNDAY and typed refusal with
  no cache entry when the snapshot is absent).

## Function-to-corpus inventory and focused associations

The focused associations below are source identifiers, not receipt claims.
Every function also participates in the all-function differential and arity
tests listed above where shown. `NOW` and `TODAY` are zero-arity functions, so
there is no meaningful formula-error argument case for them; the common arity
suite covers their too-many-argument refusal.

| function | focused tests in current source | exact corpus case IDs |
| --- | --- | --- |
| `DATE` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`date_bounds_fractional_floor_and_zero_offsets_follow_the_profile` | `date.month_rollover`<br>`date.day_rollover_leap`<br>`date.no_synthetic_1900_day`<br>`date.fractional_arguments_truncate`<br>`date.minimum_profile_date`<br>`date.maximum_profile_date`<br>`date.invalid_month`<br>`date.checked_overflow` |
| `DATEDIF` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`required_missing_and_optional_empty_slots_keep_distinct_profiles`<br>`datedif_formats_trim_case_and_apply_signed_month_remainders` | `datedif.Y`<br>`datedif.M`<br>`datedif.D`<br>`datedif.MD`<br>`datedif.YM`<br>`datedif.YD`<br>`datedif.month_day_positive`<br>`datedif.reversed_interval`<br>`datedif.invalid_format` |
| `DATEVALUE` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`fixed_parsers_accept_grouped_numeric_fallback_and_mixed_fractions`<br>`numeric_fallback_covers_exponent_percent_currency_and_fraction_forms`<br>`date_bounds_fractional_floor_and_zero_offsets_follow_the_profile`<br>`projected_datevalue_reference_preserves_each_date_and_reads_once_per_cell`<br>`value_profile_does_not_widen_existing_value_conversion` | `datevalue.iso_date`<br>`datevalue.datetime_discards_time`<br>`datevalue.en_us_numeric`<br>`datevalue.two_digit_pivot`<br>`datevalue.english_month_name`<br>`datevalue.numeric_fallback`<br>`datevalue.numeric_fallback_small`<br>`datevalue.simple_fraction_fallback`<br>`datevalue.invalid_calendar_day` |
| `DAY` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`date_bounds_fractional_floor_and_zero_offsets_follow_the_profile`<br>`component_and_month_shift_edges_use_checked_profile_bounds` | `day.datetime_floor`<br>`day.iso_text`<br>`day.out_of_domain` |
| `DAYS` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`required_missing_and_optional_empty_slots_keep_distinct_profiles`<br>`days_converts_logicals_and_single_cell_references` | `days.retain_fraction`<br>`days.reversed`<br>`days.text_and_number` |
| `DAYS360` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`required_missing_and_optional_empty_slots_keep_distinct_profiles`<br>`date_bounds_fractional_floor_and_zero_offsets_follow_the_profile` | `days360.us_february_end`<br>`days360.us_reversed_signed`<br>`days360.european_swaps_with_sign`<br>`days360.european_31st` |
| `EASTERSUNDAY` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`timestamp_api_is_explicit_and_volatile_functions_are_deterministic` | `eastersunday.explicit_1583`<br>`eastersunday.explicit_2024`<br>`eastersunday.explicit_9956`<br>`eastersunday.invalid_year`<br>`eastersunday.timestamp_after_current_easter`<br>`eastersunday.missing_timestamp` |
| `EDATE` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`date_bounds_fractional_floor_and_zero_offsets_follow_the_profile`<br>`component_and_month_shift_edges_use_checked_profile_bounds` | `edate.leap_month_clamp`<br>`edate.nonleap_month_clamp`<br>`edate.negative_month`<br>`edate.fractional_month_truncates` |
| `EOMONTH` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`date_bounds_fractional_floor_and_zero_offsets_follow_the_profile`<br>`component_and_month_shift_edges_use_checked_profile_bounds` | `eomonth.leap_target`<br>`eomonth.previous_month`<br>`eomonth.invalid_domain` |
| `HOUR` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`time_parameters_accept_finite_values_outside_the_date_domain` | `hour.midday_fraction`<br>`hour.negative_time`<br>`hour.final_second` |
| `ISOWEEKNUM` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`yearfrac_and_iso_week_edges_follow_profile_procedures`<br>`component_and_month_shift_edges_use_checked_profile_bounds` | `isoweeknum.new_year_2021`<br>`isoweeknum.first_week_2021`<br>`isoweeknum.end_2020`<br>`isoweeknum.2015_first_thursday` |
| `MINUTE` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`time_parameters_accept_finite_values_outside_the_date_domain`<br>`invalid_holiday_dates_and_leap_second_text_remain_typed_formula_values` | `minute.half_second_rounds_up`<br>`minute.just_below_half`<br>`minute.negative_half_day`<br>`minute.end_of_day_wrap` |
| `MONTH` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`date_bounds_fractional_floor_and_zero_offsets_follow_the_profile`<br>`component_and_month_shift_edges_use_checked_profile_bounds` | `month.datetime_floor`<br>`month.text_name`<br>`month.invalid_text` |
| `NETWORKDAYS` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`required_missing_and_optional_empty_slots_keep_distinct_profiles`<br>`retained_holiday_formula_error_supersedes_generated_scalar_conversion_error`<br>`invalid_holiday_dates_and_leap_second_text_remain_typed_formula_values`<br>`complete_holiday_and_workweek_sequences_survive_projected_if` | `networkdays.default_week`<br>`networkdays.reversed_interval`<br>`networkdays.custom_all_workdays`<br>`networkdays.duplicate_holiday`<br>`networkdays.no_workday_sequence`<br>`networkdays.formula_error_holiday`<br>`networkdays.reference_list_refusal` |
| `NOW` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`timestamp_api_is_explicit_and_volatile_functions_are_deterministic`<br>`explicit_timestamp_contexts_change_results_and_remain_fenced` | `now.explicit_timestamp`<br>`now.missing_timestamp` |
| `SECOND` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`time_parameters_accept_finite_values_outside_the_date_domain`<br>`invalid_holiday_dates_and_leap_second_text_remain_typed_formula_values` | `second.half_second_wrap`<br>`second.just_below_boundary`<br>`second.negative_fraction`<br>`second.final_day_fraction` |
| `TIME` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`time_parameters_accept_finite_values_outside_the_date_domain` | `time.direct_fraction`<br>`time.negative_fraction`<br>`time.multi_day`<br>`time.overflow` |
| `TIMEVALUE` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`fixed_parsers_accept_grouped_numeric_fallback_and_mixed_fractions`<br>`numeric_fallback_covers_exponent_percent_currency_and_fraction_forms`<br>`time_parameters_accept_finite_values_outside_the_date_domain`<br>`invalid_holiday_dates_and_leap_second_text_remain_typed_formula_values` | `timevalue.clock_fraction`<br>`timevalue.datetime_fraction`<br>`timevalue.numeric_fallback`<br>`timevalue.simple_fraction_fallback`<br>`timevalue.invalid_24_hour`<br>`timevalue.date_only_rejected`<br>`timevalue.invalid_second` |
| `TODAY` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`timestamp_api_is_explicit_and_volatile_functions_are_deterministic`<br>`explicit_timestamp_contexts_change_results_and_remain_fenced` | `today.explicit_timestamp`<br>`today.missing_timestamp` |
| `WEEKDAY` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`required_missing_and_optional_empty_slots_keep_distinct_profiles`<br>`component_and_month_shift_edges_use_checked_profile_bounds` | `weekday.type_1`<br>`weekday.type_2`<br>`weekday.type_3`<br>`weekday.type_11`<br>`weekday.type_12`<br>`weekday.type_13`<br>`weekday.type_14`<br>`weekday.type_15`<br>`weekday.type_16`<br>`weekday.type_17`<br>`weekday.omitted_type`<br>`weekday.invalid_type` |
| `WEEKNUM` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`required_missing_and_optional_empty_slots_keep_distinct_profiles` | `weeknum.mode_1`<br>`weeknum.mode_2`<br>`weeknum.mode_11`<br>`weeknum.mode_12`<br>`weeknum.mode_13`<br>`weeknum.mode_14`<br>`weeknum.mode_15`<br>`weeknum.mode_16`<br>`weeknum.mode_17`<br>`weeknum.mode_21`<br>`weeknum.mode_150`<br>`weeknum.omitted_mode`<br>`weeknum.noninteger_mode` |
| `WORKDAY` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`required_missing_and_optional_empty_slots_keep_distinct_profiles`<br>`date_bounds_fractional_floor_and_zero_offsets_follow_the_profile`<br>`component_and_month_shift_edges_use_checked_profile_bounds`<br>`complete_holiday_and_workweek_sequences_survive_projected_if`<br>`workday_cancellation_follows_complete_sequences_and_work_is_bounded` | `workday.forward_default`<br>`workday.backward_default`<br>`workday.zero_preserves_fraction`<br>`workday.holiday_skip`<br>`workday.custom_all_workdays`<br>`workday.all_off_nonzero`<br>`workday.reference_list_refusal` |
| `YEAR` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`date_bounds_fractional_floor_and_zero_offsets_follow_the_profile`<br>`component_and_month_shift_edges_use_checked_profile_bounds` | `year.two_digit_pivot`<br>`year.datetime`<br>`year.minimum_profile`<br>`year.invalid_text` |
| `YEARFRAC` | `all_twenty_four_functions_have_scalar_value_differential_vectors`<br>`exact_arity_and_domain_errors_cover_every_date_time_name`<br>`every_function_rejects_arguments_outside_its_declared_arity`<br>`formula_errors_propagate_through_each_argument_taking_date_function`<br>`required_missing_and_optional_empty_slots_keep_distinct_profiles`<br>`yearfrac_and_iso_week_edges_follow_profile_procedures`<br>`yearfrac_defaults_floors_dates_and_handles_thirty_first_rules` | `yearfrac.basis_0_30us`<br>`yearfrac.basis_1_actual_leap_year`<br>`yearfrac.basis_2_actual_360`<br>`yearfrac.basis_3_actual_365`<br>`yearfrac.basis_4_30e`<br>`yearfrac.reversed_is_nonnegative`<br>`yearfrac.invalid_basis` |

## Explicit per-function residuals

The source-locked focused targets and corpus cover the direct function cases. The
three rows below identify residuals whose behavior is supplied by shared
conversion/kernel paths and source review rather than by another per-function
Cartesian test. They are recorded for traceability and are not additional
freeze blockers.

| function | disposition | missing or partial requirement evidence | existing corpus anchors |
| --- | --- | --- | --- |
| `DAY` | **source-reviewed; shared boundary** | Component and malformed/out-of-domain cases are direct. The lower profile boundary is exercised through the shared DateParam/calendar kernel by lower-bound MONTH/YEAR cases and the upper-bound DAY case; formula-error and single-cell conversion use the shared argument path. | `day.datetime_floor`, `day.iso_text`, `day.out_of_domain`, `component_and_month_shift_edges_use_checked_profile_bounds` |
| `DAYS360` | **source-reviewed; shared argument path** | US/European, February/31st, fractional flooring, and optional-empty behavior are direct. Method-slot formula errors use the same common source-order argument propagation exercised across the family; no separate method-only Cartesian case is required. | `days360.us_february_end`, `days360.us_reversed_signed`, `days360.european_swaps_with_sign`, `days360.european_31st`, `formula_errors_propagate_through_each_argument_taking_date_function` |
| `NOW` | **source-reviewed** | Explicit and missing timestamp behavior is covered; ambient-state absence is established by the source-proof references below. | `now.explicit_timestamp`, `now.missing_timestamp` |
| `TIME` | **source-reviewed; shared finite-number path** | Fraction, negative, multi-day, checked overflow, and formula-error precedence are direct. Non-finite input is rejected by the shared finite Number bridge and checked kernel; formula syntax cannot supply a non-finite literal in this target. | `time.direct_fraction`, `time.negative_fraction`, `time.multi_day`, `time.overflow`, `time_parameters_accept_finite_values_outside_the_date_domain` |
| `TODAY` | **source-reviewed** | Explicit/missing timestamp behavior is covered; ambient-state absence is established by the source-proof references below. | `today.explicit_timestamp`, `today.missing_timestamp` |

## Cross-cutting audit

| disposition | requirement area | current observation |
| --- | --- | --- |
| **mapped; historical receipt observed; current refreeze pending** | `formula errors propagate in source order while typed failures remain evaluation failures` | Scalar formula-error identity is exercised for all 22 argument-taking functions; retained sequence tests exercise source-order retention and typed provider precedence for NETWORKDAYS and WORKDAY, and the matrix typed-failure case exercises no-partial publication. Shared argument dispatch supplies the common behavior; no per-function matrix Cartesian product is required. |
| **mapped; historical receipt observed; current refreeze pending** | `fixed-profile borrowed-text parsing...` | `fixed_parsers_accept_grouped_numeric_fallback_and_mixed_fractions` and `numeric_fallback_covers_exponent_percent_currency_and_fraction_forms` cover the retained grammar; `long_grouped_parser_matrix_and_holiday_storage_release_resources` covers borrowed text, matrix parsing, and bounded scratch. The historical package receipt records these identifiers, but it is invalidated. |
| **mapped; historical receipt observed; current refreeze pending** | `scalar conversion...` | ReferenceList zero-read refusal is explicit for NETWORKDAYS/WORKDAY; `projected_datevalue_reference_preserves_each_date_and_reads_once_per_cell` and `days_converts_logicals_and_single_cell_references` cover projected and single-cell conversion. The shared DateParam/shape gates are source-reviewed; no per-function Cartesian matrix is required. |
| **mapped; historical receipt observed; current refreeze pending** | `matrix broadcasting...` | Matrix tests cover DATE/TIME/WEEKDAY/IF and projected DATEVALUE; `matrix_workday_typed_failure_publishes_no_partial_result` covers typed failure and publication fencing for WORKDAY. The shared matrix path is source-reviewed rather than expanded into an unbounded family matrix. |
| **mapped; historical receipt observed; current refreeze pending** | `complete DateSequence...` | NETWORKDAYS and WORKDAY have resolver-backed complete sequence tests, including WORKDAY later typed failure and `workday_cancellation_follows_complete_sequences_and_work_is_bounded`. |
| **mapped; historical receipt observed; current refreeze pending** | `every inspected reference cell...` | The limits fixture records read order and limits; production source proves the per-cell work/read/cancellation ordering. The historical package receipt records the limits target, but it is invalidated. |
| **mapped; historical receipt observed; current refreeze pending** | `bounded parser, sequence...` | Limits cover memory, array/reference admission, text bytes, long parser scratch, matrix parsing, complete sequences, cancellation, and work; production reservations and paired drop order are source-reviewed. |
| **mapped; historical receipt observed; current refreeze pending** | `demand-cache entries retain payload/error, position, projected-shape, source, cancellation, and timestamp identity` | The per-evaluation cache and scalar/error payload policy are source-bound; `nested_munit_date_selector_remains_position_sensitive` is direct, `projected_timestamp_functions_cache_payloads_only_with_a_snapshot` proves same-invocation reuse and no-snapshot refusal, and `explicit_timestamp_contexts_change_results_and_remain_fenced` covers cross-invocation timestamp/fence identity. |
| **mapped; historical receipt observed; current refreeze pending** | `timestamp-backed volatile cache hits require the same explicit timestamp and fences` | The classifier admits NOW/TODAY/EASTERSUNDAY cache use only with an explicit timestamp; `projected_timestamp_functions_cache_payloads_only_with_a_snapshot` proves reuse only within one evaluator snapshot, while the timestamp-context test checks changed values plus source/cancellation fences. |
| **source-reviewed** | `no package, workbook, formula-cache, locale, filesystem, or process mutation` | The value API declares a read-only resolver and explicitly excludes dependency-graph recalculation, formula-cell execution, result publication, and cache publication; locale/clock are caller-profile inputs. This architectural source proof is the appropriate evidence boundary. |
| **mapped; historical receipt observed; current refreeze pending** | `independent oracle and native...` | The historical package log executes `date_time_oracle_replays_all_contract_vectors`; the native fixture records 132 rows with explicit parity/divergence policy. These corroborate the profile without replacing it, and the invalidated bundle cannot serve as the current receipt. |
| **partial; current receipt pending** | `replacement isolated gate receipts and retained raw performance controls with regressions disclosed` | The historical gate bundle recorded seven successful stages and 1,745 package tests, but it was invalidated by the corpus archive hash typo. A replacement custody bundle must bind the corrected corpus, source map, lock, gate logs, oracle replay, and retained performance evidence together. |

## Cross-cutting source-proof references

The following requirements have architectural evidence in the production source;
these references should be bound as source-review evidence rather than forcing a
mock test to prove an authority boundary that the evaluator does not expose.
They do not change the pending status of the coverage manifest.

| requirement area | source proof | remaining runtime evidence |
| --- | --- | --- |
| Resolver read-only and no package/workbook/formula-cache/process mutation | `crates/litchi-ods/src/codec/formula/evaluation/value.rs:8-12` defines the resolver as read-only and forbids I/O, refresh, and recursive formula evaluation; `value.rs:140-149` excludes dependency graphs, recalculation, spill/publication, and cache publication; `value.rs:473-480` requires coherent immutable provider observations. | Source review is sufficient for the authority boundary; no synthetic I/O mock is required. Current custody remains pending refreeze. |
| Locale-independent and clock-independent profile | `crates/litchi-ods/src/codec/formula/evaluation/timestamp.rs:1-7` requires caller-supplied timestamps; `crates/litchi-ods/src/codec/formula/evaluation.rs:338-378` stores only explicit profile options; `value.rs:64-71` documents that neither document locale nor host clock is consulted. | Semantic/oracle values and the explicit timestamp tests remain runtime evidence; ambient-state absence remains source evidence. |
| Per-evaluation demand-cache lifetime and scalar/error payload policy | `crates/litchi-ods/src/codec/formula/evaluation/value.rs:1992-2038` constructs fresh cache vectors in each `ValueEvaluator`; `value.rs:1732-1763` defines bounded scalar/error cache payloads; `value.rs:4197-4258` performs node-keyed get/put and refuses references, arrays, and borrowed text. | `nested_munit_date_selector_remains_position_sensitive` is direct; `projected_timestamp_functions_cache_payloads_only_with_a_snapshot` proves same-invocation reuse and no-snapshot refusal; `explicit_timestamp_contexts_change_results_and_remain_fenced` covers cross-invocation timestamp identity and fences. |
| Volatile cache eligibility and explicit timestamp identity | `crates/litchi-ods/src/codec/formula/evaluation/value/date_time.rs:66-121` rejects volatile cacheability without an explicit timestamp and rejects projected computed parameters; `value.rs:8177-8202` applies the same get/apply/put path to statistical and date/time reducers. | `projected_timestamp_functions_cache_payloads_only_with_a_snapshot` proves that NOW/TODAY/EASTERSUNDAY cache only with a snapshot; the timestamp-context test changes only the explicit timestamp and checks source/cancellation fences. |
| Source-version and final-cancellation fences | `crates/litchi-ods/src/codec/formula/evaluation/value.rs:1463-1527` probes source identity before and after evaluation and checks cancellation before publication. | `cancellation_and_source_fences_publish_no_partial_date_result` directly exercises the resolver-backed path. |
| Streamed reference-cell work/read/cancellation and borrowed text | `crates/litchi-ods/src/codec/formula/evaluation/value.rs:8940-8975` charges work and checks the reference-cell limit/cancellation before and after each resolver read; `value.rs:9259-9270` scans a retained rectangular descriptor in row-major order; `value.rs:9281-9311` enforces text bytes and keeps `TextValue::borrowed`. | Limits read-order, text-budget, and typed-failure cases are recorded in the invalidated historical receipt; the replacement receipt remains pending. |
| Shared DateParam/finite-number conversion and checked date kernel | `crates/litchi-ods/src/codec/formula/evaluation/value/date_time.rs:877-945` preserves formula errors, rejects non-finite/out-of-domain serials, and routes text through the shared parser; `crates/litchi-ods/src/codec/formula/evaluation/date_time/kernel.rs:370-397` checks basis/domain results for finiteness. | Component-boundary, logical/reference, overflow, and parser cases were recorded in the invalidated historical receipt; the replacement receipt remains pending. No per-function Cartesian expansion is required. |

## Receipt boundary

- The historical custody bundle is archived at
  `gates/history-provenance-typo/`. Its `freeze.json` SHA-256 is
  `fe639f5fc5c72997711d5f37c69972aad95909e8f950a0bb8f5c72f09319ecd2`, and
  it recorded candidate commit `1a666f15a3aa50a460a97f52f4d953e16ce41538`
  with isolated lock hash
  `58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`.
- That bundle is explicitly **INVALIDATED** because its provenance contained
  an incorrect corpus `archive_sha256`. Its historical `results.json` has
  SHA-256 `0a011cd5c6e6ee18e61c8327a3d27f1d41f9f3f6836d55c6148ebf56164ba2d3`,
  its `ods-tests.log` has SHA-256
  `c58cbb73c8939a3c81adda9f40111ed001f59a53a3153f28d7ab6e4cd62cb5bc`, and
  its `oracle-execution.json` has SHA-256
  `5a91dac821c0c70816c818781fdcdd58ab16bb361deca5afc1536098740ea435`.
  The log's 1,745 passed tests, including the 22 semantic tests, 11 resource
  tests, 132-vector replay, and timestamp-cache unit test, are diagnostic
  history only and do not establish current custody.
- The subsequent seven-gate run is archived at
  `gates/history-harness-checksum/`: it recorded seven successful gate stages,
  the same 1,745 package tests, and a corrected 132-row oracle receipt. That
  bundle is diagnostic only because the later performance capture failed a
  harness checksum in scalar SUM matrix mode; the failure is a harness
  custody issue, and no production source change is inferred from it.
- The corrected corpus remains independently hashed as
  `b33011089974b18b1fcc0b984adab839600acc98f09dd850e02522348316259f`.
  A replacement freeze must bind that corpus, the source map, isolated lock,
  gate logs, oracle replay, native records, and retained performance results
  together before current evidence is promoted.
- `coverage-requirements.json` remains pending with empty evidence bindings;
  no synthetic receipt or PASS status was added.
