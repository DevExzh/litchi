# OLE2 name handoff 0554 results review

**Status:** final diagnostic review; the candidate is rejected by the frozen native controls and the mandatory four-XLS supplement gate. Profile and instruction mechanism checks pass diagnostically, and final restored-baseline quality passes.

The canonical metrics report completed with status `pass`, but its frozen main admission result is false. The many-small CFB native gate and all allocator checks pass. All eight mandatory primary XLS native p50 checks fail, and 19 native control checks fail. The review retains every 83 candidate adverse rows and every 43 absolute-over-5% same-build drift rows individually in [`adverse-review.json`](adverse-review.json).

Native timing comes only from the non-instrumented executable. Allocator-instrumented elapsed samples are excluded from latency interpretation. RSS is a whole-child GNU time observation. OLE2/OOXML remains the priority; ODF is deferred and iWork is excluded.

## Gate findings

| Gate | Result | Checks | Failed |
| --- | --- | ---: | ---: |
| `primary_many_small` | **pass** | 4 | 0 |
| `primary_xls_supplement` | **fail** | 8 | 8 |
| `native_controls` | **fail** | 52 | 19 |
| `allocation` | **pass** | 72 | 0 |

The eight failed primary XLS rows are the supplement cases `xls_source_backed_open`, `xls_source_backed_open_one_cell`, `xls_owned_source_open`, and `xls_owned_source_open_one_cell`, each in repeats 1 and 2. Their p50 changes range from +2.307401% to +10.927299%; the required improvement is at least 3% in both repeats. The 19 native control failures are p50/mean rows for the owned and source-backed XLS cases, with candidate increases up to +10.927299%.

## Retained rows

The review covers **126** flags: **83** candidate adverse comparisons and **43** same-build repeat drifts. Every row below has the exact canonical object in `adverse-review.json.original`, plus an individual interpretation and disposition.

### Candidate adverse comparisons

| ID | Group/case/shape | Repeat | Metric | Baseline | Candidate | Change | Disposition |
| --- | --- | ---: | --- | ---: | ---: | ---: | --- |
| 0554-review-0001 | cfb/cfb_open/few-large | 1 | `elapsed_ns.standard_deviation` | 1954.8202091686437 | 3259.960794865526 | +66.765249% | `retain_unresolved` |
| 0554-review-0002 | cfb/cfb_open/tiny | 1 | `elapsed_ns.max` | 13230 | 215551 | +1529.259259% | `retain_unresolved` |
| 0554-review-0003 | cfb/cfb_open/tiny | 1 | `elapsed_ns.standard_deviation` | 395.6405262128514 | 6762.666444271481 | +1609.295685% | `retain_unresolved` |
| 0554-review-0004 | cfb/cfb_open/tiny | 1 | `elapsed_ns.confidence_interval_95.upper` | 2373.513348481214 | 2851.4361402982245 | +20.135669% | `retain_unresolved` |
| 0554-review-0005 | xls/xls_eager_open_one_cell/xls | 1 | `elapsed_ns.standard_deviation` | 9703.673992078773 | 10931.324781956015 | +12.651402% | `retain_unresolved` |
| 0554-review-0006 | xls/xls_owned_source_open/xls | 1 | `elapsed_ns.min` | 88840 | 94251 | +6.090725% | `retain_unresolved` |
| 0554-review-0007 | xls/xls_owned_source_open/xls | 1 | `elapsed_ns.p50` | 96465 | 105060 | +8.909967% | `retain_unresolved` |
| 0554-review-0008 | xls/xls_owned_source_open/xls | 1 | `elapsed_ns.p95` | 109741 | 118360 | +7.853947% | `retain_unresolved` |
| 0554-review-0009 | xls/xls_owned_source_open/xls | 1 | `elapsed_ns.p99` | 116210 | 122690 | +5.576112% | `retain_unresolved` |
| 0554-review-0010 | xls/xls_owned_source_open/xls | 1 | `elapsed_ns.mean` | 98056.13299999996 | 105826.46500000004 | +7.924371% | `retain_unresolved` |
| 0554-review-0011 | xls/xls_owned_source_open/xls | 1 | `elapsed_ns.confidence_interval_95.lower` | 97185.51901430624 | 105494.25073962832 | +8.549352% | `retain_unresolved` |
| 0554-review-0012 | xls/xls_owned_source_open/xls | 1 | `elapsed_ns.confidence_interval_95.upper` | 98926.74698569368 | 106158.67926037176 | +7.310391% | `retain_unresolved` |
| 0554-review-0013 | xls/xls_owned_source_open_list_worksheets/xls | 1 | `elapsed_ns.min` | 87641 | 96410 | +10.005591% | `retain_unresolved` |
| 0554-review-0014 | xls/xls_owned_source_open_list_worksheets/xls | 1 | `elapsed_ns.p50` | 96550 | 105360 | +9.124806% | `retain_unresolved` |
| 0554-review-0015 | xls/xls_owned_source_open_list_worksheets/xls | 1 | `elapsed_ns.p95` | 109631 | 118991 | +8.537731% | `retain_unresolved` |
| 0554-review-0016 | xls/xls_owned_source_open_list_worksheets/xls | 1 | `elapsed_ns.p99` | 115261 | 122400 | +6.193769% | `retain_unresolved` |
| 0554-review-0017 | xls/xls_owned_source_open_list_worksheets/xls | 1 | `elapsed_ns.max` | 119921 | 145681 | +21.480808% | `retain_unresolved` |
| 0554-review-0018 | xls/xls_owned_source_open_list_worksheets/xls | 1 | `elapsed_ns.mean` | 97196.0079999999 | 106503.93700000015 | +9.576452% | `retain_unresolved` |
| 0554-review-0019 | xls/xls_owned_source_open_list_worksheets/xls | 1 | `elapsed_ns.confidence_interval_95.lower` | 96854.80924863726 | 106190.27975368904 | +9.638624% | `retain_unresolved` |
| 0554-review-0020 | xls/xls_owned_source_open_list_worksheets/xls | 1 | `elapsed_ns.confidence_interval_95.upper` | 97537.20675136254 | 106817.59424631126 | +9.514715% | `retain_unresolved` |
| 0554-review-0021 | xls/xls_owned_source_open_one_cell/xls | 1 | `elapsed_ns.min` | 89221 | 96030 | +7.631611% | `retain_unresolved` |
| 0554-review-0022 | xls/xls_owned_source_open_one_cell/xls | 1 | `elapsed_ns.p50` | 99560 | 107705 | +8.180996% | `retain_unresolved` |
| 0554-review-0023 | xls/xls_owned_source_open_one_cell/xls | 1 | `elapsed_ns.p95` | 113390 | 121570 | +7.214040% | `retain_unresolved` |
| 0554-review-0024 | xls/xls_owned_source_open_one_cell/xls | 1 | `elapsed_ns.p99` | 119251 | 125451 | +5.199118% | `retain_unresolved` |
| 0554-review-0025 | xls/xls_owned_source_open_one_cell/xls | 1 | `elapsed_ns.max` | 126070 | 493212 | +291.220750% | `retain_unresolved` |
| 0554-review-0026 | xls/xls_owned_source_open_one_cell/xls | 1 | `elapsed_ns.mean` | 100001.44099999999 | 109158.71500000011 | +9.157142% | `retain_unresolved` |
| 0554-review-0027 | xls/xls_owned_source_open_one_cell/xls | 1 | `elapsed_ns.standard_deviation` | 6070.296482372452 | 13257.151888394677 | +118.393812% | `retain_unresolved` |
| 0554-review-0028 | xls/xls_owned_source_open_one_cell/xls | 1 | `elapsed_ns.confidence_interval_95.lower` | 99624.75066003509 | 108336.04660706292 | +8.744108% | `retain_unresolved` |
| 0554-review-0029 | xls/xls_owned_source_open_one_cell/xls | 1 | `elapsed_ns.confidence_interval_95.upper` | 100378.1313399649 | 109981.3833929373 | +9.567076% | `retain_unresolved` |
| 0554-review-0030 | xls/xls_semantic_open/xls | 1 | `elapsed_ns.max` | 461952 | 486562 | +5.327393% | `retain_unresolved` |
| 0554-review-0031 | xls/xls_source_backed_open/xls | 1 | `elapsed_ns.min` | 94921 | 101240 | +6.657115% | `retain_unresolved` |
| 0554-review-0032 | xls/xls_source_backed_open/xls | 1 | `elapsed_ns.p50` | 101855 | 112985 | +10.927299% | `retain_unresolved` |
| 0554-review-0033 | xls/xls_source_backed_open/xls | 1 | `elapsed_ns.p95` | 116951 | 126001 | +7.738284% | `retain_unresolved` |
| 0554-review-0034 | xls/xls_source_backed_open/xls | 1 | `elapsed_ns.p99` | 123230 | 130150 | +5.615516% | `retain_unresolved` |
| 0554-review-0035 | xls/xls_source_backed_open/xls | 1 | `elapsed_ns.max` | 137211 | 148811 | +8.454133% | `retain_unresolved` |
| 0554-review-0036 | xls/xls_source_backed_open/xls | 1 | `elapsed_ns.mean` | 103665.77700000002 | 113247.3529999999 | +9.242757% | `retain_unresolved` |
| 0554-review-0037 | xls/xls_source_backed_open/xls | 1 | `elapsed_ns.confidence_interval_95.lower` | 103276.75626422475 | 112896.5069386922 | +9.314536% | `retain_unresolved` |
| 0554-review-0038 | xls/xls_source_backed_open/xls | 1 | `elapsed_ns.confidence_interval_95.upper` | 104054.79773577528 | 113598.1990613076 | +9.171515% | `retain_unresolved` |
| 0554-review-0039 | xls/xls_source_backed_open_list_worksheets/xls | 1 | `elapsed_ns.min` | 95381 | 103650 | +8.669442% | `retain_unresolved` |
| 0554-review-0040 | xls/xls_source_backed_open_list_worksheets/xls | 1 | `elapsed_ns.p50` | 102405 | 113510 | +10.844197% | `retain_unresolved` |
| 0554-review-0041 | xls/xls_source_backed_open_list_worksheets/xls | 1 | `elapsed_ns.p95` | 116420 | 126720 | +8.847277% | `retain_unresolved` |
| 0554-review-0042 | xls/xls_source_backed_open_list_worksheets/xls | 1 | `elapsed_ns.p99` | 122580 | 130001 | +6.054006% | `retain_unresolved` |
| 0554-review-0043 | xls/xls_source_backed_open_list_worksheets/xls | 1 | `elapsed_ns.max` | 132080 | 398172 | +201.462750% | `retain_unresolved` |
| 0554-review-0044 | xls/xls_source_backed_open_list_worksheets/xls | 1 | `elapsed_ns.mean` | 104049.22399999999 | 115185.15900000006 | +10.702564% | `retain_unresolved` |
| 0554-review-0045 | xls/xls_source_backed_open_list_worksheets/xls | 1 | `elapsed_ns.standard_deviation` | 6024.164693306102 | 10161.359772596714 | +68.676660% | `retain_unresolved` |
| 0554-review-0046 | xls/xls_source_backed_open_list_worksheets/xls | 1 | `elapsed_ns.confidence_interval_95.lower` | 103675.39635370369 | 114554.59901362042 | +10.493524% | `retain_unresolved` |
| 0554-review-0047 | xls/xls_source_backed_open_list_worksheets/xls | 1 | `elapsed_ns.confidence_interval_95.upper` | 104423.05164629628 | 115815.7189863797 | +10.910108% | `retain_unresolved` |
| 0554-review-0048 | xls/xls_source_backed_open_one_cell/xls | 1 | `elapsed_ns.min` | 98421 | 107050 | +8.767438% | `retain_unresolved` |
| 0554-review-0049 | xls/xls_source_backed_open_one_cell/xls | 1 | `elapsed_ns.p50` | 109555 | 117330 | +7.096892% | `retain_unresolved` |
| 0554-review-0050 | xls/xls_source_backed_open_one_cell/xls | 1 | `elapsed_ns.p95` | 123880 | 131540 | +6.183403% | `retain_unresolved` |
| 0554-review-0051 | xls/xls_source_backed_open_one_cell/xls | 1 | `elapsed_ns.mean` | 110201.52800000006 | 118826.134 | +7.826213% | `retain_unresolved` |
| 0554-review-0052 | xls/xls_source_backed_open_one_cell/xls | 1 | `elapsed_ns.confidence_interval_95.lower` | 109820.02752004325 | 118491.44286344085 | +7.896024% | `retain_unresolved` |
| 0554-review-0053 | xls/xls_source_backed_open_one_cell/xls | 1 | `elapsed_ns.confidence_interval_95.upper` | 110583.02847995688 | 119160.82513655916 | +7.756883% | `retain_unresolved` |
| 0554-review-0054 | xls/xls_eager_open_list_worksheets/xls | 2 | `elapsed_ns.standard_deviation` | 10347.244527690096 | 11183.682945028988 | +8.083683% | `retain_unresolved` |
| 0554-review-0055 | xls/xls_eager_open_one_cell/xls | 2 | `elapsed_ns.standard_deviation` | 8645.067343713246 | 10788.14116665784 | +24.789556% | `retain_unresolved` |
| 0554-review-0056 | xls/xls_owned_source_open/xls | 2 | `elapsed_ns.p50` | 94250 | 99160 | +5.209549% | `retain_unresolved` |
| 0554-review-0057 | xls/xls_owned_source_open/xls | 2 | `elapsed_ns.p95` | 106640 | 114041 | +6.940173% | `retain_unresolved` |
| 0554-review-0058 | xls/xls_owned_source_open/xls | 2 | `elapsed_ns.max` | 124130 | 134821 | +8.612745% | `retain_unresolved` |
| 0554-review-0059 | xls/xls_owned_source_open/xls | 2 | `elapsed_ns.standard_deviation` | 5368.832261306065 | 5753.214484782172 | +7.159513% | `retain_unresolved` |
| 0554-review-0060 | xls/xls_owned_source_open_list_worksheets/xls | 2 | `elapsed_ns.p50` | 94175 | 99595 | +5.755243% | `retain_unresolved` |
| 0554-review-0061 | xls/xls_owned_source_open_list_worksheets/xls | 2 | `elapsed_ns.p95` | 105470 | 114540 | +8.599602% | `retain_unresolved` |
| 0554-review-0062 | xls/xls_owned_source_open_list_worksheets/xls | 2 | `elapsed_ns.p99` | 112350 | 119210 | +6.105919% | `retain_unresolved` |
| 0554-review-0063 | xls/xls_owned_source_open_list_worksheets/xls | 2 | `elapsed_ns.max` | 131581 | 147011 | +11.726617% | `retain_unresolved` |
| 0554-review-0064 | xls/xls_owned_source_open_list_worksheets/xls | 2 | `elapsed_ns.mean` | 95313.1909999999 | 100332.18599999983 | +5.265793% | `retain_unresolved` |
| 0554-review-0065 | xls/xls_owned_source_open_list_worksheets/xls | 2 | `elapsed_ns.standard_deviation` | 5354.665396645193 | 5972.382866537379 | +11.536061% | `retain_unresolved` |
| 0554-review-0066 | xls/xls_owned_source_open_list_worksheets/xls | 2 | `elapsed_ns.confidence_interval_95.lower` | 94980.90892209516 | 99961.5716586856 | +5.243857% | `retain_unresolved` |
| 0554-review-0067 | xls/xls_owned_source_open_list_worksheets/xls | 2 | `elapsed_ns.confidence_interval_95.upper` | 95645.47307790465 | 100702.80034131405 | +5.287576% | `retain_unresolved` |
| 0554-review-0068 | xls/xls_owned_source_open_one_cell/xls | 2 | `elapsed_ns.p50` | 96991 | 103330 | +6.535658% | `retain_unresolved` |
| 0554-review-0069 | xls/xls_owned_source_open_one_cell/xls | 2 | `elapsed_ns.p95` | 109401 | 116530 | +6.516394% | `retain_unresolved` |
| 0554-review-0070 | xls/xls_owned_source_open_one_cell/xls | 2 | `elapsed_ns.mean` | 98805.67399999996 | 103985.07599999994 | +5.242009% | `retain_unresolved` |
| 0554-review-0071 | xls/xls_owned_source_open_one_cell/xls | 2 | `elapsed_ns.confidence_interval_95.lower` | 97967.71727100783 | 103641.51146337696 | +5.791494% | `retain_unresolved` |
| 0554-review-0072 | xls/xls_source_backed_open_list_worksheets/xls | 2 | `elapsed_ns.max` | 145420 | 272611 | +87.464585% | `retain_unresolved` |
| 0554-review-0073 | xls/xls_source_backed_open_list_worksheets/xls | 2 | `elapsed_ns.standard_deviation` | 5657.178894087641 | 7916.599763947477 | +39.939003% | `retain_unresolved` |
| 0554-review-0074 | xls/xls_source_backed_open_one_cell/xls | 2 | `elapsed_ns.min` | 98771 | 109331 | +10.691397% | `retain_unresolved` |
| 0554-review-0075 | xls/xls_source_backed_open_one_cell/xls | 2 | `elapsed_ns.p50` | 107705 | 113110 | +5.018337% | `retain_unresolved` |
| 0554-review-0076 | xls/xls_source_backed_open_one_cell/xls | 2 | `elapsed_ns.p95` | 120441 | 130111 | +8.028827% | `retain_unresolved` |
| 0554-review-0077 | xls/xls_source_backed_open_one_cell/xls | 2 | `elapsed_ns.p99` | 126251 | 133890 | +6.050645% | `retain_unresolved` |
| 0554-review-0078 | xls/xls_source_backed_open_one_cell/xls | 2 | `elapsed_ns.max` | 138130 | 153541 | +11.156881% | `retain_unresolved` |
| 0554-review-0079 | xls/xls_source_backed_open_one_cell/xls | 2 | `elapsed_ns.mean` | 108646.96799999996 | 115170.33899999996 | +6.004191% | `retain_unresolved` |
| 0554-review-0080 | xls/xls_source_backed_open_one_cell/xls | 2 | `elapsed_ns.confidence_interval_95.lower` | 108294.20424192103 | 114806.70079354165 | +6.013707% | `retain_unresolved` |
| 0554-review-0081 | xls/xls_source_backed_open_one_cell/xls | 2 | `elapsed_ns.confidence_interval_95.upper` | 108999.7317580789 | 115533.97720645828 | +5.994735% | `retain_unresolved` |
| 0554-review-0082 | cfb/None/None | 1 | `rss.system_seconds` | 0.19 | 0.2 | +5.263158% | `retain_unresolved` |
| 0554-review-0083 | xls/None/None | 1 | `rss.system_seconds` | 0.58 | 0.61 | +5.172414% | `retain_unresolved` |

Each adverse row is retained unresolved because it exceeds the review threshold, including rows that are diagnostic percentiles, spread, confidence bounds, or RSS CPU fields.

### Same-build repeat drift

| ID | Stage | Group/case/shape | Repeats | Metric | First | Second | Change | Disposition |
| --- | --- | --- | --- | --- | ---: | ---: | ---: | --- |
| 0554-review-0084 | baseline | cfb/cfb_open/few-large | 1→2 | `elapsed_ns.max` | 98820 | 148541 | +50.314714% | `retain_diagnostic` |
| 0554-review-0085 | baseline | cfb/cfb_open/few-large | 1→2 | `elapsed_ns.standard_deviation` | 1954.8202091686437 | 3045.036265888566 | +55.770656% | `retain_diagnostic` |
| 0554-review-0086 | baseline | cfb/cfb_open/many-small | 1→2 | `elapsed_ns.max` | 266182 | 176191 | -33.808071% | `retain_diagnostic` |
| 0554-review-0087 | baseline | cfb/cfb_open/many-small | 1→2 | `elapsed_ns.standard_deviation` | 4783.4494710282015 | 3110.8272983614725 | -34.966862% | `retain_diagnostic` |
| 0554-review-0088 | baseline | cfb/cfb_open/tiny | 1→2 | `elapsed_ns.standard_deviation` | 395.6405262128514 | 488.7086596880131 | +23.523408% | `retain_diagnostic` |
| 0554-review-0089 | baseline | xls/xls_eager_open_list_worksheets/xls | 1→2 | `elapsed_ns.standard_deviation` | 12953.767907731766 | 10347.244527690096 | -20.121739% | `retain_diagnostic` |
| 0554-review-0090 | baseline | xls/xls_eager_open_one_cell/xls | 1→2 | `elapsed_ns.standard_deviation` | 9703.673992078773 | 8645.067343713246 | -10.909339% | `retain_diagnostic` |
| 0554-review-0091 | baseline | xls/xls_owned_source_open/xls | 1→2 | `elapsed_ns.max` | 502143 | 124130 | -75.279950% | `retain_diagnostic` |
| 0554-review-0092 | baseline | xls/xls_owned_source_open/xls | 1→2 | `elapsed_ns.standard_deviation` | 14029.786416485582 | 5368.832261306065 | -61.732616% | `retain_diagnostic` |
| 0554-review-0093 | baseline | xls/xls_owned_source_open_list_worksheets/xls | 1→2 | `elapsed_ns.max` | 119921 | 131581 | +9.723068% | `retain_diagnostic` |
| 0554-review-0094 | baseline | xls/xls_owned_source_open_one_cell/xls | 1→2 | `elapsed_ns.max` | 126070 | 483923 | +283.852622% | `retain_diagnostic` |
| 0554-review-0095 | baseline | xls/xls_owned_source_open_one_cell/xls | 1→2 | `elapsed_ns.standard_deviation` | 6070.296482372452 | 13503.520650025679 | +122.452407% | `retain_diagnostic` |
| 0554-review-0096 | baseline | xls/xls_semantic_open/xls | 1→2 | `elapsed_ns.standard_deviation` | 13748.510602734981 | 12896.268185016741 | -6.198798% | `retain_diagnostic` |
| 0554-review-0097 | baseline | xls/xls_source_backed_open/xls | 1→2 | `elapsed_ns.standard_deviation` | 6268.9985736469425 | 5871.685146557095 | -6.337750% | `retain_diagnostic` |
| 0554-review-0098 | baseline | xls/xls_source_backed_open_list_worksheets/xls | 1→2 | `elapsed_ns.max` | 132080 | 145420 | +10.099939% | `retain_diagnostic` |
| 0554-review-0099 | baseline | xls/xls_source_backed_open_list_worksheets/xls | 1→2 | `elapsed_ns.standard_deviation` | 6024.164693306102 | 5657.178894087641 | -6.091895% | `retain_diagnostic` |
| 0554-review-0100 | baseline | xls/xls_source_backed_open_one_cell/xls | 1→2 | `elapsed_ns.max` | 150640 | 138130 | -8.304567% | `retain_diagnostic` |
| 0554-review-0101 | baseline | xls/xls_source_backed_open_one_cell/xls | 1→2 | `elapsed_ns.standard_deviation` | 6147.811015597265 | 5684.723956485291 | -7.532552% | `retain_diagnostic` |
| 0554-review-0102 | baseline | cfb/None/None (native-r1-cfb) | 1→2 | `rss.system_seconds` | 0.19 | 0.2 | +5.263158% | `retain_diagnostic` |
| 0554-review-0103 | baseline | xls/None/None (native-r1-xls) | 1→2 | `rss.system_seconds` | 0.58 | 0.63 | +8.620690% | `retain_diagnostic` |
| 0554-review-0104 | candidate | cfb/cfb_open/few-large | 1→2 | `elapsed_ns.max` | 103620 | 95551 | -7.787107% | `retain_diagnostic` |
| 0554-review-0105 | candidate | cfb/cfb_open/few-large | 1→2 | `elapsed_ns.standard_deviation` | 3259.960794865526 | 2595.307657575179 | -20.388378% | `retain_diagnostic` |
| 0554-review-0106 | candidate | cfb/cfb_open/tiny | 1→2 | `elapsed_ns.max` | 215551 | 13600 | -93.690588% | `retain_diagnostic` |
| 0554-review-0107 | candidate | cfb/cfb_open/tiny | 1→2 | `elapsed_ns.mean` | 2431.780999999999 | 2190.368999999999 | -9.927374% | `retain_diagnostic` |
| 0554-review-0108 | candidate | cfb/cfb_open/tiny | 1→2 | `elapsed_ns.standard_deviation` | 6762.666444271481 | 422.78526490314124 | -93.748246% | `retain_diagnostic` |
| 0554-review-0109 | candidate | xls/xls_eager_open_list_worksheets/xls | 1→2 | `elapsed_ns.max` | 512232 | 458302 | -10.528432% | `retain_diagnostic` |
| 0554-review-0110 | candidate | xls/xls_owned_source_open/xls | 1→2 | `elapsed_ns.p50` | 105060 | 99160 | -5.615839% | `retain_diagnostic` |
| 0554-review-0111 | candidate | xls/xls_owned_source_open/xls | 1→2 | `elapsed_ns.mean` | 105826.46500000004 | 100037.249 | -5.470480% | `retain_diagnostic` |
| 0554-review-0112 | candidate | xls/xls_owned_source_open/xls | 1→2 | `elapsed_ns.standard_deviation` | 5353.572529405242 | 5753.214484782172 | +7.464958% | `retain_diagnostic` |
| 0554-review-0113 | candidate | xls/xls_owned_source_open_list_worksheets/xls | 1→2 | `elapsed_ns.min` | 96410 | 91100 | -5.507727% | `retain_diagnostic` |
| 0554-review-0114 | candidate | xls/xls_owned_source_open_list_worksheets/xls | 1→2 | `elapsed_ns.p50` | 105360 | 99595 | -5.471716% | `retain_diagnostic` |
| 0554-review-0115 | candidate | xls/xls_owned_source_open_list_worksheets/xls | 1→2 | `elapsed_ns.mean` | 106503.93700000015 | 100332.18599999983 | -5.794857% | `retain_diagnostic` |
| 0554-review-0116 | candidate | xls/xls_owned_source_open_list_worksheets/xls | 1→2 | `elapsed_ns.standard_deviation` | 5054.5296147769195 | 5972.382866537379 | +18.159024% | `retain_diagnostic` |
| 0554-review-0117 | candidate | xls/xls_owned_source_open_one_cell/xls | 1→2 | `elapsed_ns.max` | 493212 | 140291 | -71.555639% | `retain_diagnostic` |
| 0554-review-0118 | candidate | xls/xls_owned_source_open_one_cell/xls | 1→2 | `elapsed_ns.standard_deviation` | 13257.151888394677 | 5536.480171816501 | -58.237786% | `retain_diagnostic` |
| 0554-review-0119 | candidate | xls/xls_semantic_open/xls | 1→2 | `elapsed_ns.standard_deviation` | 12384.730909717504 | 10477.149234363262 | -15.402690% | `retain_diagnostic` |
| 0554-review-0120 | candidate | xls/xls_source_backed_open/xls | 1→2 | `elapsed_ns.max` | 148811 | 137541 | -7.573365% | `retain_diagnostic` |
| 0554-review-0121 | candidate | xls/xls_source_backed_open/xls | 1→2 | `elapsed_ns.standard_deviation` | 5653.820620961137 | 5968.241131606636 | +5.561204% | `retain_diagnostic` |
| 0554-review-0122 | candidate | xls/xls_source_backed_open_list_worksheets/xls | 1→2 | `elapsed_ns.min` | 103650 | 98380 | -5.084419% | `retain_diagnostic` |
| 0554-review-0123 | candidate | xls/xls_source_backed_open_list_worksheets/xls | 1→2 | `elapsed_ns.max` | 398172 | 272611 | -31.534362% | `retain_diagnostic` |
| 0554-review-0124 | candidate | xls/xls_source_backed_open_list_worksheets/xls | 1→2 | `elapsed_ns.standard_deviation` | 10161.359772596714 | 7916.599763947477 | -22.091138% | `retain_diagnostic` |
| 0554-review-0125 | candidate | xls/xls_source_backed_open_one_cell/xls | 1→2 | `elapsed_ns.standard_deviation` | 5393.486939764788 | 5859.963719074378 | +8.648890% | `retain_diagnostic` |
| 0554-review-0126 | candidate | cfb/None/None (native-r1-cfb) | 1→2 | `rss.system_seconds` | 0.2 | 0.19 | -5.000000% | `retain_diagnostic` |

Repeat drift is retained as within-stage diagnostic variability. Positive and negative changes are preserved alike; no drift row is used to select a rerun or relax a gate.

## Profile status and disposition

The canonical profile comparison passes all five mechanism booleans across eight matched owner rows: baseline decoder predictions match, candidate standard-root decoder calls are exactly one, many-small owner Ir decreases in both repeats, CLSID call vectors are preserved, and name-decoder/scalar rows remain separate. This is diagnostic-only evidence; it does not establish native latency or adoption.

The canonical instruction comparison also passes all six mechanism booleans across eight matched rows, including mapped decoder/scalar controls and self-instruction attribution. The evidence is diagnostic-only. Final restored-baseline quality passes all 8 commands, covering 4,228 tests in 154 groups, with 0 failures and 27 ignored tests. The metrics disposition remains `reject: one or more frozen main metrics gates failed`; the review therefore sets `adoption_allowed` to `false` and the final source remains the exact restored baseline.

## Evidence bindings

| Artifact | SHA-256 |
| --- | --- |
| `metrics-analysis.json` | `30f4bc0923d5207129d690acf9c2ed59cf216531dc7545470ec08367f4f7d39c` |
| `profile-comparison.json` | `243766cbfd383639cfd70c4276cd81274fe661810b3e3857ab7e17060462e562` |
| `instruction-comparison.json` | `916c19c492027a51300d79f661c49dd1a0b9ff53f476e919371dbe0aa972bbe1` |
| `quality.json` | `09cfa267680b490c9bbaa72bafd37b9c57ec085659a2536f5da570a787ee2668` |
| `analyze_metrics.py` | `3aeff5aba7b6b212bf0b263bf5d43e614ed41d62bfb83b2c53ff3584f8652c82` |
| `plan.json` | `d901c87af4e8d2585af997dbc850832db21db0ce0e359d3dcf7d90c7b415e98c` |
| `run.py` | `f90f8285be651b1115086194e1c3660e88d2b8880bb71f6d4e9fc839075f06d3` |

The pre-final review bytes are preserved at `review-attempts/before-final-quality/` (adverse-review SHA-256 017da19843dd4ccc335f3532f4c1ef2576cd551013d28693aa0dfa9ed9d4f314, results-review SHA-256 e6b5af85eaef175b19056f85d6eeabf3950b8c1aa9bb92e51f61f4ca144c1b7f). The complete original rows, source coverage, gate counts, interpretations, dispositions, and final evidence bindings are retained in [`adverse-review.json`](adverse-review.json).
