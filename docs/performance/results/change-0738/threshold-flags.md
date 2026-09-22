# All absolute changes exceeding 5%

Zero-based paired process rounds; no excluded observations. Full-window p99 and maximum coincide at 50 samples. Window entries are separately reported p50 diagnostics.

## Sample restoration

- primary archive → prior, windows/first10: round 0: +11.528532%, round 1: +10.932459%, round 2: +10.104646%, round 3: +9.864245%, round 4: +9.824904%, round 5: +8.479045%, round 6: +9.591353%, round 7: +9.483403%, round 8: +10.135744%.
- primary prior → restored-a, windows/first10: round 1: -5.185604%, round 4: -5.164426%.
- primary prior → restored-a, windows/last10: round 1: -7.685926%, round 2: -6.624397%, round 3: -6.063051%, round 4: -7.448635%, round 5: -6.606724%, round 6: -6.770933%, round 7: -7.052603%, round 8: -7.339831%.
- secondary archive → prior, metrics/p50: round 0: +7.591426%, round 1: +7.104619%, round 2: +5.111969%, round 3: +7.604189%, round 4: +7.168251%, round 5: +6.682646%, round 6: +5.191958%, round 7: +7.444485%, round 8: +6.774604%.
- secondary archive → prior, windows/first10: round 0: +6.812181%, round 1: +6.124968%, round 2: +6.931015%, round 3: +6.701487%, round 4: +6.553885%, round 5: +6.613315%, round 6: +6.692287%, round 7: +6.903557%, round 8: +6.363957%.
- secondary archive → prior, windows/middle30: round 3: +5.338293%, round 4: +5.176053%, round 7: +5.889007%.
- secondary prior → restored-a, metrics/p99: round 0: +16.304677%.
- secondary prior → restored-a, metrics/maximum: round 0: +16.304677%.
- secondary restored-a → restored-b, metrics/p99: round 0: -12.014596%, round 1: +7.869361%, round 7: +15.525494%.
- secondary restored-a → restored-b, metrics/maximum: round 0: -12.014596%, round 1: +7.869361%, round 7: +15.525494%.

## Startup arguments

- primary archive-standard → archive-extra, metrics/p99: round 0: -5.868623%.
- primary archive-standard → archive-extra, metrics/maximum: round 0: -5.868623%.
- primary prior-default → prior-explicit, metrics/p50: round 0: +7.043308%, round 1: +5.318609%, round 2: +7.562055%, round 3: +6.654117%, round 4: +6.908829%, round 5: +6.955973%, round 6: +8.100117%, round 7: +8.148024%, round 8: +7.712367%.
- primary prior-default → prior-explicit, metrics/mean: round 0: +5.109123%, round 4: +5.170107%, round 7: +5.593283%, round 8: +5.073829%.
- primary archive-standard → prior-default, metrics/p50: round 0: -6.252852%, round 2: -6.136715%, round 3: -6.359502%, round 4: -6.061867%, round 5: -6.306044%, round 6: -5.850337%, round 7: -6.073091%, round 8: -6.422274%.
- primary archive-standard → prior-default, metrics/p99: round 0: -5.633359%.
- primary archive-standard → prior-default, metrics/maximum: round 0: -5.633359%.
- secondary archive-standard → archive-extra, metrics/p50: round 0: +6.130729%, round 1: +6.179647%, round 2: +6.292699%, round 3: +6.847078%, round 4: +6.095277%, round 5: +5.658697%, round 6: +6.422712%, round 7: +6.142069%, round 8: +5.068060%.
- secondary prior-default → prior-explicit, metrics/p50: round 0: +7.781450%, round 1: +8.793295%, round 2: +8.211135%, round 3: +6.748772%, round 4: +7.510368%, round 5: +7.818672%, round 6: +7.269728%, round 7: +7.521904%, round 8: +7.997665%.
- secondary prior-default → prior-explicit, metrics/p99: round 2: +13.807172%.
- secondary prior-default → prior-explicit, metrics/maximum: round 2: +13.807172%.
- secondary archive-extra → prior-explicit, metrics/p99: round 2: +13.951961%.
- secondary archive-extra → prior-explicit, metrics/maximum: round 2: +13.951961%.
