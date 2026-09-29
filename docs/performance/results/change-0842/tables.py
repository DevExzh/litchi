"""Render every metric and explicitly retain all review flags."""
import driver as d
import csv,io
value=d.read(d.P/'analysis.json');stream=io.StringIO();writer=csv.writer(stream)
writer.writerow(['shape','mode','metric','before','after','ratio','low','high'])
lines=['# 0842 complete native results','','Ratios are medians of six matched block ratios. Intervals are paired bootstrap endpoints; RSS is whole-child KiB. No zero-denominator ratio is invented.','','| Shape | Mode | Metric | Before | After | Ratio | 95% interval |','| --- | --- | --- | ---: | ---: | ---: | --- |']
for row in value['summary']:
 p=row['paired'];writer.writerow([row['shape'],row['mode'],row['metric'],row['before_median'],row['after_median'],p['estimate'],p['ci95_low'],p['ci95_high']])
 ratio='undefined' if p['estimate'] is None else f"{p['estimate']:.6f}"
 interval='undefined' if p['estimate'] is None else f"[{p['ci95_low']:.6f}, {p['ci95_high']:.6f}]"
 lines.append(f"| {row['shape']} | {row['mode']} | {row['metric']} | {row['before_median']:.3f} | {row['after_median']:.3f} | {ratio} | {interval} |")
lines += ['',f"p50/RSS review flags: {len(value['regression_flags'])}. Increasing p95/p99 point estimates above 5%: {len(value['tail_review'])}. Block-p50 spread flags: {len(value['spread_flags'])}.",'','Complete block vectors and flag rows are retained in analysis.json.']
for name,text in [('native.csv',stream.getvalue()),('native.md','\n'.join(lines)+'\n')]:
 p=d.P/name
 if p.exists():assert p.read_text()==text
 else:
  with p.open('x') as f:f.write(text)
print('All84native metric rows rendered')
