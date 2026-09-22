#!/usr/bin/env python3
"""Display all native samples; pooled curves are not independent statistical units."""
import json
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
P=Path(__file__).resolve().parent
manifest=json.loads((P/'captures/manifest.json').read_text())
plt.rcParams.update({'svg.hashsalt':'change-0735','svg.fonttype':'none','font.size':10})
fig,axes=plt.subplots(1,2,figsize=(10,3.6),sharey=True,layout='constrained')
for ax,case,title in zip(axes,['primary','secondary'],['45543.ppt','41246-1.ppt']):
 for variant,color in [('baseline','#2665a8'),('candidate','#c45423')]:
  values=[]
  for row in manifest['runs']:
   if row['lane']=='native' and row['case']==case and row['variant']==variant:
    x=json.loads((P/'captures'/row['output']).read_text());values.extend(s['phase_ns']['whole_ns']/1000 for s in x['samples'])
  assert len(values)==450
  values.sort();ax.step(values,[(i+1)/len(values) for i in range(len(values))],where='post',color=color,label=variant)
 ax.axhline(.5,color='#999999',linestyle=':',linewidth=1);ax.set_title(title);ax.set_xlabel('Public owner elapsed time (µs)');ax.grid(alpha=.15);ax.legend(loc='lower right')
axes[0].set_ylabel('Empirical cumulative fraction')
fig.suptitle('All 450 native samples per line; nine processes per variant and fixture',fontsize=11)
fig.savefig(P/'native-distributions.svg',metadata={'Date':None});fig.savefig(P/'native-distributions.png',dpi=150);plt.close(fig)
(P/'plot-receipt.json').write_text(json.dumps(dict(matplotlib=matplotlib.__version__,status='passed',samples_per_line=450,interpretation='Pooled empirical distributions for visualization only. Process pairs, not pooled samples, are the independent units for reported bootstrap intervals.'),indent=2)+'\n')
print('PASS complete native sample ECDF visualization')
