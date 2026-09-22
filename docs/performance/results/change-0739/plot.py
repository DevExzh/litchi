"""Plot per-process lifecycle and phase medians, keeping both corpora separate."""
import json
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
P=Path(__file__).resolve().parent
if __name__=='__main__':
    result=json.loads((P/'analysis.json').read_text())
    fig,axes=plt.subplots(1,2,figsize=(12,4.5))
    phases=['lifecycle_ns','plan_ns','commit_ns','publication_ns','unassigned_ns','reopen_ns']
    labels=['Lifecycle','Plan','Apply','Publish','Residual','Reopen*']
    for ax,group in zip(axes,result['groups']):
        values=[[v/1e6 for v in group['metrics_ns'][p]['p50']['values']] for p in phases]
        ax.boxplot(values,tick_labels=labels,showfliers=False)
        for i,vs in enumerate(values,1):
            ax.scatter([i+(j-4)*.025 for j in range(9)],vs,s=12,alpha=.6)
        ax.set_yscale('log');ax.set_ylabel('Process median duration (ms, log scale)')
        ax.set_title('Media-rich' if 'media_rich' in group['case'] else 'Plain')
        ax.tick_params(axis='x',labelrotation=25);ax.grid(axis='y',alpha=.2)
    fig.suptitle('0739: nine fresh native processes per corpus; *reopen is outside lifecycle')
    fig.tight_layout();fig.savefig(P/'phase-medians.png',dpi=160)
