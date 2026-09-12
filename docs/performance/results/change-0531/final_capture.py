"""Capture the final binary under the complete frozen native ABBA matrix."""
import json
import subprocess
import analyze
import run

HERE = run.HERE
ORIGINAL_CHECK = run.check_source

def source_check(stage, retained_baseline=False):
    if stage == 'baseline' and retained_baseline:
        ORIGINAL_CHECK('final')
        return 'final'
    return ORIGINAL_CHECK(stage, retained_baseline)

run.check_source = source_check


def jobs():
    plan = run.plan_data()
    final_plan = json.loads((HERE/'final-native-plan.json').read_text())
    assert final_plan['primary_plan_sha256'] == run.sha(HERE/'plan.json')
    for stage, repeat in final_plan['order']:
        for original in analyze._expected_native_jobs(plan):
            if original['repeat'] != repeat:
                continue
            job = dict(original)
            job['name'] = 'final-' + job['name']
            yield stage, job


def command(stage, job):
    return ['taskset', '-c', str(run.plan_data()['cpu']), '/usr/bin/time', '-f',
            '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
            '-o', str(HERE/stage/(job['name']+'.rss.json')),
            str(run.SCRATCH/(stage+'-normal')), '--warmup', str(job['warmup']),
            '--samples', str(job['samples']), '--case', job['case'],
            '--xlsx-cell-crud-shape', job['shape'], '--json', str(HERE/stage/(job['name']+'.json'))]


def capture():
    for stage, job in jobs():
        binary = run.SCRATCH/(stage+'-normal')
        identity = json.loads((HERE/stage/'binary-normal.json').read_text())
        assert run.sha(binary) == identity['sha256']
        observed = subprocess.check_output(['ps','-eo','pid,ppid,etime,comm'],text=True).splitlines()
        run.write(HERE/stage/(job['name']+'.host.json'),dict(observed_utc=run.now(),
            processes=[r for r in observed if r.split()[-1] in ['cargo','rustc']],
            scope='Pre-child observation; not proof of an idle host.'))
        run.run(stage, job['name'], command(stage,job), binary, retained_baseline=stage=='baseline')


if __name__ == '__main__':
    capture()
