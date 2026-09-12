"""Run the frozen profile matrix with the unused vgdb interface disabled."""
import json
import run as driver
from run import HERE,FOLDER,sha,write
plan=json.loads((HERE/'profile-retry-plan.json').read_text())
assert sha(HERE/'plan.json')==plan['primary_plan_sha256']
assert sha(HERE/'run.py')==plan['frozen_run_sha256']
original=driver.run
def capture(name,command,binary=None,allow_failure=False):
    assert name.startswith('profile-')
    command=list(command)
    command.insert(command.index('valgrind')+1,plan['retry_extra_argument'])
    original(name,command,binary,allow_failure)
    write(FOLDER/(name+'.profile-binding.json'),dict(retry_plan_sha256=sha(HERE/'profile-retry-plan.json'),capture_script_sha256=sha(HERE/'capture_profiles.py'),receipt_sha256=sha(FOLDER/(name+'.receipt.json'))))
driver.run=capture
try:
    driver.capture('profile')
finally:
    driver.run=original
