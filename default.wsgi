#!/usr/bin/python

activate_this = '/home/vincentni/project/savvytech/venv/bin/activate'
execfile(activate_this, dict(__file__=activate_this))

import sys

sys.path.insert(0, '/home/vincentni/project/savvytech')

from savvytech import app as application