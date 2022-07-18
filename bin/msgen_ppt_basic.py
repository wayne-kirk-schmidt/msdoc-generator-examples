#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Explanation: A data driven ability to create a powerpoint file
Usage:
    $ msgen_ppt_basic  [ options ]
Style:
    Google Python Style Guide:
    http://google.github.io/styleguide/pyguide.html
    @name           msgen_ppt_basic
    @version        1.0.0
    @author-name    Wayne Schmidt
    @author-email   wschmidt@sumologic.com
    @license-name   APACHE 2.0
    @license-url    http://www.apache.org/licenses/LICENSE-2.0
"""

__version__ = 1.0
__author__ = "Wayne Schmidt (wschmidt@sumologic.com)"

### import os
### import sys

### import collections
### import collections.abc

from pptx import Presentation

prs = Presentation()

prs.save('/tmp/test.pptx')
