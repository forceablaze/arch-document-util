#-*- coding: utf-8 -*-

import re

def isSequenceMatchPattern(pattern, sequence):
    global config;
    #pattern = u"^\u3010[^\u3011]+\u3011\u30ea\u30b9\u30af";
    if re.match(pattern, sequence, re.UNICODE):
        return True;
    else:
        return False;
