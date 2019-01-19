#-*- coding: utf-8 -*-

import re

def isUTF8(data):
    try:
        data.decode('utf8')
    except UnicodeDecodeError:
        return False
    else:
        return True

def isShift_JIS(data):
    try:
        data.decode('shift-jis')
    except UnicodeDecodeError:
        return False
    else:
        return True

def isSequenceMatchPattern(pattern, sequence):
    global config;
    #pattern = u"^\u3010[^\u3011]+\u3011\u30ea\u30b9\u30af";
    if re.match(pattern, sequence, re.UNICODE):
        return True;
    else:
        return False;
