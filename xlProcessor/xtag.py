#coding=utf-8
'''
Created on 2015年7月10日

@author: Anyson
'''

import re

var_tag = '{% (.*?) %}'
var_pattern = re.compile(var_tag,re.S)

class XTag(object):
    '''
    XTag
    '''
    
    @staticmethod
    def parse(source=None, content={}):
        source_temp = source
        items = re.findall(var_pattern, source_temp)
        print items
        
        for var in items:
            s = "{% " + var + " %}"
            print s
            source_temp = re.sub(s, content.get(var, ''), source_temp)
            
        return source_temp