#coding=utf-8
'''
Created on 2015年7月10日

@author: Anyson
'''
"""Definitions for openpyxl shared exception classes."""

class InvalidFileException(Exception):
    """Error for trying to open a non file."""
    
class InvalidTemplateException(Exception):
    """Error for html content template."""
    
class InvalidExcelException(Exception):
    """Error for excel content."""
