# coding=utf-8
# 

import os, configparser



def getCurrentDirectory():
	"""
	Get the absolute path to the directory where this module is in.

	This piece of code comes from:

	http://stackoverflow.com/questions/3430372/how-to-get-full-path-of-current-files-directory-in-python
	"""
	return os.path.dirname(os.path.abspath(__file__))



def _load_config(config_file='cmbhk.config'):
	"""
	Read the config file, convert it to a config object.
	"""
	cfg = configparser.ConfigParser()
	cfg.read(config_file)
	return cfg



# initialized only once when this module is first imported by others
if not 'config' in globals():
	config = _load_config()



def getStartRow():
	"""
	The row where the holding section starts, i.e., the row number of the
	headers.
	"""
	global config
	return int(config['setting']['startrow'])



def getCustodian():
	global config
	return config['setting']['custodian']