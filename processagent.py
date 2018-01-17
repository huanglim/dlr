
from var_settings import INVALID_VALUES

class ProcessAgent(object):
	"""docstring for ProcessRequest"""
	def __init__(self):
		pass
		
	def process_request(driver, function, value=''):
		try:
			getattr(driver, function)(value)
		except Exception as e:
			raise
			
	def makeup_functions(self, request):
		if not isinstance(request, Mapping):
			logging.error('the request type should be dict')
			raise AttributionError('the request type should be dict')

		functions['logon_check'] = ''
		
		for key, val in request.items():
			if str(val).lower().strip() not in INVALID_VALUES and val is not None:
				functions[function_mapping[key]] = val 

		functions['run_report'] = 5
		functions['export_report'] = ''

		return functions