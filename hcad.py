import requests

def login():
	req = requests.get("https://public.hcad.org/records/Real.asp", verify = False)
	return dict(req.cookies)