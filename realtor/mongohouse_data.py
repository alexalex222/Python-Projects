from urllib2 import Request, urlopen

request = Request('http://mongohouse.io/api/soldrecords')

response_body = urlopen(request).read()
print(response_body)