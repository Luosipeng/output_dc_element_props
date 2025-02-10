"""
    Create the single line diagram using command lines
"""

import etap.api
import json
import etap


from configuration.configuration import base_address


## Step 0: Test the connection
print("Connecting...")
e = etap.api.connect(base_address)
response = e.application.filepaths()
print(response)

response = e.application.pid()
print(response)
# ping
print("Pinging...")
ping_result = e.application.ping()
print(str(ping_result))

response = e.projectdata.setelementprop("UTIL","U10","Bus","B10")
response = e.projectdata.getelementprop("UTIL","U10","Bus")
print(response)