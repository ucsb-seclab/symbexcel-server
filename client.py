import xmlrpc.client
import pickle
import sys
import os
import time

USER = 'symbexcel'
PASS = 'c3f409286244438d436935fb0016a0b9'

s = xmlrpc.client.ServerProxy('http://%s:%s@127.0.0.1:8000/supersecretendpointV3' % (USER, PASS))
blob   = open(sys.argv[1], 'rb').read()
start = time.time()
print('Submitted: %s' % os.path.basename(sys.argv[1]))
try:
    result = s.process(blob).data
    result = pickle.loads(result)
except Exception as e:
    print('[%s] Exception' % os.path.basename(sys.argv[1]))
    print(repr(e))
    sys.exit()

# assert(result['names']['auto_open'] == '=Macro1!$A$1')
took = time.time() - start

cells = 0
# print('MACROSHEETS')
for name, sheet in result['macrosheets'].items():
    # print(name, len(sheet))
    for coordinate, content in sheet.items():
        if any(content):
            cells += 1
            # print(coordinate, content)

# print('WORKSHEETS')
for name, sheet in result['worksheets'].items():
    # print(name, len(sheet))
    for coordinate, (value, formula) in sheet.items():
        if any(content):
            cells += 1
            print(coordinate, value, formula)

print(result['names'])
print("[%s] Took: %s (%d)" % (os.path.basename(sys.argv[1]), took, cells))
