import os
import time
import excel
import base64
import winreg
import psutil
import logging
import threading
import win32com
import signal

from socketserver import ThreadingMixIn
from concurrent.futures import ThreadPoolExecutor
from xmlrpc.server import SimpleXMLRPCServer, SimpleXMLRPCRequestHandler
from defusedxml.xmlrpc import monkey_patch
monkey_patch()


USER    = b'symbexcel:c3f409286244438d436935fb0016a0b9'
AUTH    = 'Basic ' + base64.b64encode(USER).decode()
TIMEOUT = 60*10

logging.root.handlers = []
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S',
    handlers=[
        logging.FileHandler("server.log"),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger('server')
logger.setLevel(logging.DEBUG)


class PoolMixIn(ThreadingMixIn):
    pool = ThreadPoolExecutor(max_workers=os.cpu_count())

    def process_request(self, request, client_address):
        x = self.pool.submit(self.process_request_thread, request, client_address)


class SimpleThreadedXMLRPCServer(PoolMixIn, SimpleXMLRPCServer):
    def _dispatch(self, method, params):
        return SimpleXMLRPCServer._dispatch(self, method, params)


class RequestHandler(SimpleXMLRPCRequestHandler):
    rpc_paths = ('/supersecretendpointV3',)

    def do_POST(self):
        if self.headers.get('Authorization') == AUTH:
            SimpleXMLRPCRequestHandler.do_POST(self)

    def do_GET(self):
        pass


def setup_registry_keys():
    PATH = "SOFTWARE\\MICROSOFT\\OFFICE\\16.0\\EXCEL\\SECURITY"
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, PATH)
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, PATH, 0, winreg.KEY_WRITE) as key:
        winreg.SetValueEx(key, "ExtensionHardening", 0, winreg.REG_DWORD, 0)
        winreg.CloseKey(key)

    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, PATH, 0, winreg.KEY_WRITE) as key:
        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)

def excel_processes():
    for process in psutil.process_iter():
        parent = process.parent()
        if process.name() == 'EXCEL.EXE' and (not parent or parent.name() != 'explorer.exe'):
            yield process

def kill_stale_excel():
    while True:
        time.sleep(5)
        for process in excel_processes():
            if time.time() - process.create_time() > TIMEOUT:
                logger.debug('[~] Killing stale Excel: %d' % process.pid)
                process.kill()

def clean_exit():
    for process in excel_processes():
        process.kill()

def run_server(host="0.0.0.0", port=8000):
    setup_registry_keys()

    win32com.client.gencache.EnsureDispatch("Excel.Application").Quit()

    kill = threading.Thread(target=kill_stale_excel, daemon=True)
    kill.start()

    server = SimpleThreadedXMLRPCServer((host, port),
                                        requestHandler=RequestHandler,
                                        allow_none=True)

    server.register_function(excel.start_excel)
    server.register_function(excel.process, 'process')
    server.register_function(excel.get_cell_info, 'get_cell_info')
    server.register_function(excel.get_workbook_info, 'get_workbook_info')
    server.register_function(excel.execute_formula)

    logger.info('Server started: listening on {} port {}'.format(host, port))

    try:
        server.serve_forever()
    except:
        clean_exit()

if __name__ == '__main__':
    run_server()

# Delete win32com temp files
# shutil.rmtree(Path.home().joinpath("AppData\Local\Temp\gen_py"),
#               ignore_errors=True)
