
## symbexcel COM server

This component allows symbexcel to interact with Office [VBA](https://docs.microsoft.com/en-us/office/vba/api/overview/excel).
The [symbexcel server](server.py) implements a multithreaded XML-RPC server, and it exposes serveral functions which are used by the [symbexcel client](https://github.com/ucsb-seclab/symbexcel/blob/main/symbexcel/excel_wrapper/com_wrapper.py) to process Excel files.


## Quick Start

1. Install Windows 10 and Office 2019 in a virtual machine

2. Install Python 3.9.6 from the Microsoft Store

3. Install Windows Terminal from the Microsoft Store (recommended)

4. Clone this repository inside the VM
```bash
    git clone git@github.com:ucsb-seclab/symbexcel-server.git
```

5. Install dependencies
```bash
    cd symbexcel-server
    pip install -e requirements.txt
```

6. Run a test
```bash
    Z:\> python excel.py .\tests\test.xlsm
    [~] Running a test, NOT starting the server..
    Defined Names:
    {'TEST_NAME': ('=Macro1!$A$3', 1)}
    Macro Sheets:
    {'Macro1': {'$A$1': (False, '=ALERT("FORMULA1")'),
                '$A$2': (2.0, '=SUM(1, 1)'),
                '$A$3': (False, '=ALERT(TEST_NAME)'),
                '$B$1': (10.0, None)}}
```

7. Run the server:
```bash
    Z:\> python server.py
    14:53:18 [INFO] Server started: listening on 0.0.0.0 port 8000
```
