import datetime
import hashlib
import logging
import os
import pickle
import tempfile
import traceback
from os.path import basename
from xmlrpc.client import Binary

import pythoncom
import pywintypes
import win32com.client as win32
from joblib import Memory

import sys
pywintypes.datetime = pywintypes.TimeType

CACHE  = os.path.join(tempfile.gettempdir(), "CACHE")
MEMORY = Memory(CACHE, verbose=0)

try:
    os.mkdir(CACHE, 777)
except FileExistsError:
    pass

logger = logging.getLogger('excel')
logger.setLevel(logging.DEBUG)

# Depending on the last argument (nocache) we call the function or use MEMORY.cache.
def cache(func):
    def _cache(*args, **kwargs):
        nocache = args[-1]
        assert(isinstance(nocache, bool))
        if nocache:
            return func(*args, **kwargs)
        else:
            f = MEMORY.cache(func)
            return f(*args, **kwargs)
    return _cache

def start_excel(blob):
    data = blob.data
    sha1 = hashlib.sha1(data).hexdigest()
    path = os.path.join(CACHE, sha1 + '.bin')
    with open(path, 'wb') as f:
        f.write(data)
    return path

@cache
def process(path, nocache):
    logger.debug('Process: %s' % path)
    try:
        return ExcelProcess(path).process()
    except Exception as e:
        logger.error('[Process Exception] - %s: %s' % (basename(path), e))
        raise e

@cache
def get_cell_info(path, sheet_name, col, row, idx, nocache):
    logger.debug('GetCellInfo: %s' % path)
    try:
        return ExcelProcess(path).get_cell_info(sheet_name, col, row, idx)
    except Exception as e:
        logger.error('[GetCellInfo Exception] - %s: %s' % (path, e))
        raise

@cache
def get_workbook_info(path, idx, nocache):
    logger.debug('GetWorkbookInfo: %s' % path)
    try:
        return ExcelProcess(path).get_workbook_info(idx)
    except Exception as e:
        logger.error('[GetWorkbookInfo Exception] - %s: %s' % (path, e))
        raise

def execute_formula(path, sheet_name, col, row, formula, accessed):
    logger.debug('[Execute Formula]: %s %s!$%s$%s' % (path, sheet_name, col, row))

    # update cells
    logger.debug('Updating cells...')
    excel = ExcelProcess(path)
    for cell_sheet_name, cell_column, cell_row, cell_formula, cell_value in accessed['cells'].values():
        sheet = excel.book.Sheets[cell_sheet_name]
        cell  = sheet.Range(f'{cell_column}{cell_row}')
        if cell.Formula != '' and cell.Formula == cell_formula:
            cell.Calculate()
        elif cell.Formula != '':
            logger.error('Mismatching formula during delegation')
            cell.Formula = formula
        else:
            cell.Value = cell_value

        if cell.Value != cell_value:
            logger.error(f'Unexpected unmatching value (expected {cell_value}, got {cell.Value})')
            # todo: we could just overwrite the cell here and ignore the formula

    # update names
    logger.debug('Updating names...')
    for name, value in accessed['names'].items():
        sheet = excel.book.Sheets[sheet_name]
        sheet.Names.Add(Name=name, RefersTo=value)

    try:
        result = excel.execute_formula(sheet_name, col, row, formula)
    except:
        logger.exception('Something went wrong during the formula execution')
        raise RuntimeError('Something went wrong during the formula execution')

    new_accessed = {
        'cells': dict(),
        'names': dict()
    }

    for cell_sheet_name, cell_column, cell_row, cell_formula, cell_value in accessed['cells'].values():
        sheet = excel.book.Sheets[cell_sheet_name]
        cell  = sheet.Range(f'{cell_column}{cell_row}')
        cell_formula = cell.Formula if cell.Formula != '' else None
        new_accessed['cells'][cell.Address] = (cell_sheet_name, cell_column, cell_row, cell_formula, cell.Value)
    for name, name_value in accessed['names'].items():
        sheet = excel.book.Sheets[sheet_name]
        new_accessed['names'][name] = sheet.Names(name).Value

    return result, new_accessed

def get_from_range(name):
    try:
        r = name.RefersToRange
        n = r.Worksheet.Name
        a = r.Address
        return "'%s'!%s" % (n, a)
    except pywintypes.com_error:
        return None

def load_defined_names(excel, book):
    names = {}

    for name in excel.Names:
        value = get_from_range(name)
        if value:
            names[name.Name] = (value, name.RefersToRange.Count)
        else:
            names[name.Name] = (name.RefersTo, False)

    for name, (value, count) in names.items():
        if value.endswith("\x00'"):
            return None

    return names

def specialcells(urange, t):
    try:
        for cell in urange.SpecialCells(t):
            yield cell
    except pywintypes.com_error:
        pass

def convert_date(value, formula=None):
    if isinstance(value, pywintypes.TimeType):
        value = datetime.datetime.strptime(str(value).split('+')[0], '%Y-%m-%d %H:%M:%S')
    return value, formula

def is_protected(urange):
    try:
        locked = urange.Locked
        urange.Locked = False
        urange.Locked = locked
        return False
    except:
        return True

def load_cells(sheet):
    cells   = {}
    urange  = sheet.UsedRange

    if is_protected(urange):
        try:
            return {cell.Address: convert_date(cell.Value, cell.Formula) for cell in urange}
        except pywintypes.com_error:
            return {}

    # To speed up this thing even more, we could combine "close" cells into Ranges
    # and the accessing .Formula and .Value.
    for cell in specialcells(urange, win32.constants.xlCellTypeFormulas):
        cells[cell.Address] = convert_date(cell.Value, cell.FormulaR1C1)

    for cell in specialcells(urange, win32.constants.xlCellTypeConstants):
        cells[cell.Address] = convert_date(cell.Value, None)

    return cells

def load_macrosheets(excel, book):
    macrosheets = {}
    for sheet in book.Excel4MacroSheets:
        macrosheets[sheet.Name] = load_cells(sheet)
    return macrosheets

def load_worksheets(excel, book):
    worksheets = {}
    for sheet in book.Worksheets:
        worksheets[sheet.Name] = load_cells(sheet)
    return worksheets

def load_comments(excel, book):
    comments = {}
    for sheet in list(book.Excel4MacroSheets) + list(book.Worksheets):
        comments[sheet.Name] = {c.Parent.Address: c.Text() for c in sheet.Comments}
    return comments

# https://www.codeproject.com/Articles/640258/Deconstruction-of-a-VBA-Code-Module
def load_vba(excel, book):
    vba = {}

    if not book.HasVBProject:
        return vba

    try:
        project = book.VBProject
    except pywintypes.com_error:
        return None

    if project.Protection == 1:
        return None

    for component in list(project.VBComponents):
        module  = component.CodeModule
        if not module.CountOfLines:
            continue

        index = module.CountOfDeclarationLines + 1

        while index < module.CountOfLines:
            name, kind = module.ProcOfLine(index)
            start  = module.ProcStartLine(name, kind)
            length = module.ProcCountLines(name, kind)
            vba[name] = module.Lines(index, length)
            index  = start + length + 1

    return vba

class ExcelProcess():

    def __init__(self, path):
        self.path  = path
        self.excel = self.open_excel()
        self.book  = self.open_workbook()
        try:
            self.excel.Calculation = -4135 # xlCalculationManual
        except:
            raise NotImplementedError('We are not able to process this file (VBA related problem)')

    def open_excel(self):
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")

        excel.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
        excel.EnableEvents   = False
        excel.DisplayAlerts  = False
        excel.ScreenUpdating = False
        excel.Interactive    = False
        excel.DisplayStatusBar = False
        excel.AlertBeforeOverwriting = False
        excel.EnableCheckFileExtensions = False
        excel.WarnOnFunctionNameConflict = False
        # excel.Visible = 1

        return excel

    def open_corrupted_workbook(self):
        try:
            return self.excel.Workbooks.Open(self.path,
                                             Password='',
                                             CorruptLoad=1) # xlRepairFile
        except:
            traceback.print_exc()
            raise TimeoutError('Workbooks.Open error')

    def open_workbook(self):
        try:
            return self.excel.Workbooks.Open(self.path,
                                             Password='')
        except:
            return self.open_corrupted_workbook()

    def execute_formula(self, sheet_name, col, row, formula):
        logger.debug('Ready to execute formula..')
        sheet = self.book.Sheets[sheet_name]

        # backup current cell
        logger.debug('Backing up current cell')
        current_cell = sheet.Range(f'{col}{row}')
        tmp_current_formula = current_cell.Formula
        tmp_current_value = current_cell.Value

        # rewrite current cell to r1c1_formula
        logger.debug(f'Writing {formula} to current cell')
        current_cell.Formula = f'={formula}'

        # backup next cell
        logger.debug('Backing up next cell')
        next_cell = sheet.Range(f'{col}{row+1}')
        tmp_next_formula = next_cell.Formula
        tmp_next_value = next_cell.Value

        # rewrite next cell to HALT
        logger.debug('Writing HALT to next cell')
        next_cell.Formula = '=HALT()'
        # execute goto
        cell = sheet.Range(f'{col}{row}')
        # cell.Activate()

        trampoline = f'GOTO({sheet_name}!{cell.GetAddressLocal(ReferenceStyle=-4150)})'
        logger.debug(f'Executing trampoline {trampoline}...')

        goto_result = self.excel.ExecuteExcel4Macro(trampoline)
        assert goto_result is True
        # read result
        result = cell.Value

        # restore next cell
        logger.debug('Restoring next cell')
        next_cell.Formula = tmp_next_formula
        # assert next_cell.Value == tmp_value

        # restore current cell
        logger.debug('Restoring current cell')
        current_cell.Formula = tmp_current_formula
        # assert next_cell.Value == tmp_value

        logger.debug('Executed formula: %s -- Result: %s' % (formula, result))
        return result

    def process(self):
        self.book   = self.open_workbook()
        names       = load_defined_names(self.excel, self.book)
        macrosheets = load_macrosheets(self.excel, self.book)
        worksheets  = load_worksheets(self.excel, self.book)
        comments    = load_comments(self.excel, self.book)
        vba         = load_vba(self.excel, self.book)

        # Trick to remove null bytes from defined names.
        if names is None:
            self.path += '.trick'
            try:
                self.book.SaveAs(self.path, win32.constants.xlExcel12) # .xlsb
            except:
                traceback.print_exc()
                raise NotImplementedError('Cant save %s with xlsb format, skipping for now' % self.path)
            self.book.Close()
            book  = self.open_workbook()
            names = load_defined_names(self.excel, book)

        # TODO: add Connections.
        r = pickle.dumps({
            'rpath'      : self.path,
            'names'      : names,
            'macrosheets': macrosheets,
            'worksheets' : worksheets,
            'comments'   : comments,
            'vba'        : vba,
        })

        return r

    # https://berndplumhoff.gitbook.io/sulprobil/excel/excel-vba-solutions/sbgetcell-1
    def get_cell_info(self, sheet_name, col, row, idx):
        sheet = self.book.Sheets[sheet_name]
        cell  = sheet.Range(f'{col}{row}')

        if idx == 2:
            return cell.Row

        if idx == 3:
            return cell.Column

        if idx == 5:
            return cell.Value

        if idx == 7:
            return cell.NumberFormatLocal

        if idx == 8:
            import constants
            return constants.HorizontalAlignment.get(cell.HorizontalAlignment)

        if idx == 17:
            return cell.RowHeight

        if idx == 19:
            return cell.Font.Size

        if idx == 20:
            return cell.Font.Bold

        if idx == 21:
            return cell.Font.Italic

        if idx == 23:
            return cell.Font.Strikethrough

        if idx == 24:
            return cell.Font.ColorIndex

        if idx == 38:
            return cell.Interior.ColorIndex

        if idx == 50:
            import constants
            return constants.VerticalAlignment.get(cell.VerticalAlignment)

        raise NotImplementedError(f'[get_cell_info][%s] Index {idx} not implemented.' % basename(self.path))

    def get_workbook_info(self, idx):
        import constants

        self.book.Activate()
        index = constants.DocumentProperties.get(idx)
        if index is None:
            raise NotImplementedError(f'[get_workbook_info][%s] Index {idx} not implemented.' % basename(self.path))

        return str(self.book.BuiltinDocumentProperties(index))

    def __del__(self):
        try:
            if self.excel:
                self.excel.Quit()
        except:
            pass

        if pythoncom:
            pythoncom.CoUninitialize()


if __name__ == "__main__":
    import sys
    import pprint

    print('[~] Running a test, NOT starting the server..')

    with open(sys.argv[1], 'rb') as f:
        rpath = start_excel(Binary(f.read()))

    result = pickle.loads(process(rpath, True))

    print('Defined Names: ')
    pprint.pprint(result['names'])

    print('Macro Sheets:')
    pprint.pprint(result['macrosheets'])
