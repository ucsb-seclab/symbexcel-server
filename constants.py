import win32com.client as win32

VerticalAlignment = {
    win32.constants.xlVAlignTop         : 1,
    win32.constants.xlVAlignCenter      : 2,
    win32.constants.xlVAlignBottom      : 3,
    win32.constants.xlVAlignJustify     : 4,
    win32.constants.xlVAlignDistributed : 5,
}

HorizontalAlignment = {
    win32.constants.xlGeneral : 1,
    win32.constants.xlLeft    : 2,
    win32.constants.xlCenter  : 3,
    win32.constants.xlRight   : 4,
    win32.constants.xlFill    : 5,
    win32.constants.xlJustify : 6,
    win32.constants.xlCenterAcrossSelection : 7,
    win32.constants.xlDistributed : 8,
}


# Keywords is the number 36 in the myOnlineTrainingHub documents,
# but the number 4 in the Excel VBA reference (https://docs.microsoft.com/en-us/office/vba/api/excel.workbook.builtindocumentproperties)
DocumentProperties = {
    36: 4,
}
