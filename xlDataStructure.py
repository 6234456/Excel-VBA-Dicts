# -*- coding: utf-8 -*-
__author__ = 'yang'

'''
    titledDict added
'''

from copy import deepcopy
import json

import win32com.client as win32
import os

xlUp = -4162
xltoLeft = -4159

class xlDict:

    excel = win32.gencache.EnsureDispatch('Excel.Application')

    def __init__(self, sht=None, keyCol=1, valCol=2, startRow=1, endRow=None, testKey = lambda x: True, ignoreNullVal = True, setNullValTo = None, reversed = False, data = None):
        """
            load the data on the spreadsheet to dict
            @param
                    sht             Excel.Worksheet Object
                    keyCol          can be a single value           -> the normal Dicts instance
                                    or tuple with two elements      -> (keyColFrom, keyColTo)
                                    or list                         -> [keyLvl1, keyLvl2, ...]
                    valCol          can be a single value           -> the normal Dicts instance
                                    or tuple with two elements      -> (keyColFrom, keyColTo)
                                    or list with more than one      -> [keyLvl1, keyLvl2, ...]
                    startRow        from which row to start
                    endRow          ends at which row
                    testKey         a function returning Bool value, to filter the keys based on the result of which
                    testVal         a function returning Bool value, to filter the value based on the result of which
                    ignoreNullVal   whether to ignore the key if the corresponding value is None, default True
                    setNullValTo    set the None value to the given value, valid only when ignoreNullVal is False
        """

        if data is None:

            if not type(keyCol) is list:
                if type(keyCol) is tuple:
                    keyCol = range(keyCol[0], keyCol[1]+1)
                elif isinstance(keyCol, (int)):
                    keyCol = [keyCol]
                else:
                    raise TypeError("the type of {0} is invalid!\n Integer, Tuple with two elements or List is required.")

            # level is the total level of keys
            self.level = len(keyCol)

            if not type(valCol) is list:
                if type(valCol) is tuple:
                    valCol = range(valCol[0], valCol[1]+1)
                elif isinstance(valCol, (int)):
                    valCol = [valCol]
                else:
                    raise TypeError("the type of {0} is invalid!\n Integer, Tuple with two elements or List is required.")

            # deep is the len of valCol
            self.deep = len(valCol)

            if not endRow:
                endRow = sht.Cells(sht.Rows.Count, valCol[0]).End(xlUp).Row

            # load the value into memory
            keys, vals = [],[]

            if not reversed:
                # load keys
                for i in keyCol:
                    if startRow == endRow:
                        keys.append([sht.Range(sht.Cells(startRow, i), sht.Cells(endRow, i)).Value])
                    else:
                        keys.append(list(sht.Range(sht.Cells(startRow, i), sht.Cells(endRow, i)).Value))

                #load values
                for i in valCol:
                    if startRow == endRow:
                        vals.append([sht.Range(sht.Cells(startRow, i), sht.Cells(endRow, i)).Value])
                    else:
                        vals.append(list(sht.Range(sht.Cells(startRow, i), sht.Cells(endRow, i)).Value))
            else:
                # load keys
                for i in keyCol:
                    if startRow == endRow:
                        tmp = [sht.Range(sht.Cells(startRow, i), sht.Cells(endRow, i)).Value]
                    else:
                        tmp = list(sht.Range(sht.Cells(startRow, i), sht.Cells(endRow, i)).Value)

                    tmp.reverse()
                    keys.append(tmp)

                #load values
                for i in valCol:
                    if startRow == endRow:
                        tmp = [sht.Range(sht.Cells(startRow, i), sht.Cells(endRow, i)).Value]
                    else:
                        tmp = list(sht.Range(sht.Cells(startRow, i), sht.Cells(endRow, i)).Value)

                    tmp.reverse()
                    vals.append(tmp)

            if startRow == endRow:
                vals, keys = [list(i) for i in zip(*vals)], list(keys)
            else:
                vals, keys = [list(sum(i, ())) for i in zip(*vals)], [list(sum(j, ())) for j in keys]

            if self.level == 1:
                self.raw = xlDict.__singleCol(keys[0], vals, 0, len(vals), testKey, ignoreNullVal, setNullValTo)
                # if self.deep == 1:
                #     xlDict.__reduceOneLevel(self.raw)
            else:
                self.raw = xlDict.__setDict(keys, vals, 0, len(vals), testKey, ignoreNullVal, setNullValTo)
        else:
            if isinstance(data, dict):
                self.raw = data
                if data:
                    self.level, self.deep = xlDict.__getLevelAndDeep(data, 0)
                else:
                    self.level, self.deep = 0, 0
            elif isinstance(data, xlDict):
                self.raw, self.level, self.deep = data.raw, data.level, data.deep
            else:
                raise TypeError("dict type required!")

    def __getitem__(self, item):
        return self.raw[item]

    def __setitem__(self, key, value):
         self.raw[key] = value

    def __str__(self):
        return str(self.raw)

    def __and__(self, other):
        res, tmp = dict(), dict()
        if isinstance(other, dict):
            tmp = other
        elif isinstance(other, xlDict):
            tmp = other.raw
        else:
            raise TypeError("dict or xlDict expected!")

        for k, v in self.raw.items():
            if k in tmp:
                res[k] = v

        return xlDict(data=res)

    def __or__(self, other):
        res, tmp = deepcopy(self.raw), dict()
        if isinstance(other, dict):
            tmp = other
        elif isinstance(other, xlDict):
            tmp = other.raw
        else:
            raise TypeError("dict or xlDict expected!")

        for k, v in tmp.items():
            if k not in self.raw:
                res[k] = v
        return xlDict(data=res)

    def __sub__(self, other):
        res, tmp = dict(), dict()
        if isinstance(other, dict):
            tmp = other
        elif isinstance(other, xlDict):
            tmp = other.raw
        else:
            raise TypeError("dict or xlDict expected!")

        for k, v in self.raw.items():
            if k not in tmp:
                res[k] = v
        return xlDict(data=res)

    def __xor__(self, other):
        return (self | other) - (self & other)

    def update(self, other):
        if isinstance(other, dict):
            self.raw.update(other)
        elif isinstance(other, xlDict):
            self.raw.update(other.raw)
        else:
            raise TypeError("dict or xlDict expected!")
        return self

    def reduced(self):
        tmp = deepcopy(self.raw)
        xlDict.__reduceOneLevel(tmp)
        return xlDict(data=tmp)

    def reduceToRoot(self):
        tmp = self.simplify()

        if self.level > 1:
            xlDict.__reduceOneLevel(tmp)
            for i in range(self.level):
                for k, v in tmp.items():
                    if isinstance(v, dict):
                        xlDict.__sumDict(v, tmp, k)
        return xlDict(data=tmp)

    def unload(self, sht, keyCol=1, valCol=2, rowStart=1, rowEnd=None):
        if self.level != 1:
            raise AttributeError("keys are ambiguous")
        else:
            if rowEnd is None:
                rowEnd = sht.Cells(sht.Rows.Count, keyCol).End(xlUp).Row

            # load keys
            keyVal = sht.Range(sht.Cells(rowStart, keyCol), sht.Cells(rowEnd, keyCol)).Value
            if keyVal is None:
                raise AttributeError("The Key Column is empty!")

            keyVal = sum(keyVal, ())

            if self.deep > 1:
                val = tuple([self.raw.get(i, (None,)*self.deep) for i in keyVal][:len(keyVal)])
            elif self.deep == 1:
                #val = tuple([(self.raw.get(i, None),) for i in keyVal])
                val = list()
                for i in keyVal:
                    tmp = self.raw.get(i, None)
                    if tmp is not None and isinstance(tmp, (list, tuple)):
                        tmp = tmp[0]
                    val.append((tmp,))
                val = tuple(val)

            sht.Range(sht.Cells(rowStart, valCol), sht.Cells(rowEnd, valCol + self.deep -1)).Value = val

            return self

    def dumpKeys(self, sht, startRow=1, startCol=1):
        for k in self.raw.keys():
            sht.Cells(startRow, startCol).Value = k
            startRow += 1


    def dump(self, sht, startRow=1, startCol=1, titled=False):
        # dump the keys
        cnt, karr = 1, list()

        if titled:
            startRow = startRow + 1

        startRow_ = startRow

        for k in self.raw.keys():
            if titled and cnt == 1:
                for i in self.raw[k]:
                    karr.append(i.keys()[0])
                cnt = 0

            sht.Cells(startRow, startCol).Value = k
            startRow += 1


        self.unload(sht, startCol, startCol+1, startRow_)


        if titled:
            sht.Range(sht.Cells(startRow, startCol+1), sht.Cells(startRow, startCol + len(karr))).Value = tuple(karr)

    def append(self, sht=None, keyCol=1, valCol=2, startRow=1, endRow=None, testKey = lambda x: True, ignoreNullVal = True, setNullValTo = None, reversed = False, data = None):
        other = xlDict(sht,keyCol,valCol,startRow,endRow,testKey,ignoreNullVal,setNullValTo,reversed,data)
        return self.update(other)

    def simplify(self):
        if self.level <= 1:
            return self
        else:
            tmp = deepcopy(self.raw)
            xlDict.__simplify(tmp)
            return tmp

    def map(self, keyFun=lambda k: k, valFun=lambda v: v):
        res = dict()
        for k, v in self.raw.items():
            res[keyFun(k)] = valFun(v)
        self.raw = res

        return self

    def toJSON(self):
        return json.dumps(self.raw, indent=4, sort_keys=True)

    def titledDict(self, title):
        '''
            @param  title list or tuple which contains the title corresponding to the values in the val list
                    default for more than one val
        '''
        return self.map(valFun=lambda v: dict(zip(title, v)))

########################################################
    @staticmethod
    def __simplify(data):
        '''
            to simplify data[x][x][x][z] = y to data[x][z] = y
        '''
        for k, v in data.items():
            if type(v) is dict:
                xlDict.__simplify(v)
                if (k in v) and len(v) == 1:
                    data[k] = v[k]


    @staticmethod
    def __getLevelAndDeep(data, l):
        '''
            the data dict has to be square
        '''
        for k, v in data.items():
            if not type(v) is dict:
                if isinstance(v, (str, int, float)):
                    return (l + 1, 1)

                if isinstance(v, (tuple, list)):
                    return (l + 1, len(v))
            else:
                xlDict.__getLevelAndDeep(v, l + 1)


    @staticmethod
    def __singleCol(keyCol, valCols, _start, _end, testKey, ignoreNullVal, setNullValTo):
        data = dict()

        for i in range(_start, _end):
            tmpV = valCols[i]
            j = keyCol[i]
            if j and testKey(j):

                if any(tmpV):
                    # not null
                    data[j] = tmpV
                elif not ignoreNullVal:
                    # if null and should not be ignored
                    # and exists default value
                    if setNullValTo:
                        data[j] = setNullValTo
                    else:
                        data[j] = tmpV

        return data


    @staticmethod
    def __setDict(keyCols, valCols, _start, _end, testKey, ignoreNullVal, setNullValTo, name=None):
        res, lastKey, tmp_start, cnt = dict(), name, _start, False

        if len(keyCols) == 2:
            for i in range(_start, _end):
                e = keyCols[0][i]
                if e:
                    if cnt:
                        # if there exists valid key
                        res[lastKey] = xlDict.__singleCol(keyCols[-1], valCols, tmp_start, i, testKey, ignoreNullVal, setNullValTo)

                    tmp_start = i
                    lastKey = e
                    cnt = True

            res[lastKey] = xlDict.__singleCol(keyCols[-1], valCols, tmp_start, _end, testKey, ignoreNullVal, setNullValTo)
        elif len(keyCols) > 2:
            for i in range(_start, _end):
                e = keyCols[0][i]
                if e:
                    if cnt:
                        res[lastKey] = xlDict.__setDict(keyCols[1:], valCols, tmp_start, i, testKey, ignoreNullVal, setNullValTo, lastKey)

                    tmp_start = i
                    lastKey = e
                    cnt = True

            res[lastKey] = xlDict.__setDict(keyCols[1:], valCols, tmp_start, _end, testKey, ignoreNullVal, setNullValTo, lastKey)
        return res

    @staticmethod
    def reduceDict(data):
        if isinstance(data, xlDict):
            return xlDict.__reduceOneLevel(data.raw)
        elif isinstance(data, dict):
            return xlDict.__reduceOneLevel(data)
        else:
            raise TypeError("dict or xlDict expected!")

    @staticmethod
    def __reduceOneLevel(data):
        for k, v in data.items():
            if not type(v) is dict:
                data[k] = sum([i for i in v if i is not None])
            else:
                xlDict.__reduceOneLevel(v)

    @staticmethod
    def __sumDict(data, upper, tar):
        '''
            data    target dict
            upper   the dict which contains data
            tar     the index of data
        '''
        for k, v in data.items():
            if not type(v) is dict:
                upper[tar] = sum(data.values())
                break
            else:
                xlDict.__sumDict(v, data, k)


    @staticmethod
    def __reduceToRoot(data):
        xlDict.__reduceOneLevel(data)
        xlDict.__sumDict(data)
        return data

    @staticmethod
    def getWorkBook(wbName, readOnly=True):
        return xlDict.excel.Workbooks.Open("{0}{1}{2}".format(os.getcwd(), os.path.sep, wbName), ReadOnly=readOnly)

    @staticmethod
    def createWorkBook(wbName):
        wb = xlDict.excel.Workbooks.Add()
        wb.SaveAs("{0}{1}{2}".format(os.getcwd(), os.path.sep, wbName))
        wb.Close(SaveChanges=False)
        return xlDict.excel.Workbooks.Open("{0}{1}{2}".format(os.getcwd(), os.path.sep, wbName), ReadOnly=False)



if __name__ == "__main__":
    application = win32.gencache.EnsureDispatch('Excel.Application')
    wb = application.Workbooks.Open("{0}{1}{2}".format(os.getcwd(), os.path.sep, "info.xlsx"), ReadOnly=True)

    sht = wb.Worksheets("Entity")
    # d = xlDict(sht, [1, 2], (4, 89))

    # print(d.simplify())
    # print(d.reduceToRoot())
    # print(sht.Range(sht.Cells(1,1), sht.Cells(1,4)).Value)

    # print(d.reduced())
    # d.reduceToRoot().unload(sht1)
    wb.Close(SaveChanges=True)
