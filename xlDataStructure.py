# -*- coding: utf-8 -*-
__author__ = 'yang'

from copy import deepcopy

xlUp = -4162

class xlDict:


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
                elif isinstance(keyCol, (int, long)):
                    keyCol = [keyCol]
                else:
                    raise TypeError("the type of {0} is invalid!\n Integer, Tuple with two elements or List is required.")

            self.level = len(keyCol)

            if not type(valCol) is list:
                if type(valCol) is tuple:
                    valCol = range(valCol[0], valCol[1]+1)
                elif isinstance(valCol, (int, long)):
                    valCol = [valCol]
                else:
                    raise TypeError("the type of {0} is invalid!\n Integer, Tuple with two elements or List is required.")

            self.deep = len(valCol)

            if not endRow:
                endRow = sht.Cells(sht.Rows.Count, valCol[0]).End(xlUp).Row

            # load the value into memory
            keys, vals = [],[]

            if not reversed:
                # load keys
                for i in keyCol:
                    keys.append(list(sht.Range(sht.Cells(startRow, i), sht.Cells(endRow, i)).Value))

                #load values
                for i in valCol:
                    vals.append(list(sht.Range(sht.Cells(startRow, i), sht.Cells(endRow, i)).Value))
            else:
                # load keys
                for i in keyCol:
                    tmp = list(sht.Range(sht.Cells(startRow, i), sht.Cells(endRow, i)).Value)
                    tmp.reverse()
                    keys.append(tmp)

                #load values
                for i in valCol:
                    tmp = list(sht.Range(sht.Cells(startRow, i), sht.Cells(endRow, i)).Value)
                    tmp.reverse()
                    vals.append(tmp)

            vals, keys = [list(sum(i, ())) for i in zip(*vals)], [list(sum(j, ())) for j in keys]

            if self.level == 1:
                self.raw = xlDict.__singleCol(keys[0], vals, 0, len(vals), testKey, ignoreNullVal, setNullValTo)
                if self.deep == 1:
                    xlDict.__reduceOneLevel(self.raw)
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

        for k, v in self.raw.iteritems():
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

        for k, v in tmp.iteritems():
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

        for k, v in self.raw.iteritems():
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
                for k, v in tmp.iteritems():
                    if isinstance(v, dict):
                        xlDict.__sumDict(v, tmp, k)
        return tmp

    def unload(self, sht, keyCol=1, valCol=2, rowStart=1, rowEnd=None):
        if self.level != 1:
            raise AttributeError("keys are ambiguous")
        else:
            if rowEnd is None:
                rowEnd = sht.Cells(sht.Rows.Count, keyCol).End(xlUp).Row

            # load keys
            keyVal = sht.Range(sht.Cells(rowStart, keyCol), sht.Cells(rowEnd, keyCol)).Value
            keyVal = sum(keyVal, ())

            if self.deep > 1:
                val = tuple([self.raw.get(i, (None,)*self.deep) for i in keyVal][:len(keyVal)])
            elif self.deep == 1:
                val = tuple([(self.raw.get(i, None),) for i in keyVal])


            sht.Range(sht.Cells(rowStart, valCol), sht.Cells(rowEnd, valCol + self.deep -1)).Value = val

            return self


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

########################################################
    @staticmethod
    def __simplify(data):
        '''
            to simplify data[x][x][x][z] = y to data[x][z] = y
        '''
        for k, v in data.iteritems():
            if type(v) is dict:
                xlDict.__simplify(v)
                if (k in v) and len(v) == 1:
                    data[k] = v[k]


    @staticmethod
    def __getLevelAndDeep(data, l):
        '''
            the data dict has to be square
        '''
        for k, v in data.iteritems():
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
        for k, v in data.iteritems():
            if not type(v) is dict:
                data[k] = sum(v)
            else:
                xlDict.__reduceOneLevel(v)

    @staticmethod
    def __sumDict(data, upper, tar):
        '''
            data    target dict
            upper   the dict which contains data
            tar     the index of data
        '''
        for k, v in data.iteritems():
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

if __name__ == "__main__":
    import win32com.client as win32
    import os
    application = win32.gencache.EnsureDispatch('Excel.Application')
    wb = application.Workbooks.Open("{0}{1}{2}".format(os.getcwd(), os.path.sep, "data.xlsx"), ReadOnly=False)

    sht = wb.Worksheets("data")
    sht1 = wb.Worksheets("res")
    d = xlDict(sht, 4, (5, 6))

    print(sht1.Range(sht1.Cells(1, 1), sht1.Cells(5, 1)).Value)
    d.unload(sht1)
    print(d.simplify())
    # print(d.reduceToRoot())
    # print(d.reduced())
    # d.reduced().unload(sht1)
    wb.Close(SaveChanges=True)
