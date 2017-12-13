# -*- coding: utf-8 -*-
import wx.grid


class Table(wx.grid.GridTableBase):
    def __init__(self):
        wx.grid.GridTableBase.__init__(self)
        # 初期テーブルは3行3列
        self.header = [''] * 3
        self.data = [[''] * 3] * 3
    
    # 表の先頭行と2行目以降のデータを差し替え
    def SetHeaderAndData(self, headerList, dataList):
        self.header = headerList
        self.data = dataList
    
    # 表の先頭行と2行目以降のデータを取得
    def getHeaderAndData(self):
        result = []
        result.append(self.header)
        for i in range(len(self.data)):
            result.append(self.data[i])
        return result
    
    # 割当表から，日付・時限・被割当者のデータを全て取得
    # 重複あり
    def getAssignDays_inAssignTable(self):
        chars_data = self.getChars_fromData()
        indexes = self.getConmaIndexes_fromData(chars_data)
        date_chars = []  # 日付を1文字ずつ格納
        time_chars = []  # 時限を1文字ずつ格納
        name_chars = []  # 1人の名前を1文字ずつ格納
        name_chars_List = []  # 1人ずつ名前を格納
        AssignDay_chars_Set = []  # 日付，時限，名前リストのセット
        AssignedList = []  # 1セットを重ねていく
        
        # 例：11/4, 5限, Andy, Dan
        # ['1', '1', '/', '4']['5', '限']['A', 'n', 'd', 'y']['D', 'a', 'n']
        
        # Assignなら2列目から「日付,時限,名前,名前,・・・」
        #   ┌  ―――――――――――――――――――――――――――> j
        # ｜│['a', 'b', 'c', ',', 'a', 'b'...] ['a', 'b', 'c', ',', 'a'...]
        # ｜│['a', 'b', 'c', ',', 'a', 'b'...] ['a', 'b', 'c', ',', 'a'...]
        # ｜│       …
        # Ｖ│ ―――――――――――――――> k
        # i │['a', 'b', 'c', ',', 'a', 'b'...] ['a', 'b', 'c', ',', 'a'...]
        #   └
        
        # 1行・2列目から開始，日付部分だけ取り出す
        for i in range(len(chars_data)):
            for j in range(1, len(chars_data[i])):
                # カンマ数が2より多い場合
                # ( カンマの位置情報が3つ以上格納されている )
                # ≒日付，時刻，名前のセットがある
                if int(len(indexes[i][j])) > 2:
                    # 日付の部分(最初のカンマまで)===============
                    for k in range(0, indexes[i][j][0]):
                        date_chars.append(chars_data[i][j][k])
                    # 時限の部分(最初のカンマから次のカンマまで)==========
                    for k in range(indexes[i][j][0] + 1, indexes[i][j][1]):
                        time_chars.append(chars_data[i][j][k])
                    # 名前の部分=================================
                    for k in range(indexes[i][j][1] + 1, len(chars_data[i][j])):
                        if chars_data[i][j][k] == ',' \
                                or chars_data[i][j][k] == '，':
                            name_chars_List.append(''.join(name_chars))
                            name_chars = []
                        else:
                            name_chars.append(chars_data[i][j][k])
                    # 最後の人の名前（最後はカンマがない===================
                    name_chars_List.append(''.join(name_chars))
                    name_chars = []
                    # セットをまとめる===========================
                    AssignDay_chars_Set.append(''.join(date_chars))
                    AssignDay_chars_Set.append(''.join(time_chars))
                    AssignDay_chars_Set.append(name_chars_List)
                    date_chars = []
                    time_chars = []
                    name_chars_List = []
                    # 1セットを連ねる============================
                    AssignedList.append(AssignDay_chars_Set)
                    AssignDay_chars_Set = []
        
        # AssignedListの中身
        #   ┌  ――――――――――――――――> j
        #   ｜               ―――――――――> k
        # ｜│['1/1', '2限', ['A', 'B', 'C', ...] ]
        # ｜│['1/2', '1限', ['A', 'B', 'C', ...] ]
        # ｜│       …
        # i │['1/5', '3限', ['A', 'B', 'C', ...] ]
        #   └
        return AssignedList
    
    # 割当表から．被割当者を全て取得する（負担数計算用）
    # 重複あり（A,B,C,B,A,Aのように，1人の割当者に複数の被割当者）
    def getNames_inAssignTable(self):
        chars_data = self.getChars_fromData()
        indexes = self.getConmaIndexes_fromData(chars_data)
        
        tmpChars = []  # 文字群を一時的に格納
        names = []  # 名前を1人ずつ格納
        
        # indexes[i][k][0]:日付後ろカンマ，indexes[i][k][1]:時限後ろカンマ
        for i in range(len(chars_data)):
            for j in range(1, len(chars_data[i])):
                if int(len(indexes[i][j])) > 0:
                    for k in range(indexes[i][j][1] + 1, len(chars_data[i][j])):
                        if chars_data[i][j][k] == ',' \
                                or chars_data[i][j][k] == '，':
                            names.append(''.join(tmpChars))
                            tmpChars = []
                        else:
                            tmpChars.append(chars_data[i][j][k])
                    # セル内最後の人の分
                    if tmpChars != []:
                        names.append(''.join(tmpChars))
                        tmpChars = []
        
        return names
    
    # 被割当表から，被割当者の一覧を取得する（負担数計算用）
    # CSVは名前が重複しないように作る
    def getNames_inRAssignTable(self):
        rowNum = self.GetNumberRows()
        names = []
        for i in range(rowNum):
            names.append(self.data[i][0])
        
        return names
    
    # 1文字ずつのデータに変換して返す
    def getChars_fromData(self):
        # Assignなら2列目から「日付,時限,名前,名前,・・・」
        # RAssignなら列ごとに「名前,名前,・・・」
        rowNum = self.GetNumberRows()
        colNum = self.GetNumberCols()
        chars_col = []
        chars_data = []
        
        # 1字ずつに分解したリストに置き換える
        #   ┌  ――――――――――――――――――――――――> j
        # ｜│  ['abc,abc,...'] ['abc,abc,...'] ['abc,abc,...']
        # ｜│  ['abc,abc,...'] ['abc,abc,...'] ['abc,abc,...']
        # ｜│       …
        # Ｖ│
        # i │  ['abc,abc,...'] ['abc,abc,...'] ['abc,abc,...']
        #   └
        #
        # ↓
        #
        #   ┌  ―――――――――――――――――――――――――――> j
        # ｜│['a', 'b', 'c', ',', 'a', 'b'...] ['a', 'b', 'c', ',', 'a'...]
        # ｜│['a', 'b', 'c', ',', 'a', 'b'...] ['a', 'b', 'c', ',', 'a'...]
        # ｜│       …
        # Ｖ│ ―――――――――――――――> k
        # i │['a', 'b', 'c', ',', 'a', 'b'...] ['a', 'b', 'c', ',', 'a'...]
        #   └
        for i in range(rowNum):
            for j in range(colNum):
                chars_col.append(list(self.data[i][j]))
            chars_data.append(chars_col)
            chars_col = []
        
        return chars_data
    
    # 1文字ずつのデータに対して，カンマの位置情報リストを返す
    def getConmaIndexes_fromData(self, chars_data):
        indexes = []  # 日付・時限部分の終了位置
        colIndexes = []
        unitIndexes = []
        
        # カンマ毎に位置を記録
        #   ┌  ―――――――――――――――――――――――――――> j
        # ｜│  ['a','b','c'...] ['a','b','c'...] ['a','b','c'...]
        # ｜│  ['a','b','c'...] ['a','b','c'...] ['a','b','c'...]
        # ｜│       …
        # Ｖ│   ――――――> k
        # i │  ['a','b','c'...] ['a','b','c'...] ['a','b','c'...]
        #   └
        for i in range(len(chars_data)):
            for j in range(len(chars_data[i])):
                for k in range(len(chars_data[i][j])):
                    if chars_data[i][j][k] == ',' or chars_data[i][j][k] == '，':
                        unitIndexes.append(k)
                colIndexes.append(unitIndexes)
                unitIndexes = []
            indexes.append(colIndexes)
            colIndexes = []
        
        return indexes
    
    def GetColLabelValue(self, col):
        return self.header[col]
    
    def GetRowLabelValue(self, row):
        return str(row + 1)
    
    def GetNumberCols(self):
        return len(self.data[0])
    
    def GetNumberRows(self):
        return len(self.data)
    
    def IsEmptyCell(self, row, col):
        return False
    
    def GetValue(self, row, col):
        return self.data[row][col]
    
    # 読み取り専用にする場合は処理を書かない(pass)
    def SetValue(self, row, col, value):
        self.data[row][col] = value
