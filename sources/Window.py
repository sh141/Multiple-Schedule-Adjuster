# -*- coding: utf-8 -*-
import csv
import os.path
import wx
import wx.grid
import ScheduleTable


class GUI(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, parent=None)
        self.appTitle = '†割り当てし者† (Multiple Schedule Adjuster)'
        self.appForeColor = '#000000'
        self.appBackColor = '#DDDDDD'
        self.buttonColor = '#CCFFCC'
        
        self.AssignGrid_MaxSize = (1000, 200)
        self.RAssignGrid_MaxSize = (300, 600)
        self.AssignedDaysGrid_MaxSize = (700, 500)
        # 割り当てをする側
        self.mode_Assign = 'Assign'
        self.schedule_Assign = ScheduleTable.Table()
        # 割り当てを受ける側
        self.mode_RAssign = 'RAssign'
        self.schedule_RAssign = ScheduleTable.Table()
        # 割り当て日程の一覧
        self.mode_AssignedDays = 'AssignedDays'
        self.assignedDays = ScheduleTable.Table()
        
        # ================================================
        self.SetTitle(self.appTitle)
        self.SetMinSize((300, 200))
        self.winFont = wx.Font(14, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL,
                               wx.FONTWEIGHT_NORMAL)
        self.SetFont(self.winFont)
        self.SetForegroundColour(self.appForeColor)
        self.SetBackgroundColour(self.appBackColor)
        
        self.initGrid()
        self.initUI()  # UI設定
        self.Show()
    
    # CSV読み込み
    def readCSV(self, csvPath: str, mode: str):
        data = []
        # 要指定：CSVの文字コード
        with open(csvPath, 'r', encoding='shift-jis') as f:
            reader = csv.reader(f)
            header = next(reader)  # 先頭行のデータ
            for row in reader:
                data.append(row)  # 2行目から全てのデータ
        
        # データ表示========================================
        if mode == self.mode_Assign:
            # 表データ差し替え
            self.schedule_Assign.SetHeaderAndData(header, data)
            # ファイル名の表示===========================
            dir = csvPath.split('\\')
            self.assignLabel.SetLabel(dir[len(dir) - 1])
        
        if mode == self.mode_RAssign:
            # 表データ差し替え
            self.schedule_RAssign.SetHeaderAndData(header, data)
            # ファイル名の表示===========================
            dir = csvPath.split('\\')
            self.rAssignLabel.SetLabel(dir[len(dir) - 1])
        
        # テーブル差し替え＋表示更新＋サイズ調整＋ウィンドウサイズ調整
        self.reloadTables_resizeWindow()
    
    # CSV保存
    def saveCSV(self, csvPath: str, mode: str):
        # ファイルが既にあれば上書き，無ければ新規作成
        if os.path.exists(csvPath):
            state = 'w'
        else:
            state = 'a'
        
        with open(csvPath, state, newline='') as f:
            writer = csv.writer(f)
            if mode == self.mode_Assign:
                data = self.schedule_Assign.getHeaderAndData()
                for i in range(len(data)):
                    writer.writerow(data[i])
            
            if mode == self.mode_RAssign:
                data = self.schedule_RAssign.getHeaderAndData()
                for i in range(len(data)):
                    writer.writerow(data[i])
            
            if mode == self.mode_AssignedDays:
                data = self.assignedDays.getHeaderAndData()
                for i in range(len(data)):
                    writer.writerow(data[i])
    
    # CSV読み込み時のダイアログ
    def openDialog(self, mode: str):
        filter = "csv file(*.csv;*.txt) | *.csv;*.txt | All file(*.*) | *.*"
        # ファイル選択ダイアログ
        dialog = wx.FileDialog(None, 'CSV読込み')
        dialog.SetWildcard(filter)  # 標準で選択するファイル形式
        status = dialog.ShowModal()  # ダイアログを表示
        # キャンセルでない場合はファイルパスからCSV読み込み
        if status != wx.ID_CANCEL:
            filePath = dialog.GetPath()
            self.readCSV(filePath, mode)
    
    # CSV保存時のダイアログ
    def saveDialog(self, mode: str):
        filter = "csv file(*.csv;*.txt) | *.csv;*.txt | All file(*.*) | *.*"
        # ファイル保存＆上書き時に確認するダイアログ
        dialog = wx.FileDialog(None, 'CSV保存',
                               style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        dialog.SetWildcard(filter)  # 標準で選択するファイル形式
        status = dialog.ShowModal()  # ダイアログを表示
        # キャンセルでない場合はファイルパスにCSV保存
        if status != wx.ID_CANCEL:
            filePash = dialog.GetPath()
            self.saveCSV(filePash, mode)
    
    # 割当の負担数計算，負担表を更新
    def createAssigns(self):
        # 割当者に割り当てられる全ての被割当者
        AssignNameData = self.schedule_Assign.getNames_inAssignTable()
        # 被割当者の一覧
        RAssignNameData = self.schedule_RAssign.getNames_inRAssignTable()
        AssignNums = []  # 被割当者ごとの負担数
        
        # 被割当者ごとに，何人に割り当てを受けているか計算
        for i in range(len(RAssignNameData)):
            assignCount = 0
            for j in range(len(AssignNameData)):
                if RAssignNameData[i] == AssignNameData[j]:
                    assignCount += 1
            AssignNums.append(assignCount)
        
        # 被割当表の割当数を書き換え
        rows = self.schedule_RAssign.GetNumberRows()
        for i in range(rows):
            self.schedule_RAssign.SetValue(i, 1, AssignNums[i])
        
        # テーブル差し替え＋表示更新＋サイズ調整＋ウィンドウサイズ調整
        self.reloadTables_resizeWindow()
    
    # 割当日程の重複矛盾を判定，割当日程表を更新
    def checkDuplicateAssigns(self):
        AssignDaysList = self.schedule_Assign.getAssignDays_inAssignTable()
        DuplicateNames = []  # 重複した名前の一覧を格納
        OrganizedAssignDays = [['日付', '時間', ['名前'], ['重複者']]]
        
        for i in range(len(AssignDaysList)):
            time_duplicate = False
            for x in reversed(range(len(OrganizedAssignDays))):
                # 重複する日付があればそこで時間重複を確認
                if AssignDaysList[i][0] == OrganizedAssignDays[x][0]:
                    # 重複する時間があればそこで名前重複を確認
                    if AssignDaysList[i][1] == OrganizedAssignDays[x][1]:
                        for k in range(len(AssignDaysList[i][2])):
                            name_duplicate = False
                            for z in reversed(
                                    range(len(OrganizedAssignDays[x][2]))):
                                # 重複する名前があれば重複リストに追記
                                if AssignDaysList[i][2][k] == \
                                        OrganizedAssignDays[x][2][z]:
                                    OrganizedAssignDays[x][3].append(
                                        AssignDaysList[i][2][k]
                                    )
                                    name_duplicate = True  # 名前重複ありフラグ
                            # 重複する名前が無ければ名前を登録
                            if name_duplicate == False:
                                OrganizedAssignDays[x][2].append(
                                    AssignDaysList[i][2][k]
                                )
                        time_duplicate = True  # 時間重複ありフラグ
            # 重複する日付も時間も無ければ日付，時間，名前を全て登録
            if time_duplicate == False:
                OrganizedAssignDays.append([
                    AssignDaysList[i][0],
                    AssignDaysList[i][1],
                    AssignDaysList[i][2],
                    []  # 重複者がいないので空欄
                ])
        
        # 割当日程表から名前が重複している行を記録
        duplicateRowsList = []
        for i in range(len(OrganizedAssignDays)):
            findFlag = False
            for k in range(len(OrganizedAssignDays[i][2])):
                for x in range(len(DuplicateNames)):
                    if OrganizedAssignDays[i][2][k] == DuplicateNames[x]:
                        duplicateRowsList.append(i)
                        findFlag = True
                        break
                if findFlag == True:
                    break
        
        # 割当日程リストの被割当者リストと重複者リストを1つの文字列に変換
        for i in range(len(OrganizedAssignDays)):
            OrganizedAssignDays[i][2] = ','.join(OrganizedAssignDays[i][2])
            OrganizedAssignDays[i][3] = ','.join(OrganizedAssignDays[i][3])
        
        # 割当日程表の書き込み
        header = OrganizedAssignDays[0]
        del OrganizedAssignDays[0]
        self.assignedDays.SetHeaderAndData(header, OrganizedAssignDays)
        
        # 割当日程表の着色をリセット
        maxRows = self.assignedDays.GetNumberRows()  # 割当日程表の行数
        maxCols = self.assignedDays.GetNumberCols()  # 割当日程表の列数
        for i in range(maxRows):
            for p in range(maxCols):
                self.grid_assignedDays.SetCellBackgroundColour(
                    i - 1, p, '#FFFFFF')
        
        # 割当日程表で重複者がいる行を着色
        for i in range(maxRows):
            if len(OrganizedAssignDays[i - 1][3]) > 0:
                for p in range(maxCols):
                    self.grid_assignedDays.SetCellBackgroundColour(
                        i - 1, p, '#FF9999')
        
        self.reloadTables_resizeWindow()
    
    # テーブル・ウィンドウ更新
    def reloadTables_resizeWindow(self):
        prevSize = self.GetSize()  # 今のウィンドウサイズ
        
        # テーブル再設置
        self.grid_Assign.SetTable(self.schedule_Assign)
        self.grid_Assign.Refresh()  # リロード
        # レイアウト再調整===========================
        self.assignLayout.Fit(self)
        self.grid_Assign.AutoSize()
        
        # テーブル再設置
        self.grid_RAssign.SetTable(self.schedule_RAssign)
        self.grid_RAssign.Refresh()  # リロード
        # レイアウト再調整===========================
        self.rAssignLayout.Fit(self)
        self.grid_RAssign.AutoSize()
        
        # テーブル再設置
        self.grid_assignedDays.SetTable(self.assignedDays)
        self.grid_assignedDays.Refresh()  # リロード
        # レイアウト再調整===========================
        self.assignDaysLayout.Fit(self)
        self.grid_assignedDays.AutoSize()
        
        self.mainLayout.Fit(self)  # ウィンドウ再調整
        
        self.SetSize(prevSize)  # ウィンドウサイズを維持
    
    # ボタンの処理==================================
    def pressAssignButton(self, evt):
        self.createAssigns()
        self.checkDuplicateAssigns()
    
    def pressOpenButton_Assign(self, evt):
        self.openDialog(self.mode_Assign)
    
    def pressOpenButton_RAssign(self, evt):
        self.openDialog(self.mode_RAssign)
    
    def pressSaveButton_Assign(self, evt):
        self.saveDialog(self.mode_Assign)
    
    def pressSaveButton_RAssign(self, evt):
        self.saveDialog(self.mode_RAssign)
    
    def pressSaveButton_AssignedDays(self, evt):
        self.saveDialog(self.mode_AssignedDays)
    
    # ==============================================
    # テーブル初期化
    def initGrid(self):
        self.grid_Assign = wx.grid.Grid(self)
        self.grid_Assign.CreateGrid(5, 5)
        self.grid_Assign.SetTable(self.schedule_Assign)
        self.grid_Assign.SetMaxSize(self.AssignGrid_MaxSize)
        self.grid_Assign.AutoSize()
        
        self.grid_RAssign = wx.grid.Grid(self)
        self.grid_RAssign.CreateGrid(7, 2)
        self.grid_RAssign.SetTable(self.schedule_RAssign)
        self.grid_RAssign.SetMaxSize(self.RAssignGrid_MaxSize)
        self.grid_RAssign.AutoSize()
        
        self.grid_assignedDays = wx.grid.Grid(self)
        self.grid_assignedDays.CreateGrid(5, 3)
        self.grid_assignedDays.SetTable(self.assignedDays)
        self.grid_assignedDays.SetMaxSize(self.AssignedDaysGrid_MaxSize)
        self.grid_assignedDays.AutoSize()
    
    # UI作成
    def initUI(self):
        #
        # 割り当て部レイアウト===============================
        #
        
        # 表のデータ取得ボタン
        AssignButton = wx.Button(self, -1, '割当実行')
        AssignButton.SetBackgroundColour(self.buttonColor)
        AssignButton.Bind(wx.EVT_BUTTON, self.pressAssignButton)
        
        notice_Assign = wx.StaticText(self, -1, '■割当表')
        # ファイル名ラベル
        self.assignLabel = wx.StaticText(self, -1, 'none')
        # ファイル読込みボタン
        openButton_Assign = wx.Button(self, -1, 'CSV読込')
        openButton_Assign.SetBackgroundColour(self.buttonColor)
        openButton_Assign.Bind(wx.EVT_BUTTON,
                               self.pressOpenButton_Assign)
        # ファイル保存ボタン
        saveButton_Assign = wx.Button(self, -1, 'CSV保存')
        saveButton_Assign.SetBackgroundColour(self.buttonColor)
        saveButton_Assign.Bind(wx.EVT_BUTTON,
                               self.pressSaveButton_Assign)
        # ボタン用レイアウト
        buttonsLayout_Assign = wx.BoxSizer(wx.VERTICAL)
        buttonsLayout_Assign.Add(AssignButton)
        buttonsLayout_Assign.Add(wx.StaticLine(self), 0,
                                 wx.EXPAND | wx.TOP | wx.BOTTOM, 10)
        buttonsLayout_Assign.Add(notice_Assign, 0, wx.ALIGN_CENTER)
        buttonsLayout_Assign.Add(self.assignLabel, 0, wx.LEFT, 10)
        buttonsLayout_Assign.Add(openButton_Assign)
        buttonsLayout_Assign.Add(saveButton_Assign, 0, wx.TOP, 10)
        
        self.assignLayout = wx.BoxSizer(wx.HORIZONTAL)
        self.assignLayout.Add(buttonsLayout_Assign)
        self.assignLayout.Add(self.grid_Assign)
        
        #
        # 被割り当て部レイアウト===============================
        #
        
        notice_RAssign = wx.StaticText(self, -1, '■負担表')
        # ファイル名ラベル
        self.rAssignLabel = wx.StaticText(self, -1, 'none')
        # ファイル読込みボタン
        openButton_RAssign = wx.Button(self, -1, 'CSV読込')
        openButton_RAssign.SetBackgroundColour(self.buttonColor)
        openButton_RAssign.Bind(wx.EVT_BUTTON,
                                self.pressOpenButton_RAssign)
        # ファイル保存ボタン
        saveButton_RAssign = wx.Button(self, -1, 'CSV保存')
        saveButton_RAssign.SetBackgroundColour(self.buttonColor)
        saveButton_RAssign.Bind(wx.EVT_BUTTON,
                                self.pressSaveButton_RAssign)
        # ボタン用レイアウト
        buttonsLayout_RAssign = wx.BoxSizer(wx.VERTICAL)
        buttonsLayout_RAssign.Add(notice_RAssign, 0, wx.ALIGN_CENTER)
        buttonsLayout_RAssign.Add(self.rAssignLabel, 0, wx.LEFT, 10)
        buttonsLayout_RAssign.Add(openButton_RAssign)
        buttonsLayout_RAssign.Add(saveButton_RAssign, 0, wx.TOP, 10)
        
        # 割り当て日程一覧の部分
        notice_assignedDays = wx.StaticText(self, -1, '■割り当て日程の一覧')
        saveButton_AssignedDays = wx.Button(self, -1, 'CSV保存')
        saveButton_AssignedDays.SetBackgroundColour(self.buttonColor)
        saveButton_AssignedDays.Bind(wx.EVT_BUTTON,
                                     self.pressSaveButton_AssignedDays)
        
        self.AssignedDaysTitleLayout = wx.BoxSizer(wx.HORIZONTAL)
        self.AssignedDaysTitleLayout.Add(notice_assignedDays)
        self.AssignedDaysTitleLayout.Add(saveButton_AssignedDays)
        
        self.assignDaysLayout = wx.BoxSizer(wx.VERTICAL)
        self.assignDaysLayout.Add(self.AssignedDaysTitleLayout)
        self.assignDaysLayout.Add(self.grid_assignedDays)
        
        # 被割り当て部レイアウトの結合
        self.rAssignLayout = wx.BoxSizer(wx.HORIZONTAL)
        self.rAssignLayout.Add(buttonsLayout_RAssign)
        self.rAssignLayout.Add(self.grid_RAssign)
        self.rAssignLayout.Add(self.assignDaysLayout, 0, wx.LEFT, 10)
        
        #
        # 全体のレイアウト実装=======================================
        #
        
        self.mainLayout = wx.BoxSizer(wx.VERTICAL)
        self.mainLayout.Add(self.assignLayout)
        self.mainLayout.Add(wx.StaticLine(self), 0,
                            wx.EXPAND | wx.TOP | wx.BOTTOM, 10)
        self.mainLayout.Add(self.rAssignLayout)
        self.mainLayout.Fit(self)
        
        self.SetSizer(self.mainLayout)
