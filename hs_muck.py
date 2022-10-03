import asyncio
import os
import threading
import pyperclip
from ctypes import windll
from tkinter import *
from tkinter.ttk import *
import sys
import winreg
import re

error_lines = ''
try:
    _win_v = sys.getwindowsversion()
    if _win_v.major == 6 and _win_v.minor == 1:
        windll.user32.SetProcessDPIAware()
    else:
        windll.shcore.SetProcessDpiAwareness(2)
except Exception as e:
    error_lines += str(e) + '\n'

hs_dir = ''
try:
    reg_pos = r'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Hearthstone'
    handle = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_pos)
    hs_dir, _type = winreg.QueryValueEx(handle, 'InstallLocation')
    winreg.CloseKey(handle)
except Exception as e:
    error_lines += str(e) + '\n'
hs_dir += '\\Logs\\'
if not os.path.exists(hs_dir):
    error_lines += '找不到Log文件夹。关闭工具，把所有文件移动到\nHearthstone.exe所在文件夹后再试。'
HS_LOG_PATH = hs_dir + 'Power.log'

APPDATA_DIR = os.getenv('LOCALAPPDATA')
log_config_p0 = APPDATA_DIR + '\\Blizzard\\Hearthstone\\'
log_config_path = APPDATA_DIR + r'\Blizzard\Hearthstone\log.config'
if not os.path.exists(log_config_path):
    try:
        with open(log_config_path, 'w', encoding='utf-8') as ff:
            print("""[Achievements]
LogLevel=1
FilePrinting=True
ConsolePrinting=False
ScreenPrinting=False
Verbose=False
[Arena]
LogLevel=1
FilePrinting=True
ConsolePrinting=False
ScreenPrinting=False
Verbose=False
[FullScreenFX]
LogLevel=1
FilePrinting=True
ConsolePrinting=False
ScreenPrinting=False
Verbose=False
[LoadingScreen]
LogLevel=1
FilePrinting=True
ConsolePrinting=False
ScreenPrinting=False
Verbose=False
[Power]
LogLevel=1
FilePrinting=True
ConsolePrinting=False
ScreenPrinting=False
Verbose=True""", file=ff)
            error_lines += '已创建log.config，需重启游戏。\n'
    except Exception as e:
        error_lines += str(e) + '\n'
if not error_lines:
    error_lines = '（检测正常）\n'
R1 = re.compile(r'.*TAG_CHANGE Entity=\[entityName=(.*) id=(\d*) zone=PLAY zonePos=(\d*) cardId=(.*) player=(\d)] '
                r'tag=(.*) value=(.*)')
R2 = re.compile(r'FULL_ENTITY - Updating \[entityName=(.*) id=(\d*) zone=(.*) zonePos=(\d*) cardId=(.*) '
                r'player=(\d)] CardID=(.*)')
R3 = re.compile(r'FULL_ENTITY - Creating ID=(\d*) CardID=(.*)')
R4 = re.compile(r'tag=(.*) value=(.*)')
R5 = re.compile(r'SUB_SPELL_START - SpellPrefabGUID=(.*) Source=(\d*) TargetCount=(\d*)')
R6 = re.compile(r'Targets\[0] = \[entityName=(.*) id=(\d*) zone=(.*) zonePos=(\d*) cardId=(.*) player=(\d)]')
R7 = re.compile(r'Source = \[entityName=(.*) id=(\d*) zone=(.*) zonePos=(\d*) cardId=(.*) player=(\d)]')
label = None


class MainWindow(Tk):
    def __init__(self):
        global error_lines, label
        super().__init__()
        self.title('地标谜题辅助 右击窗口有说明')
        self.geometry('410x410+1480+40')
        self.wm_attributes('-topmost', True)
        self.fontsize = 12
        self.loop = None
        self.Label = Label(self, font=('consolas', self.fontsize, 'bold'), text=error_lines, image=None)
        label = self.Label
        self.Label.pack()
        self.Label.place(x=-1, y=-1)
        button_lines = error_lines + '\n单击按钮开始检测'
        self.Button1 = Button(self, text=button_lines, command=self.main_start)
        self.Button1.pack()
        self.Button1.place()

        self.Menu = Menu(self, tearoff=False)
        self.Menu.add_cascade(label='切换客户端语言为【英语】')
        self.Menu.add_cascade(label='点击这里复制套牌代码', command=self.copy_code)
        self.Menu.add_cascade(label='导入游戏，打天梯，进入谜题')
        self.Menu.add_cascade(label='随从和地标有隐藏词条，用1到20的数字表示')
        self.Menu.add_cascade(label='十个随从和十个地标一一对应，但一开始打乱了位置')
        self.Menu.add_cascade(label='对应随从的四个词条一定包括对应地标的两个词条，以此为依据反推')
        self.Menu.add_cascade(label='用手牌揭示随从或地标目前是？的词条，每回合两次')
        self.Menu.add_cascade(label='用技能判断随从是否对应当前地标的地标，星星特性说明猜对，每回合一次')
        self.Menu.add_cascade(label='回合结束后的烟火数量表示目前放对位置的总数')
        self.Menu.add_cascade(label='回合结束后你失去一点生命，你只有十个回合')
        self.Menu.add_cascade(label='第2版开始，辅助工具显示部分计算可得信息')
        self.Menu.add_separator()
        self.Menu.add_cascade(label='字体大小 +', command=self.font_plus)
        self.Menu.add_cascade(label='字体大小 -', command=self.font_minus)
        self.Menu.add_cascade(label='结束后可打开日志配置路径手动删除log.config', command=self.pop_dir)
        self.Menu.add_separator()
        self.Menu.add_cascade(label='第2版')
        self.bind('<Button-3>', self.popupmenu)
        self.logviewer = self.LogViewer(HS_LOG_PATH)
        self.mainloop()

    def popupmenu(self, event):
        try:
            self.Menu.post(event.x_root, event.y_root)
        finally:
            self.Menu.grab_release()

    def copy_code(event):
        pyperclip.copy(
            'AAECAaoIDLSKBLaKBKyfBNugBOCgBJbUBKDUBKnUBPzbBMviBJakBfCuBQmf1ASo2QS12QT03ASz3QS14gSl5ATF7QTK7QQA')

    def font_plus(self):
        self.fontsize += 1
        self.Label['font'] = ('consolas', self.fontsize, 'bold')

    def font_minus(self):
        if self.fontsize > 6:
            self.fontsize -= 1
            self.Label['font'] = ('consolas', self.fontsize, 'bold')

    def pop_dir(event):
        os.startfile(log_config_p0)

    def get_loop(self, loop):
        self.loop = loop
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    def main_start(self):
        self.Button1.destroy()
        coroutine1 = self.handler()
        new_loop = asyncio.new_event_loop()
        t = threading.Thread(target=self.get_loop, args=(new_loop,))
        t.daemon = True
        t.start()
        asyncio.run_coroutine_threadsafe(coroutine1, new_loop)

    class LogViewer:
        def __init__(self, path):
            self.path = path
            self.file = None
            self.last_size = 0
            self.pos = 0
            locations = {a: [0, [0, 0], ''] for a in range(20, 30)}
            minions = {a: [0, [0, 0, 0, 0], ''] for a in range(10, 20)}
            self.entities = {**locations, **minions}
            self.fireworks = 0
            self.fireworks_updated = False
            self.data = [[0] * 10 for _ in range(6)]
            self.score = {(location, minion): '0' for location in range(20, 30) for minion in range(10, 20)}
            # 0=初始 1=错 2=错但已经知道对 3=对
            self.full_answers = []
            self.output = ''
            self.updated = False

        def update(self):
            if not self.path:
                return
            if os.path.exists(self.path):
                if self.file is None:
                    self.file = open(self.path, 'r', encoding='utf-8')
                else:
                    size = os.path.getsize(self.path)
                    if size == self.last_size:
                        return
                    if size < self.last_size:
                        self.pos = 0
                    self.last_size = size
                    self.file.seek(self.pos)
                    lines = self.file.readlines()
                    self.pos = self.file.tell()
                    cache_type = ''
                    cache = []
                    for line in lines:
                        line0 = line[19:].strip()
                        if '() - ' in line0:
                            s_type, s_info = line0.split('() - ', 1)
                            s_info1 = s_info.strip()
                            level = (len(s_info) - len(s_info1)) // 2

                            if cache and level == 0:
                                self.cache_handle(cache_type, cache)
                                cache = []
                            cache.append((level, s_info1))
                            cache_type = s_type
                        # else:
                        #     print(line, end='')
                        #     print('特殊语句')
                    if cache:
                        self.cache_handle(cache_type, cache)
                    if self.updated:
                        self.cal_output()
            else:
                pass
                # print('文件消失，实体摆烂', self.path)
                # self.path = ''
                # self.file = None

        def cal_output(self):
            entities = self.entities.items()
            minions = sorted([a[1] for a in entities if a[0] < 20], key=lambda b: b[0])
            locations = sorted([a[1] for a in entities if a[0] >= 20], key=lambda b: b[0])
            for j in range(2):
                for i in range(10):
                    self.data[j][i] = locations[i][1][j]
            for j in range(4):
                for i in range(10):
                    self.data[j + 2][i] = minions[i][1][j]
            self.output = f'上次猜对{self.fireworks}个位置\n\n'
            for i in range(6):
                self.output += ''.join([f'{self.data[i][j] if self.data[i][j] else "?":>4}' for j in range(10)]) + (
                    '\n\n' if i == 1 else '\n')
            guesses_record = [en[0] for en in self.score.items() if en[1] == '3']
            for record in guesses_record:
                self.output += f'\n{self.entities[record[1]][2]}【在】' \
                               f'{self.entities[record[0]][2]}（{self.entities[record[0]][0]}）'
            guesses_record = [en[0] for en in self.score.items() if en[1] == '1']
            self.output += '\n'
            for record in guesses_record:
                self.output += f'\n{self.entities[record[1]][2]}【不在】' \
                               f'{self.entities[record[0]][2]}（{self.entities[record[0]][0]}）'

        def cache_handle(self, _type, cache):
            last_id = 0
            sub_spell_type = None
            if _type == 'PowerTaskList.DebugPrintPower':
                for log in cache:
                    line = log[1]
                    if line.startswith('TAG_CHANGE'):
                        if self.fireworks_updated:
                            minions = [en for en in self.entities.items() if en[0] < 20]
                            _new = (set((19 + i, [minion[0] for minion in minions if minion[1][0] == i][0]) for i in range(1, 11)), self.fireworks)
                            if len(self.full_answers) > 1:
                                _last = self.full_answers[-1]
                                _add = _new[0] - _last[0]
                                if len(_add) == 2:
                                    _lost = _last[0] - _new[0]
                                    if len(_lost) == 2:
                                        _cha = _new[1] - _last[1]
                                        if _cha == 2:
                                            for pair in _add:
                                                for t in range(20, 30):
                                                    self.score[(t, pair[1])] = '2'
                                                for t in range(10, 20):
                                                    self.score[(pair[0], t)] = '2'
                                                self.score[pair] = '3'
                                        elif _cha == -2:
                                            for pair in _lost:
                                                for t in range(20, 30):
                                                    self.score[(t, pair[1])] = '2'
                                                for t in range(10, 20):
                                                    self.score[(pair[0], t)] = '2'
                                                self.score[pair] = '3'
                                        elif _cha == 0:
                                            for pair in _add | _lost:
                                                if self.score[pair] == '0':
                                                    self.score[pair] = '1'
                                        else:
                                            pass
                            self.full_answers.append(_new)
                            self.fireworks_updated = False
                        result = R1.match(line)
                        if result:
                            _name = result.group(1)
                            _id = int(result.group(2))
                            _pos = int(result.group(3))
                            _tag = result.group(6)
                            _value = result.group(7)
                            if _tag == 'ZONE_POSITION':
                                if 10 <= _id < 30:
                                    self.entities[_id][0] = int(_value)
                                    self.entities[_id][2] = _name
                                    self.updated = True
                            elif 'CARDTEXT_ENTITY_' in _tag:
                                _tag_num = int(_tag[-1])
                                self.entities[_id][1][_tag_num] = (int(_value) - 30) % 21
                                self.updated = True
                        if line == 'TAG_CHANGE Entity=GameEntity tag=STATE value=COMPLETE':
                            self.cal_output()
                            # print(self.output)
                            label['text'] = self.output
                            self.full_answers = []
                            self.score = {(location, minion): '0' for location in range(20, 30) for minion in
                                          range(10, 20)}
                            self.updated = False
                    else:
                        result = R2.match(line)
                        if result:
                            last_id = int(result.group(2))
                            if 10 <= last_id < 30:
                                _name = result.group(1)
                                self.entities[last_id][2] = _name
                        result = R3.match(line)
                        if result:
                            last_id = int(result.group(1))
                        result = R4.match(line)
                        if result:
                            _tag = result.group(1)
                            _value = result.group(2)
                            if _tag == 'ZONE_POSITION':
                                if 10 <= last_id < 30:
                                    self.entities[last_id][0] = int(_value)
                                    self.updated = True
                            elif 'CARDTEXT_ENTITY_' in _tag:
                                _tag_num = int(_tag[-1])
                                self.entities[last_id][1][_tag_num] = (int(_value) - 30) % 21
                                self.updated = True
            elif _type == 'GameState.SendOption':
                for log in cache:
                    line = log[1]
                    if line == 'selectedOption=0 selectedSubOption=-1 selectedTarget=0 selectedPosition=0':
                        self.fireworks = 0
            elif _type == 'GameState.DebugPrintPower':
                for log in cache:
                    line = log[1]
                    result = R5.match(line)
                    if result:
                        sub_spell_type = result.group(1)
                    result = R6.match(line)
                    if result:
                        name = result.group(1)
                        _id = int(result.group(2))
                        _pos = int(result.group(4))
                        answer = True if sub_spell_type.startswith('LOOTFX_Confuse_Targeted_ImpactOnly_FX') else False
                        loc_id = [en[0] for en in self.entities.items() if en[1][0] == _pos][0]
                        if answer:
                            for t in range(20, 30):
                                self.score[(t, _id)] = '2'
                            for t in range(10, 20):
                                self.score[(loc_id, t)] = '2'
                            self.score[(loc_id, _id)] = '3'
                        else:
                            if self.score[(loc_id, _id)] == '0':
                                self.score[(loc_id, _id)] = '1'
                        self.updated = True
                    result = R7.match(line)
                    if result and sub_spell_type.startswith('ReuseFX_Generic_AE_Flare_Super'):
                        self.updated = True
                        self.fireworks_updated = True
                        self.fireworks += 1
                    # 'ReuseFX_Generic_AE_Flare_Super:d852894a04f4afd45b05f0e21cdd6524'  # 烟花
                    # 'ReuseFX_Mercenaries_SlowDown_Impact_Super:24c7d126f2e2c9f47bf6b2d375ad49c9' # 错
                    # 'LOOTFX_Confuse_Targeted_ImpactOnly_FX:ec7a4d1e90ce55b47bc2f7bdc46462b8'  # 对

        def __del__(self):
            if self.file:
                self.file.close()

    async def handler(self):
        while True:
            try:
                self.logviewer.update()
                if self.logviewer.updated:
                    # print(self.logviewer.output)
                    self.Label['text'] = self.logviewer.output
                    self.logviewer.updated = False
                await asyncio.sleep(0.005)
            except Exception as ee:
                print(ee, ee.args)


if __name__ == '__main__':
    MainWindow()
