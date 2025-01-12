# -*- coding: utf-8 -*-
# Before use,execute `CWeChatRobot.exe /regserver` in cmd by admin user
import datetime
import os
import ctypes
import pickle
import queue
import socket
import subprocess
import time
import json
import ctypes.wintypes
import socketserver
import threading
import uuid
import winreg

# need `pip install comtypes`
import comtypes.client
from comtypes.client import GetEvents
from comtypes.client import PumpEvents
import re
import requests
import urllib.parse
import random
import signal
import webbrowser
import sys
import configparser
import logging
import psutil
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

wxpusherOfficialAccountId = "gh_bf214c93111c"
wechatRobotProcessList = []
software_version = "Beta Version V0.0.6"
dispatcher = None
isEqualizationDistributionModel = True
skipSamePostIn1Second = True
socketServerThread = {"id": None, "port": 10808}
lastestMessage = {"sign": "", "msgid": ""}
# 创建一个锁对象
lock = threading.Lock()


def search_str_list(objects, target):
    for item in objects:
        try:
            if target in item:
                # print(f"Match found: {obj}")
                return item
        except Exception as e:
            continue
            # print("No match found.")
    return None


def parse_first_level(json_str):
    pattern = r'"(.*?)":\s*(".*?"|\d+)'
    matches = re.findall(pattern, json_str)
    first_level_data = {key: value for key, value in matches}
    return first_level_data


def convert_data_to_json(data, skipRegex=False):
    result = {}
    try:
        result = json.loads(data)
    except Exception as e:
        result = {}
    if skipRegex == False:
        # 使用正则表达式匹配字段值
        pattern = re.compile(r"<(\w+)><!\[CDATA\[(.*?)\]\]></\1>")
        matches = re.findall(pattern, data)
        for match in matches:
            result[match[0]] = match[1]

        # 匹配空字段
        empty_fields = re.findall(r"<(\w+)></\1>", data)
        for field in empty_fields:
            result[field] = ""

    if len(result.keys()) == 0:
        try:
            temp = parse_first_level(data)
            result["bizmsgfromuser"] = temp["wxid"]
            result["fromusername"] = temp["sender"]
            result["fromusername"] = temp["sender"]
            result["title"] = extract_middle_text(
                data, "<title><![CDATA[", "]]></title>"
            )
            result["url"] = extract_middle_text(data, "<url><![CDATA[", "]]></")
        except Exception as e:
            return {}
    return result


def get_url_content(url):
    response = requests.get(url, verify=False, allow_redirects=False)
    return response.text


def extract_links(text):
    pattern = r"http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+"
    links = re.findall(pattern, text)
    return links


def extract_middle_text(source, before_text, after_text, all_matches=False):
    results = []
    start_index = source.find(before_text)

    while start_index != -1:
        source_after_before_text = source[start_index + len(before_text) :]
        end_index = source_after_before_text.find(after_text)

        if end_index == -1:
            break

        results.append(source_after_before_text[:end_index])
        if not all_matches:
            break

        source = source_after_before_text[end_index + len(after_text) :]
        start_index = source.find(before_text)

    return results if all_matches else results[0] if results else ""


class _WeChatRobotClient:
    # 创建和管理与微信机器人的连接。这个类使用了单例模式；在整个程序运行期间，只会创建一个_WeChatRobotClient的实例。
    _instance = None

    @classmethod
    def instance(cls) -> "_WeChatRobotClient":
        if not cls._instance:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        # 分别用于控制微信机器人和接收微信事件
        self.robot = comtypes.client.CreateObject("WeChatRobot.CWeChatRobot")
        self.event = comtypes.client.CreateObject("WeChatRobot.RobotEvent")
        self.com_pid = self.robot.CStopRobotService(0)

    @classmethod
    def __del__(cls):
        # 如果_WeChatRobotClient的实例存在，就尝试结束与之关联的COM进程，并将_instance设置为None
        import psutil

        if cls._instance is not None:
            try:
                com_process = psutil.Process(cls._instance.com_pid)
                com_process.kill()
            except psutil.NoSuchProcess:
                pass
        cls._instance = None


class WeChatEventSink:
    """
    接收消息的默认回调，可以自定义，并将实例化对象作为StartReceiveMsgByEvent参数
    自定义的类需要包含以下所有成员
    """

    def OnGetMessageEvent(self, msg):
        pass
        # msg = json.loads(msg[0])
        # 打包后能显示，但是不打包的情况下不会显示这个，很奇怪
        # print("WeChatEventSink 接收消息的默认回调 ：", msg)


def messageHandler(data):
    msg = data.strip().decode("utf-8")
    rawmsg = json.loads(msg)
    # print(f"来消息了: ", msg)
    # print('wxpusherOfficialAccountId：',wxpusherOfficialAccountId)
    if wxpusherOfficialAccountId in msg:
        # print(f"自定义类收到的Received data: {msg}")
        parsedMsg = convert_data_to_json(msg)
        # print(f"来消息了: ", parsedMsg)
        try:
            if parsedMsg["fromusername"] == wxpusherOfficialAccountId:
                # 使用 with 语句获取锁
                with lock:
                    if (lastestMessage["msgid"] == parsedMsg["msgid"]) or (
                        parsedMsg["sign"] == lastestMessage["sign"]
                    ):
                        print(
                            f"检测到检测消息是同一条，跳过处理！",
                            # lastestMessage["msgid"],
                            # lastestMessage["sign"],
                            # " ---- ",
                            # parsedMsg["msgid"],
                            # parsedMsg["sign"],
                        )
                        return
                    else:
                        lastestMessage["msgid"] = parsedMsg["msgid"]
                        lastestMessage["sign"] = parsedMsg["sign"]
                messageUrl = parsedMsg["url"]
                # print(f"开始获取消息详细内容: ",messageUrl)
                if messageUrl:
                    messageContent = get_url_content(messageUrl)
                    if messageContent:
                        messageBody = extract_middle_text(
                            messageContent, "<body>", "</body>"
                        )
                        if messageBody:
                            findUrls = extract_links(messageBody)
                            # print(f"开始提取推送消息里的链接: ",findUrls)
                            if findUrls and len(findUrls):
                                mpUrl = search_str_list(
                                    findUrls, "https://mp.weixin.qq.com/"
                                ) or search_str_list(
                                    findUrls, "http://mp.weixin.qq.com/"
                                )
                                print(f"开始提取推送消息里的公众号文章链接: ", mpUrl)
                                if mpUrl:
                                    decodePostUrl = urllib.parse.unquote(mpUrl)
                                    if (
                                        check_and_execute(decodePostUrl, 1) == False
                                    ) and skipSamePostIn1Second:
                                        print(
                                            f"1秒内重复出现同一公众号文章链接，考虑到可能是不同微信订阅了同一个主题，故跳过: ",
                                            mpUrl,
                                        )
                                        return
                                    # 初始化COM线程
                                    # comtypes.CoInitialize()
                                    # 主线程中已经注入，此处禁止调用StartService和StopService
                                    global isEqualizationDistributionModel
                                    if isEqualizationDistributionModel:
                                        # 添加项目到队列
                                        global dispatcher
                                        dispatcher.add_item(decodePostUrl)
                                        # 开始调度
                                        dispatcher.dispatch()
                                    else:
                                        robot = comtypes.client.CreateObject(
                                            "WeChatRobot.CWeChatRobot"
                                        )
                                        event = comtypes.client.CreateObject(
                                            "WeChatRobot.RobotEvent"
                                        )
                                        wx = WeChatRobot(rawmsg["pid"], robot, event)
                                        print(f"开始打开链接: {decodePostUrl}")
                                        wx.OpenBrowser(decodePostUrl)
                                        delay = random.uniform(6, 7)
                                        print(f"等待 {delay:.2f} 秒...")
                                        time.sleep(delay)
                                        print("阅读该文章结束")
                                    # comtypes.CoUninitialize()  # 反初始化COM线程
        except Exception as e:
            print(f"该信息无法解析 data: {msg}")
            raise e


# 用于存储上次执行的时间戳和入参值
last_execution = {}
lock = threading.Lock()  # 创建线程锁

# time_interval  设置时间间隔（单位：秒）
def check_and_execute(input_value, time_interval=0.5):
    current_time = time.time()
    global last_execution
    # 使用线程锁来保证线程安全
    with lock:
        # 如果该入参值在上次执行的时间间隔内已经出现过，则不执行某个函数
        if input_value in last_execution and (
            current_time - last_execution[input_value] < time_interval
        ):
            # print(f"Input value '{input_value}' appeared too soon. Not executing the function.")
            return False
        else:
            # 记录当前时间戳
            last_execution[input_value] = current_time
            # print(f"Input value '{input_value}' is unique. Executing the function.")
            return True


class WechatMessageHandler(socketserver.BaseRequestHandler):
    def handle(self):
        comtypes.CoInitialize()
        conn = self.request
        ptr_data = b""  # 初始化一个空字节串用于存储接收到的数据

        try:
            while True:
                data = conn.recv(10240)  # 接收数据，每次最多接收1024字节
                if not data:
                    break  # 如果接收到的数据为空，跳出循环
                ptr_data += data  # 将接收到的数据添加到ptr_data

                if data[-1] == 0xA:  # 如果最后一个字节是换行符（0xA），则跳出循环
                    break

            # msg = json.loads(ptr_data.decode('utf-8'))  # 将接收到的数据解码为JSON格式的消息
            messageHandler(ptr_data)  # 调用消息回调函数处理消息
            conn.sendall(ptr_data)  # 发送确认消息给客户端
        except OSError:
            pass  # 捕获操作系统错误
        except json.JSONDecodeError:
            pass  # 捕获JSON解析错误
        finally:
            conn.close()  # 关闭连接
            comtypes.CoUninitialize()  # 反初始化COM线程


class ChatSession:
    def __init__(self, pid, robot, wxid):
        self.pid = pid
        self.robot = robot
        self.chat_with = wxid

    def SendText(self, msg):
        return self.robot.CSendText(self.pid, self.chat_with, msg)

    def SendImage(self, img_path):
        return self.robot.CSendImage(self.pid, self.chat_with, img_path)

    def SendFile(self, filepath):
        return self.robot.CSendFile(self.pid, self.chat_with, filepath)

    def SendMp4(self, mp4path):
        return self.robot.CSendImage(self.pid, self.chat_with, mp4path)

    def SendArticle(self, title, abstract, url, img_path=None):
        return self.robot.CSendArticle(
            self.pid, self.chat_with, title, abstract, url, img_path
        )

    def SendCard(self, shared_wxid, nickname):
        return self.robot.CSendCard(self.pid, self.chat_with, shared_wxid, nickname)

    def SendAtText(self, wxid: list or str or tuple, msg, auto_nickname=True):
        if "@chatroom" not in self.chat_with:
            return 1
        return self.robot.CSendAtText(
            self.pid, self.chat_with, wxid, msg, auto_nickname
        )

    def SendAppMsg(self, appid):
        return self.robot.CSendAppMsg(self.pid, self.chat_with, appid)


class WeChatRobot:
    def __init__(self, pid: int = 0, robot=None, event=None):
        self.pid = pid
        self.robot = robot or _WeChatRobotClient.instance().robot
        self.event = event or _WeChatRobotClient.instance().event
        self.AddressBook = []

    def StartService(self) -> int:
        """
        注入DLL到微信以启动服务

        Returns
        -------
        int
            0成功,非0失败.

        """
        status = self.robot.CStartRobotService(self.pid)
        return status

    def IsWxLogin(self) -> int:
        """
        获取微信登录状态

        Returns
        -------
        bool
            微信登录状态.

        """
        return self.robot.CIsWxLogin(self.pid)

    def SendText(self, receiver: str, msg: str) -> int:
        """
        发送文本消息

        Parameters
        ----------
        receiver : str
            消息接收者wxid.
        msg : str
            消息内容.

        Returns
        -------
        int
            0成功,非0失败.

        """
        return self.robot.CSendText(self.pid, receiver, msg)

    def SendImage(self, receiver: str, img_path: str) -> int:
        """
        发送图片消息

        Parameters
        ----------
        receiver : str
            消息接收者wxid.
        img_path : str
            图片绝对路径.

        Returns
        -------
        int
            0成功,非0失败.

        """
        return self.robot.CSendImage(self.pid, receiver, img_path)

    def SendFile(self, receiver: str, filepath: str) -> int:
        """
        发送文件

        Parameters
        ----------
        receiver : str
            消息接收者wxid.
        filepath : str
            文件绝对路径.

        Returns
        -------
        int
            0成功,非0失败.

        """
        return self.robot.CSendFile(self.pid, receiver, filepath)

    def SendArticle(
        self,
        receiver: str,
        title: str,
        abstract: str,
        url: str,
        img_path: str or None = None,
    ) -> int:
        """
        发送XML文章

        Parameters
        ----------
        receiver : str
            消息接收者wxid.
        title : str
            消息卡片标题.
        abstract : str
            消息卡片摘要.
        url : str
            文章链接.
        img_path : str or None, optional
            消息卡片显示的图片绝对路径，不需要可以不指定. The default is None.

        Returns
        -------
        int
            0成功,非0失败.

        """
        return self.robot.CSendArticle(
            self.pid, receiver, title, abstract, url, img_path
        )

    def SendCard(self, receiver: str, shared_wxid: str, nickname: str) -> int:
        """
        发送名片

        Parameters
        ----------
        receiver : str
            消息接收者wxid.
        shared_wxid : str
            被分享人wxid.
        nickname : str
            名片显示的昵称.

        Returns
        -------
        int
            0成功,非0失败.

        """
        return self.robot.CSendCard(self.pid, receiver, shared_wxid, nickname)

    def SendAtText(
        self,
        chatroom_id: str,
        at_users: list or str or tuple,
        msg: str,
        auto_nickname: bool = True,
    ) -> int:
        """
        发送群艾特消息，艾特所有人可以将AtUsers设置为`notify@all`
        无目标群管理权限请勿使用艾特所有人
        Parameters
        ----------
        chatroom_id : str
            群聊ID.
        at_users : list or str or tuple
            被艾特的人列表.
        msg : str
            消息内容.
        auto_nickname : bool, optional
            是否自动填充被艾特人昵称. 默认自动填充.

        Returns
        -------
        int
            0成功,非0失败.

        """
        if "@chatroom" not in chatroom_id:
            return 1
        return self.robot.CSendAtText(
            self.pid, chatroom_id, at_users, msg, auto_nickname
        )

    def GetSelfInfo(self) -> dict:
        """
        获取个人信息

        Returns
        -------
        dict
            调用成功返回个人信息，否则返回空字典.

        """
        self_info = self.robot.CGetSelfInfo(self.pid)
        return json.loads(self_info)

    def StopService(self) -> int:
        """
        停止服务，会将DLL从微信进程中卸载

        Returns
        -------
        int
            COM进程pid.

        """
        com_pid = self.robot.CStopRobotService(self.pid)
        return com_pid

    def GetAddressBook(self) -> list:
        """
        获取联系人列表

        Returns
        -------
        list
            调用成功返回通讯录列表，调用失败返回空列表.

        """
        try:
            friend_tuple = self.robot.CGetFriendList(self.pid)
            self.AddressBook = [dict(i) for i in list(friend_tuple)]
        except IndexError:
            self.AddressBook = []
        return self.AddressBook

    def GetFriendList(self) -> list:
        """
        从通讯录列表中筛选出好友列表

        Returns
        -------
        list
            好友列表.

        """
        if not self.AddressBook:
            self.GetAddressBook()
        friend_list = [
            item
            for item in self.AddressBook
            if (item["wxType"] == 3 and item["wxid"][0:3] != "gh_")
        ]
        return friend_list

    def GetChatRoomList(self) -> list:
        """
        从通讯录列表中筛选出群聊列表

        Returns
        -------
        list
            群聊列表.

        """
        if not self.AddressBook:
            self.GetAddressBook()
        chatroom_list = [item for item in self.AddressBook if item["wxType"] == 2]
        return chatroom_list

    def GetOfficialAccountList(self) -> list:
        """
        从通讯录列表中筛选出公众号列表

        Returns
        -------
        list
            公众号列表.

        """
        if not self.AddressBook:
            self.GetAddressBook()
        official_account_list = [
            item
            for item in self.AddressBook
            if (item["wxType"] == 3 and item["wxid"][0:3] == "gh_")
        ]
        return official_account_list

    def GetFriendByWxRemark(self, remark: str) -> dict or None:
        """
        通过备注搜索联系人

        Parameters
        ----------
        remark : str
            好友备注.

        Returns
        -------
        dict or None
            搜索到返回联系人信息，否则返回None.

        """
        if not self.AddressBook:
            self.GetAddressBook()
        for item in self.AddressBook:
            if item["wxRemark"] == remark:
                return item
        return None

    def GetFriendByWxNumber(self, wx_number: str) -> dict or None:
        """
        通过微信号搜索联系人

        Parameters
        ----------
        wx_number : str
            联系人微信号.

        Returns
        -------
        dict or None
            搜索到返回联系人信息，否则返回None.

        """
        if not self.AddressBook:
            self.GetAddressBook()
        for item in self.AddressBook:
            if item["wxNumber"] == wx_number:
                return item
        return None

    def GetFriendByWxNickName(self, nickname: str) -> dict or None:
        """
        通过昵称搜索联系人

        Parameters
        ----------
        nickname : str
            联系人昵称.

        Returns
        -------
        dict or None
            搜索到返回联系人信息，否则返回None.

        """
        if not self.AddressBook:
            self.GetAddressBook()
        for item in self.AddressBook:
            if item["wxNickName"] == nickname:
                return item
        return None

    def GetChatSession(self, wxid: str) -> "ChatSession":
        """
        创建一个会话，没太大用处

        Parameters
        ----------
        wxid : str
            联系人wxid.

        Returns
        -------
        'ChatSession'
            返回ChatSession类.

        """
        return ChatSession(self.pid, self.robot, wxid)

    def GetWxUserInfo(self, wxid: str) -> dict:
        """
        通过wxid查询联系人信息

        Parameters
        ----------
        wxid : str
            联系人wxid.

        Returns
        -------
        dict
            联系人信息.

        """
        userinfo = self.robot.CGetWxUserInfo(self.pid, wxid)
        return json.loads(userinfo)

    def GetChatRoomMembers(self, chatroom_id: str) -> dict or None:
        """
        获取群成员信息

        Parameters
        ----------
        chatroom_id : str
            群聊id.

        Returns
        -------
        dict or None
            获取成功返回群成员信息，失败返回None.

        """
        info = dict(self.robot.CGetChatRoomMembers(self.pid, chatroom_id))
        if not info:
            return None
        members = info["members"].split("^G")
        data = self.GetWxUserInfo(chatroom_id)
        data["members"] = []
        for member in members:
            member_info = self.GetWxUserInfo(member)
            data["members"].append(member_info)
        return data

    def CheckFriendStatus(self, wxid: str) -> int:
        """
        获取好友状态码

        Parameters
        ----------
        wxid : str
            好友wxid.

        Returns
        -------
        int
            0x0: 'Unknown',
            0xB0:'被删除',
            0xB1:'是好友',
            0xB2:'已拉黑',
            0xB5:'被拉黑',

        """
        return self.robot.CCheckFriendStatus(self.pid, wxid)

    # 接收消息的函数
    def StartReceiveMessage(self, port: int = 10808) -> int:
        """
        启动接收消息Hook

        Parameters
        ----------
        port : int
            socket的监听端口号.如果要使用连接点回调，则将端口号设置为0.

        Returns
        -------
        int
            启动成功返回0,失败返回非0值.

        """
        status = self.robot.CStartReceiveMessage(self.pid, port)
        return status

    def StopReceiveMessage(self) -> int:
        """
        停止接收消息Hook

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        status = self.robot.CStopReceiveMessage(self.pid)
        return status

    def GetDbHandles(self) -> dict:
        """
        获取数据库句柄和表信息

        Returns
        -------
        dict
            数据库句柄和表信息.

        """
        tables_tuple = self.robot.CGetDbHandles(self.pid)
        tables = [dict(i) for i in tables_tuple]
        dbs = {}
        for table in tables:
            dbname = table["dbname"]
            if dbname not in dbs.keys():
                dbs[dbname] = {"Handle": table["Handle"], "tables": []}
            dbs[dbname]["tables"].append(
                {
                    "name": table["name"],
                    "tbl_name": table["tbl_name"],
                    "root_page": table["rootpage"],
                    "sql": table["sql"],
                }
            )
        return dbs

    def ExecuteSQL(self, handle: int, sql: str) -> list:
        """
        执行SQL

        Parameters
        ----------
        handle : int
            数据库句柄.
        sql : str
            SQL.

        Returns
        -------
        list
            查询结果.

        """
        result = self.robot.CExecuteSQL(self.pid, handle, sql)
        if len(result) == 0:
            return []
        query_list = []
        keys = list(result[0])
        for item in result[1:]:
            query_dict = {}
            for key, value in zip(keys, item):
                query_dict[key] = (
                    value if not isinstance(value, tuple) else bytes(value)
                )
            query_list.append(query_dict)
        return query_list

    def BackupSQLiteDB(self, handle: int, filepath: str) -> int:
        """
        备份数据库

        Parameters
        ----------
        handle : int
            数据库句柄.
        filepath : int
            备份文件保存位置.

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        filepath = filepath.replace("/", "\\")
        save_path = filepath.replace(filepath.split("\\")[-1], "")
        if not os.path.exists(save_path):
            os.makedirs(save_path)
        return self.robot.CBackupSQLiteDB(self.pid, handle, filepath)

    def VerifyFriendApply(self, v3: str, v4: str) -> int:
        """
        通过好友请求

        Parameters
        ----------
        v3 : str
            v3数据(encryptUserName).
        v4 : str
            v4数据(ticket).

        Returns
        -------
        int
            成功返回0,失败返回非0值..

        """
        return self.robot.CVerifyFriendApply(self.pid, v3, v4)

    def AddFriendByWxid(self, wxid: str, message: str or None) -> int:
        """
        wxid加好友

        Parameters
        ----------
        wxid : str
            要添加的wxid.
        message : str or None
            验证信息.

        Returns
        -------
        int
            请求发送成功返回0,失败返回非0值.

        """
        return self.robot.CAddFriendByWxid(self.pid, wxid, message)

    def AddFriendByV3(self, v3: str, message: str or None, add_type: int = 0x6) -> int:
        """
        v3数据加好友

        Parameters
        ----------
        v3 : str
            v3数据(encryptUserName).
        message : str or None
            验证信息.
        add_type : int
            添加方式(来源).手机号: 0xF;微信号: 0x3;QQ号: 0x1;朋友验证消息: 0x6.

        Returns
        -------
        int
            请求发送成功返回0,失败返回非0值.

        """
        return self.robot.CAddFriendByV3(self.pid, v3, message, add_type)

    def GetWeChatVer(self) -> str:
        """
        获取微信版本号

        Returns
        -------
        str
            微信版本号.

        """
        return self.robot.CGetWeChatVer()

    def GetUserInfoByNet(self, keyword: str) -> dict or None:
        """
        网络查询用户信息

        Parameters
        ----------
        keyword : str
            查询关键字，可以是微信号、手机号、QQ号.

        Returns
        -------
        dict or None
            查询成功返回用户信息,查询失败返回None.

        """
        userinfo = self.robot.CSearchContactByNet(self.pid, keyword)
        if userinfo:
            return dict(userinfo)
        return None

    def AddBrandContact(self, public_id: str) -> int:
        """
        关注公众号

        Parameters
        ----------
        public_id : str
            公众号id.

        Returns
        -------
        int
            请求成功返回0,失败返回非0值.

        """
        return self.robot.CAddBrandContact(self.pid, public_id)

    def ChangeWeChatVer(self, version: str) -> int:
        """
        自定义微信版本号，一定程度上防止自动更新

        Parameters
        ----------
        version : str
            版本号，类似`3.7.0.26`

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        return self.robot.CChangeWeChatVer(self.pid, version)

    def HookImageMsg(self, save_path: str) -> int:
        """
        开始Hook未加密图片

        Parameters
        ----------
        save_path : str
            图片保存路径(绝对路径).

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        return self.robot.CHookImageMsg(self.pid, save_path)

    def UnHookImageMsg(self) -> int:
        """
        取消Hook未加密图片

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        return self.robot.CUnHookImageMsg(self.pid)

    def HookVoiceMsg(self, save_path: str) -> int:
        """
        开始Hook语音消息

        Parameters
        ----------
        save_path : str
            语音保存路径(绝对路径).

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        return self.robot.CHookVoiceMsg(self.pid, save_path)

    def UnHookVoiceMsg(self) -> int:
        """
        取消Hook语音消息

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        return self.robot.CUnHookVoiceMsg(self.pid)

    def DeleteUser(self, wxid: str) -> int:
        """
        删除好友

        Parameters
        ----------
        wxid : str
            被删除好友wxid.

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        return self.robot.CDeleteUser(self.pid, wxid)

    def SendAppMsg(self, wxid: str, appid: str) -> int:
        """
        发送小程序

        Parameters
        ----------
        wxid : str
            消息接收者wxid.
        appid : str
            小程序id (在xml中是username，不是appid).

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        return self.robot.CSendAppMsg(self.pid, wxid, appid)

    def EditRemark(self, wxid: str, remark: str or None) -> int:
        """
        修改好友或群聊备注

        Parameters
        ----------
        wxid : str
            wxid或chatroom_id.
        remark : str or None
            要修改的备注.

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        return self.robot.CEditRemark(self.pid, wxid, remark)

    def SetChatRoomName(self, chatroom_id: str, name: str) -> int:
        """
        修改群名称.请确认具有相关权限再调用。

        Parameters
        ----------
        chatroom_id : str
            群聊id.
        name : str
            要修改为的群名称.

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        return self.robot.CSetChatRoomName(self.pid, chatroom_id, name)

    def SetChatRoomAnnouncement(
        self, chatroom_id: str, announcement: str or None
    ) -> int:
        """
        设置群公告.请确认具有相关权限再调用。

        Parameters
        ----------
        chatroom_id : str
            群聊id.
        announcement : str or None
            公告内容.

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        return self.robot.CSetChatRoomAnnouncement(self.pid, chatroom_id, announcement)

    def SetChatRoomSelfNickname(self, chatroom_id: str, nickname: str) -> int:
        """
        设置群内个人昵称

        Parameters
        ----------
        chatroom_id : str
            群聊id.
        nickname : str
            要修改为的昵称.

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        return self.robot.CSetChatRoomSelfNickname(self.pid, chatroom_id, nickname)

    def GetChatRoomMemberNickname(self, chatroom_id: str, wxid: str) -> str:
        """
        获取群成员昵称

        Parameters
        ----------
        chatroom_id : str
            群聊id.
        wxid : str
            群成员wxid.

        Returns
        -------
        str
            成功返回群成员昵称,失败返回空字符串.

        """
        return self.robot.CGetChatRoomMemberNickname(self.pid, chatroom_id, wxid)

    def DelChatRoomMember(
        self, chatroom_id: str, wxid_list: str or list or tuple
    ) -> int:
        """
        删除群成员.请确认具有相关权限再调用。

        Parameters
        ----------
        chatroom_id : str
            群聊id.
        wxid_list : str or list or tuple
            要删除的成员wxid或wxid列表.

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        return self.robot.CDelChatRoomMember(self.pid, chatroom_id, wxid_list)

    def AddChatRoomMember(
        self, chatroom_id: str, wxid_list: str or list or tuple
    ) -> int:
        """
        添加群成员.请确认具有相关权限再调用。

        Parameters
        ----------
        chatroom_id : str
            群聊id.
        wxid_list : str or list or tuple
            要添加的成员wxid或wxid列表.

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        return self.robot.CAddChatRoomMember(self.pid, chatroom_id, wxid_list)

    def OpenBrowser(self, url: str) -> int:
        """
        打开微信内置浏览器

        Parameters
        ----------
        url : str
            目标网页url.

        Returns
        -------
        int
            成功返回0,失败返回非0值.

        """
        return self.robot.COpenBrowser(self.pid, url)

    def GetHistoryPublicMsg(self, public_id: str, offset: str = "") -> str:
        """
        获取公众号历史消息，一次获取十条推送记录

        Parameters
        ----------
        public_id : str
            公众号id.
        offset : str, optional
            起始偏移，为空的话则从新到久获取十条，该值可从返回数据中取得. The default is "".

        Returns
        -------
        str
            成功返回json数据，失败返回错误信息或空字符串.

        """
        ret = self.robot.CGetHistoryPublicMsg(self.pid, public_id, offset)[0]
        try:
            ret = json.loads(ret)
        except json.JSONDecodeError:
            pass
        return ret

    def ForwardMessage(self, wxid: str, msgid: int) -> int:
        """
        转发消息，只支持单条转发

        Parameters
        ----------
        wxid : str
            消息接收人wxid.
        msgid : int
            消息id，可以在实时消息接口中获取.

        Returns
        -------
        int
            成功返回0，失败返回非0值.

        """
        return self.robot.CForwardMessage(self.pid, wxid, msgid)

    def GetQrcodeImage(self) -> bytes:
        """
        获取二维码，同时切换到扫码登录

        Returns
        -------
        bytes
            二维码bytes数据.
        You can convert it to image object,like this:
        >>> from io import BytesIO
        >>> from PIL import Image
        >>> buf = wx.GetQrcodeImage()
        >>> image = Image.open(BytesIO(buf)).convert("L")
        >>> image.save('./qrcode.png')

        """
        data = self.robot.CGetQrcodeImage(self.pid)
        return bytes(data)

    def GetA8Key(self, url: str) -> dict or str:
        """
        获取A8Key

        Parameters
        ----------
        url : str
            公众号文章链接.

        Returns
        -------
        dict
            成功返回A8Key信息，失败返回空字符串.

        """
        ret = self.robot.CGetA8Key(self.pid, url)
        try:
            ret = json.loads(ret)
        except json.JSONDecodeError:
            pass
        return ret

    def SendXmlMsg(self, wxid: str, xml: str, img_path: str = "") -> int:
        """
        发送原始xml消息

        Parameters
        ----------
        wxid : str
            消息接收人.
        xml : str
            xml内容.
        img_path : str, optional
            图片路径. 默认为空.

        Returns
        -------
        int
            发送成功返回0，发送失败返回非0值.

        """
        return self.robot.CSendXmlMsg(self.pid, wxid, xml, img_path)

    def Logout(self) -> int:
        """
        退出登录

        Returns
        -------
        int
            成功返回0，失败返回非0值.

        """
        return self.robot.CLogout(self.pid)

    def GetTransfer(self, wxid: str, transcationid: str, transferid: str) -> int:
        """
        收款

        Parameters
        ----------
        wxid : str
            转账人wxid.
        transcationid : str
            从转账消息xml中获取.
        transferid : str
            从转账消息xml中获取.

        Returns
        -------
        int
            成功返回0，失败返回非0值.

        """
        return self.robot.CGetTransfer(self.pid, wxid, transcationid, transferid)

    def SendEmotion(self, wxid: str, img_path: str) -> int:
        """
        发送图片消息

        Parameters
        ----------
        wxid : str
            消息接收者wxid.
        img_path : str
            图片绝对路径.

        Returns
        -------
        int
            0成功,非0失败.

        """
        return self.robot.CSendEmotion(self.pid, wxid, img_path)

    def GetMsgCDN(self, msgid: int) -> str:
        """
        下载图片、视频、文件

        Parameters
        ----------
        msgid : int
            msgid.

        Returns
        -------
        str
            成功返回文件路径，失败返回空字符串.

        """
        path = self.robot.CGetMsgCDN(self.pid, msgid)
        if path != "":
            while not os.path.exists(path):
                time.sleep(0.5)
        return path


def get_wechat_pid_list() -> list:
    """
    获取所有微信pid

    Returns
    -------
    list
        微信pid列表.

    """
    import psutil

    pid_list = []
    process_list = psutil.pids()
    for pid in process_list:
        try:
            if psutil.Process(pid).name() == "WeChat.exe":
                pid_list.append(pid)
        except psutil.NoSuchProcess:
            pass
    return pid_list


def start_wechat() -> "WeChatRobot" or None:
    """
    启动微信

    Returns
    -------
    WeChatRobot or None
        成功返回WeChatRobot对象,失败返回None.

    """
    pid = _WeChatRobotClient.instance().robot.CStartWeChat()
    if pid != 0:
        return WeChatRobot(pid)
    return None


def register_msg_event(
    wx_pid: int, event_sink: "WeChatEventSink" or None = None
) -> None:
    """
    通过COM组件连接点接收消息，真正的回调
    只会收到wx_pid对应的微信消息

    Parameters
    ----------
    wx_pid: 微信PID
    event_sink : object, optional
        回调的实现类，该类要继承`WeChatEventSink`类或实现其中的方法.

    Returns
    -------
    None
        .

    """
    robot = comtypes.client.CreateObject("WeChatRobot.CWeChatRobot")
    event = comtypes.client.CreateObject("WeChatRobot.RobotEvent")
    wx = WeChatRobot(wx_pid, robot, event)
    # _WeChatRobotClient.instance()
    event = wx.event
    if event is not None:
        sink = event_sink or WeChatEventSink()
        connection_point = GetEvents(event, sink)
        assert connection_point is not None
        event.CRegisterWxPidWithCookie(wx_pid, connection_point.cookie)
        while True:
            try:
                PumpEvents(2)
            except KeyboardInterrupt:
                break
        del connection_point


def start_socket_server(
    port: int = 10808,
    request_handler: "socketserver.BaseRequestHandler" = socketserver.BaseRequestHandler,
    main_thread=True,
) -> int or None:
    """
    创建消息监听线程

    Parameters
    ----------
    port : int
        socket的监听端口号.

    request_handler : ReceiveMsgBaseServer
        用于处理消息的类，需要继承自socketserver.BaseRequestHandler或ReceiveMsgBaseServer

    main_thread : bool
        是否在主线程中启动server

    Returns
    -------
    int or None
        main_thread为False时返回线程id,否则返回None.

    """
    ip_port = ("127.0.0.1", port)
    try:
        s = socketserver.ThreadingTCPServer(ip_port, request_handler)
        if main_thread:
            s.serve_forever()
        else:
            socket_server = threading.Thread(target=s.serve_forever)
            socket_server.daemon = True
            socket_server.start()
            return socket_server.ident
    except KeyboardInterrupt:
        pass
    except Exception as e:
        print(e)
    return None


def stop_socket_server(thread_id: int) -> None:
    """
    强制结束消息监听线程

    Parameters
    ----------
    thread_id : int
        消息监听线程ID.

    Returns
    -------
    None
        .

    """
    if not thread_id:
        return
    import inspect

    try:
        tid = comtypes.c_long(thread_id)
        res = 0
        if not inspect.isclass(SystemExit):
            exec_type = type(SystemExit)
            res = comtypes.pythonapi.PyThreadState_SetAsyncExc(
                tid, comtypes.py_object(exec_type)
            )
        if res == 0:
            raise ValueError("invalid thread id")
        elif res != 1:
            ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, None)
            raise SystemError("PyThreadState_SetAsyncExc failed")
    except (ValueError, SystemError):
        pass


def exception_hook(exc_type, exc_value, exc_traceback):
    print("程序执行出错，请前往 error.log 查看错误信息")
    if issubclass(exc_type, KeyboardInterrupt):
        print("检测到用户同时按下 Ctrl+C，开始终止当前进程……")
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
    else:
        logging.error(
            "Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback)
        )


def set_up_logger():
    logging.basicConfig(filename="./error.log", level=logging.ERROR, encoding="utf-8")
    sys.excepthook = exception_hook


defaultConfig = {
    "ui_port": 28682,
    "hook_port": 10808,
    "main_thead_port": 28683,
    "robot_thead_port": 28684,
}


def load_config():
    config = configparser.ConfigParser()
    # 如果配置文件存在且不为空，则加载配置
    if os.path.exists("config.ini") and os.stat("config.ini").st_size > 0:
        config.read("config.ini")
    else:
        # 否则，写入默认配置
        config["Settings"] = defaultConfig
        with open("config.ini", "w") as config_file:
            config.write(config_file)
    return config


config = load_config()
# 获取配置值
ui_port = int(config.get("Settings", "ui_port"))
hook_port = int(config.get("Settings", "hook_port"))


def start_hello(instance, wechatBotIndex):
    self_info = instance.GetSelfInfo()
    print(f"第[{wechatBotIndex}] 号微信昵称为：{str(self_info.get('wxNickName'))}")
    chat_with = instance.GetFriendByWxNickName(
        "文件传输助手"
    ) or instance.GetFriendByWxNumber("filehelper")
    if chat_with:
        print(f"第[{wechatBotIndex}] 号微信已经成功获取到 文件传输助手，开始发送 启动欢迎提示语 ▷▷▷ ")
    else:
        print(
            f"第[{wechatBotIndex}] 号微信未取到 文件传输助手 联系人，无法发送 启动欢迎语 ，请搜索文件传输助手，将其添加到通讯录再试试吧~",
            instance.GetFriendByWxNickName("文件传输助手"),
            instance.GetFriendByWxNumber("filehelper"),
        )
        return False
    filehelper = instance.GetChatSession(chat_with.get("wxid"))
    filehelper.SendText(
        f"第[{wechatBotIndex}] 号微信当前过检测微信个人昵称：{str(self_info.get('wxNickName'))}，欢迎使用 幻生版免费过检测机器人"
    )
    return str(self_info.get("wxNickName"))


def show_interfaces():
    robot = WeChatRobot(0).robot
    print(robot.CGetWeChatVer())
    interfaces = [i for i in dir(robot) if "_" not in i and i[0] == "C"]
    for interface in interfaces:
        print(interface)


def statusMsg(status):
    if status == 0:
        return "操作成功"
    elif status == 1:
        return "操作失败"
    else:
        return "未知错误"


def search_obj_list(objects, keyName, target):
    # print(f"Searching for target {target} in objects...")
    for obj in objects:
        try:
            value = obj[keyName]
            # print(f"Checking {keyName} in object: {value}")
            if value == target:
                # print(f"Match found: {obj}")
                return obj
        except Exception as e:
            continue
            # print("No match found.")
    return None


# 返回一个可用的进程号列表
def get_available_pids(n, min_pid=0, max_pid=None):
    """
    返回当前系统未使用的指定数量的进程号列表
    :param n: 要返回的进程号数量
    :param min_pid: 进程号的最小值，默认为0
    :param max_pid: 进程号的最大值，默认为None，表示使用系统最大进程号
    :return: 未使用的进程号列表，如果未找到足够的进程号，则返回空列表
    """
    current_pids = {p.pid for p in psutil.process_iter(["pid"])}
    max_pid = max_pid or os.sysconf("SC_CLK_TCK") * os.sysconf("SC_CHILD_MAX")
    available_pids = [
        pid for pid in range(min_pid, max_pid + 1) if pid not in current_pids
    ]
    return available_pids[:n]


def get_unused_ports(num_ports, min_port=1024, max_port=65535):
    """
    返回当前系统未使用的指定数量的端口号列表

    :param num_ports: 要返回的端口号数量
    :param min_port: 端口号的最小值，默认为1024
    :param max_port: 端口号的最大值，默认为65535
    :return: 未使用的端口号列表，如果未找到足够的端口号，则返回空列表
    """
    unused_ports = []
    try:
        for port in range(min_port, max_port + 1):
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                # 检查端口是否可用
                result = s.connect_ex(("127.0.0.1", port))
                if result != 0:
                    unused_ports.append(port)
                if len(unused_ports) == num_ports:
                    break
    except Exception as e:
        print(f"发生异常: {e}")
    return unused_ports


def find_and_kill_process_using_port(port):
    # 遍历系统所有进程
    for proc in psutil.process_iter(["pid", "name", "connections"]):
        try:
            # 获取进程的连接信息
            connections = proc.connections()
            for conn in connections:
                # 判断是否有连接使用了指定的端口
                if conn.laddr.port == port:
                    print(
                        f"Found process {proc.pid} ({proc.name()}) using port {port}. Killing process..."
                    )
                    proc.kill()  # 强制终止进程
                    break
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass


def startListen(fastStartHelper=True):
    kill_all_robot_processes()
    allProcessingWechatPids = get_wechat_pid_list()
    listenWechatPids = []
    if len(allProcessingWechatPids) < 1:
        startWechatChoose = True
        if fastStartHelper == False:
            startWechatChoose = input("未找到已启动的微信，请问是否需要启动微信？请输入 y/n：") == "y"
        if startWechatChoose:
            startWxResult = start_wechat()
            if startWxResult:
                allProcessingWechatPids = get_wechat_pid_list()
                if len(allProcessingWechatPids) < 1:
                    print(f"启动异常，停止运行！")
                    return
        else:
            print(f"自动启动微信失败，请手动启动微信再试吧 ~ ")
            return
    if len(allProcessingWechatPids) > 1:
        isEqualization = "y"
        if fastStartHelper == False:
            isEqualization = input(
                "是否均衡分配，检测文章随机分配给当前可过检的微信，输入 y/直接回车 表示系统随机分配微信过检，输入 n/其他 表示检测文章仅订阅推送的微信过检："
            )
        isEqualizationDistributionModel = isEqualization == "y"
        print(f"{'已' if isEqualizationDistributionModel else '不' }启用均衡分配模式")
        if fastStartHelper == False:
            print(f"找到已启动了 {len(allProcessingWechatPids)} 个微信，请输入您想监听的微信：")
            print(f"默认或其他操作 表示为监听所有的微信进程 ")
        else:
            print(f"找到已启动了 {len(allProcessingWechatPids)} 个微信")
        for wechat_process_index, wechat_process_id in enumerate(
            allProcessingWechatPids
        ):
            print(f"[{wechat_process_index+1}] 号微信 进程号为  {wechat_process_id} ")
        wechat_chose_menu_index = -1
        if fastStartHelper == False:
            try:
                wechat_chose_menu_index = int(input("请输入您选择的微信序号："))
            except:
                print("检测到未选择 单个微信，自动选择所有微信进程")
        if (
            wechat_chose_menu_index <= len(allProcessingWechatPids)
            and wechat_chose_menu_index >= 1
        ):
            print(
                f"开始监听 序号为 {wechat_chose_menu_index} 进程号为 {allProcessingWechatPids[wechat_chose_menu_index - 1]} 的微信 >>> "
            )
            listenWechatPids.append(
                {
                    "index": wechat_chose_menu_index,
                    "wechatPid": allProcessingWechatPids[wechat_chose_menu_index - 1],
                }
            )
        else:
            print("开始监听 所有微信进程 >>> ")
            for wechat_process_index, wechat_process_id in enumerate(
                allProcessingWechatPids
            ):
                listenWechatPids.append(
                    {
                        "index": wechat_process_index + 1,
                        "wechatPid": wechat_process_id,
                    }
                )
    elif len(allProcessingWechatPids) == 1:
        print("开始监听 当前已启动的微信进程 >>> ")
        for wechat_process_index, wechat_process_id in enumerate(
            allProcessingWechatPids
        ):
            listenWechatPids.append(
                {
                    "index": wechat_process_index + 1,
                    "wechatPid": wechat_process_id,
                }
            )

    if len(listenWechatPids) > 0:
        unusedPort = get_unused_ports(1, 20808)
        try:
            if fastStartHelper == False:
                unusedPortRes = int(input("请输入您电脑可用的端口号用于监听微信消息（直接回车则程序随机分配）："))
                if unusedPortRes > 80:
                    unusedPort[0] = unusedPortRes
                    print(f"您指定的端口号为 {unusedPort[0]} >>> ")
            else:
                print(f"随机分配可用端口来与微信进行通信，端口为 {unusedPort[0]} >>> ")
        except:
            print(f"随机分配可用端口来与微信进行通信，端口为 {unusedPort[0]} >>> ")
        # selectedPid = get_available_pids(1, 10807)
        # if len(selectedPid) == 0:
        #     print("未找到可用的进程号，请稍后重试！")
        #     return
        if len(unusedPort) == 0:
            print("未找到可用的端口号，请稍后重试！")
            return
        processId = start_socket_server(
            port=unusedPort[0], request_handler=WechatMessageHandler, main_thread=False
        )
        print(f"启动监听进程完毕，当前监听进程号 为 {processId} ")
        global socketServerThread
        socketServerThread["port"] = unusedPort[0]
        if processId:
            socketServerThread["id"] = processId

        # for listenWechatPid in listenWechatPids:
        #     listenWechat(
        #         listenWechatPid["index"],
        #         listenWechatPid["wechatPid"],
        #     )
        threads = []
        for listenWechatPid in listenWechatPids:
            thread = threading.Thread(
                target=listenWechat,
                args=(
                    listenWechatPid["index"],
                    listenWechatPid["wechatPid"],
                    unusedPort[0],
                ),
            )
            thread.start()
            threads.append(thread)

        print(f"\n======== ▷ 微信自动过检测助手初始化完毕✅ ，如需退出进程，请按 Ctrl + C 即可！ ◁ ========\n")
        # 监听保持所有微信进程的消息处理
        for thread in threads:
            thread.join()


# 杀死所有可能还在运行中的微信监听进程！
def kill_all_robot_processes():
    for process in psutil.process_iter():
        try:
            if (
                process.name() == "CWeChatRobot.exe" or "微信自动过检测助手" in process.name()
            ) and process.pid != os.getpid():
                process.terminate()
                print(f"已杀死进程 {process.name()} 进程ID {process.pid}")
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass


def listenWechat(wechatBotIndex, wechatPid, messagePort):
    try:
        tempWechatBotInstance = {
            "index": wechatBotIndex,
            "pid": wechatPid,
            "instance": None,
        }
        comtypes.CoInitialize()
        robot = comtypes.client.CreateObject("WeChatRobot.CWeChatRobot")
        event = comtypes.client.CreateObject("WeChatRobot.RobotEvent")
        wx = WeChatRobot(wechatPid, robot, event)
        if wx.IsWxLogin() == False:
            print(f" 第[{wechatBotIndex}] 号微信还没登录，监听个锤子，跳过！")
            return
        tempWechatBotInstance["instance"] = wx
        # 注入DLL到微信以启动服务
        installDllRes = wx.StartService()
        print(f"注入DLL到 第[{wechatBotIndex}] 号微信以启动服务：", statusMsg(installDllRes))
        listenMessageRes = wx.StartReceiveMessage(messagePort)
        print(f"启动 第[{wechatBotIndex}] 号微信接收消息Hook完毕 ：", statusMsg(listenMessageRes))
        if "失败" in statusMsg(listenMessageRes):
            print(
                f"启动 第[{wechatBotIndex}] 号微信接收消息Hook失败，请检查是否已经用管理员执行 安装程序.bat，如果执行了还是这样，请尝试手动在软件根目录用管理员权限执行 .\CWeChatRobot.exe /regserver 或者 微信版本不符合！"
            )
            return
        if start_hello(wx, wechatBotIndex) == False:
            print("由于 未获取到 文件传输助手，尝试获取 WxPusher消息推送平台 公众号历史消息二次确认 >>> ")
            wxpusherNewMessages = wx.GetHistoryPublicMsg("wxpusher")
            if wxpusherNewMessages:
                print("获取到的微信pusher的最新十条历史推文成功！")
            else:
                print("获取到的微信pusher的历史推文也失败了，无法判断是否正常监听消息！")
        # print(f"第[{wechatBotIndex}] 号微信机器人正常启动了！")
        # 创建工作对象（处理消息的人）
        dispatcher.add_worker(Worker(tempWechatBotInstance))

        # 公众号列表里找不到
        # officialAccountList = wx.GetOfficialAccountList()
        # if not wx.AddressBook:
        #     wx.GetAddressBook()
        # officialAccountList = [
        #     item
        #     for item in wx.AddressBook
        #     if (item["wxType"] != 2 and item["wxid"][0:3] == "gh_")
        # ]
        # wxpusherOfficialAccount = (
        #     search_obj_list(officialAccountList, "wxNickName", "WxPusher消息推送平台")
        #     or wx.GetFriendByWxNickName("WxPusher消息推送平台")
        #     or wx.GetFriendByWxNumber("wxpusher")
        # )

        # wxpusherOfficialAccountId = ""
        # if wxpusherOfficialAccount:
        #     wxpusherOfficialAccountId = wxpusherOfficialAccount["wxid"]
        #     print("获取到的WxPusher消息推送平台公众号：", wxpusherOfficialAccountId)
        # else:
        #     wxpusherOfficialAccountId = "gh_bf214c93111c"
        wechatRobotProcessList.append(tempWechatBotInstance)
        print(f"第[{wechatBotIndex}] 号微信机器人启动完毕✅ ，开始持续监听WxPusher公众号消息 ▷▷▷ ")
        register_msg_event(wechatPid)
        comtypes.CoUninitialize()
        # print(f"停止服务，将DLL从 第[{wechatBotIndex}] 号微信进程中卸载：", processId)
        # if processId:
        #     stop_socket_server(processId)
        # wx.StopService()
        # print(
        #     f"执行完毕， 第[{wechatBotIndex}] 号微信已经停止监听！",
        # )
    except OSError as e:
        print("未初始化程序，请手动以管理员身份执行下 安装程序.bat ！")
        countdown_and_exit(20)
        return


def get_wechat_process_path():
    # 从当前进程中获取微信的安装路径
    for proc in psutil.process_iter(["pid", "name", "exe"]):
        if proc.info["name"] == "WeChat.exe":
            return proc.info["exe"]
    return None


def get_wechat_from_registry():
    # 从注册表中获取微信的安装路径
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Tencent\WeChat")
    value, _ = winreg.QueryValueEx(key, "InstallPath")
    if value:
        return value + "\WeChat.exe"
    return None


def get_wechat_path_from_shortcut(shortcut_path):
    # 从微信的快捷方式中获取微信的路径
    if os.path.exists(shortcut_path):
        shortcut = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f".lnk\\{shortcut_path}")
        target_id_list = winreg.QueryValueEx(shortcut, "Target")[0]
        shell_link = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            f"Software\\Classes\\lnkfile\\shellex\\ContextMenuHandlers\\ShellLink",
        )
        pidl_data = winreg.QueryValueEx(shell_link, "")[0]

        # 将ID列表转换为文件路径
        import ctypes

        shell = ctypes.windll.shell32
        path_ptr = shell.SHGetPathFromIDListW(ctypes.c_void_p(target_id_list))
        try:
            wechat_path = ctypes.wstring_at(path_ptr)
        finally:
            shell.SHFreeNameMappings(path_ptr)

        return wechat_path
    return None


def find_wechat_path():
    # 尝试从进程、注册表、桌面快捷方式获取微信的路径
    wechat_path = get_wechat_process_path()
    if wechat_path:
        return wechat_path

    wechat_path = get_wechat_from_registry()
    if wechat_path:
        return wechat_path

    # 尝试从桌面查找微信的快捷方式
    desktop_path = os.path.join(os.path.expanduser("&#126;"), "Desktop")
    wechat_shortcuts = [
        f
        for f in os.listdir(desktop_path)
        if f.lower().endswith(".lnk") and ("wechat" in f.lower() or "微信" in f.lower())
    ]
    for shortcut in wechat_shortcuts:
        shortcut_full_path = os.path.join(desktop_path, shortcut)
        wechat_path = get_wechat_path_from_shortcut(shortcut_full_path)
        if wechat_path:
            return wechat_path

    return None


def is_valid_wechat_path(path):
    # 微信标准安装路径的一部分
    standard_wechat_path = r"\WeChat.exe"

    # 检查路径是否以标准微信路径的一部分结尾
    if not path.endswith(standard_wechat_path):
        return False

    # 获取路径中的文件夹部分
    folder_path = os.path.dirname(path)

    # 检查文件夹是否存在
    if not os.path.exists(folder_path):
        return False

    # 检查文件是否存在
    if not os.path.isfile(path):
        return False

    # 如果以上检查都通过，则认为是有效的微信路径
    return True


def launch_wechat(count):
    wechat_path = find_wechat_path()
    if wechat_path:
        for i in range(count):
            subprocess.Popen([wechat_path])
            time.sleep(0.1)  # 等待一段时间，防止操作系统拒绝过快打开多个相同的应用
    else:
        wechatPathInputConfirm = input("无法找到微信路径，很抱歉，请问是否需要手动输入微信路径？请输入 y/n：") == "y"
        if wechatPathInputConfirm:
            wechatPath = input("请输入完整的微信路径：")
            if is_valid_wechat_path(wechatPath):
                for i in range(count):
                    subprocess.Popen([wechatPath])
                    time.sleep(0.1)  # 等待一段时间，防止操作系统拒绝过快打开多个相同的应用


def stopListen():
    print(f"停止服务，将DLL从 {str(len(wechatRobotProcessList))} 个微信进程中卸载 >>> ")
    stopAllWechatBot()
    print(
        "停止服务，执行完毕，请关闭当前窗口！",
    )


def disclaimer():
    print(
        """
程序初始化中……正在开机……欢迎使用 幻生公开版自动过检测机器人  >>>

⚠️ 【免责声明】
------------------------------------------
项目开源地址：https://github.com/Huansheng1/wechat-auto-read-helper
------------------------------------------
0、本项目免费！免费！免费！仅供学习交流和测试，如果你在闲鱼买了，请去退款！！！！！
1、此软件仅用于学习研究，不保证其合法性、准确性、有效性，请根据情况自行判断，本人对此不承担任何保证责任。
2、由于此软件仅用于学习研究，您必须在下载后 24 小时内将所有内容从您的计算机或手机或任何存储设备中完全删除，若违反规定引起任何事件本人对此均不负责。
3、请勿将此软件用于任何商业或非法目的，若违反规定请自行对此负责。
4、此软件涉及应用与本人无关，本人对因此引起的任何隐私泄漏或其他后果不承担任何责任。
5、本人对任何软件引发的问题概不负责，包括但不限于由软件错误引起的任何损失和损害。
6、如果任何单位或个人认为此软件可能涉嫌侵犯其权利，应及时通知并提供身份证明，所有权证明，我们将在收到认证文件确认后删除此软件。
7、所有直接或间接使用、查看此软件的人均应该仔细阅读此声明。本人保留随时更改或补充此声明的权利。一旦您使用或复制了此软件，即视为您已接受此免责声明。
------------------------------------------
    """
    )
    user_input = input("请问你同意以上免责声明吗？(请输入 y 或 n，回车代表同意): ")
    return user_input.lower() != "n"


def join_group():
    webbrowser.open("https://t.me/huan_sheng")


def go_to_publish_page():
    webbrowser.open("https://github.com/Huansheng1/my-qinglong-js")


def openSupportImg():
    img = mpimg.imread(
        os.path.join(os.path.abspath(os.path.dirname(__file__)), "support-author.jpg")
    )
    plt.imshow(img)
    plt.show()


def countdown_and_exit(seconds):
    for i in range(seconds, 0, -1):
        print(f"倒计时 {i} 秒...")
        time.sleep(1)
    print("时间到，程序将退出。")
    os.kill(os.getpid(), signal.SIGTERM)


def entryPoint():
    if not disclaimer():
        print("你拒绝了免责声明，很抱歉，程序退出。")
        return

    while True:
        # 在你的主函数中调用 set_up_logger
        set_up_logger()
        # 打印菜单
        print("\n\n-------------------【程序菜单】-------------------")
        print("0. 一键启动自动过检测助手（默认配置，无需配置）")
        print("1. 自定义启动自动过检测助手（手动选择各种配置）")
        print("2. 多开微信")
        print("3. 打赏作者，加速更新")
        print("4. 加入聊天群")
        print("5. 前往发布页")
        print("q. 退出本程序")
        user_input = str(input("\n请输入您选择的菜单编号(直接回车代表着 一键启动过检测助手)："))
        if user_input == "":
            user_input = "0"
        if user_input in ["1", "0"]:
            # print("【提示】：进入后如想暂停请使用 CTRL + C 终止小助手任务！\n")
            startAutoReadHelper(user_input == "0")
        elif user_input == "2":
            processingWechatCount = len(get_wechat_pid_list())
            if processingWechatCount == 0:
                # 提示用户输入想要启动的微信数量
                number_of_wechats_to_launch = int(input("请输入想要启动的微信数量: ") or 0) or 0
                # 启动指定数量的微信实例
                launch_wechat(number_of_wechats_to_launch + processingWechatCount)
            else:
                print("检测到当前已经有启动的微信，启动第三方现成软件多开微信：")
                # 替换为你的exe文件的路径和参数
                exe_path = "PC微信多开器.exe"
                # 使用subprocess.run运行exe文件
                subprocess.run([exe_path])
        elif user_input == "3":
            openSupportImg()
        elif user_input == "4":
            join_group()
        elif user_input == "5":
            go_to_publish_page()
        elif user_input.lower() == "q":
            print("【提示】：如果卡住了，可以直接关闭本窗口即可！\n")
            countdown_and_exit(10)
            stopListen()
            break
        else:
            print("\n菜单编号无效，请重新输入正确的菜单编号。")
        time.sleep(1)


def stopAllWechatBot():
    for wechatBotIndex, wxRobotChat in wechatRobotProcessList:
        if wxRobotChat is None:
            continue
        print(f"开始关闭对进程ID为 {wxRobotChat.pid} 的微信监听 >>> ")
        if wxRobotChat and wxRobotChat["pid"]:
            comtypes.CoInitialize()
            robot = comtypes.client.CreateObject("WeChatRobot.CWeChatRobot")
            event = comtypes.client.CreateObject("WeChatRobot.RobotEvent")
            wx = WeChatRobot(wxRobotChat["pid"], robot, event)
            wx.StopService()
            comtypes.CoUninitialize()
            # wxRobotChat["instance"].StopService()
        print(
            f"执行完毕， 第[{wechatBotIndex}] 号微信已经停止监听！",
        )
    global socketServerThread
    if socketServerThread and socketServerThread["id"]:
        stop_socket_server(socketServerThread["id"])
    find_and_kill_process_using_port(socketServerThread["port"])
    print(
        f"全部的微信监听已经停止！",
    )


def fetch_config_data(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print("网络异常，获取配置数据失败:", e)
        countdown_and_exit(5)


def parse_version(version_str):
    # 提取版本号中的数字部分
    match = re.search(r"V(\d+\.\d+\.\d+)", version_str)
    if match:
        return match.group(1)
    else:
        return None


def compare_versions(version1, version2):
    def normalize(version):
        # 将版本号字符串拆分成数字部分和非数字部分
        parts = version.split(".")
        normalized_parts = []
        for part in parts:
            if part.isdigit():
                normalized_parts.append(int(part))
            else:
                normalized_parts.append(part)
        return normalized_parts

    # 规范化版本号
    normalized_version1 = normalize(version1)
    normalized_version2 = normalize(version2)

    # 逐位比较版本号
    for v1, v2 in zip(normalized_version1, normalized_version2):
        if v1 < v2:
            return -1
        elif v1 > v2:
            return 1

    # 如果前面的位数都相同，比较长度
    if len(normalized_version1) < len(normalized_version2):
        return -1
    elif len(normalized_version1) > len(normalized_version2):
        return 1
    else:
        return 0


def check_for_updates(local_version, server_version, update_link):
    local_version_num = parse_version(local_version)
    server_version_num = parse_version(server_version)
    compareVersionResult = compare_versions(local_version_num, server_version_num)
    if compareVersionResult < 0:
        print(f"检测到有新版本 {server_version}。")
        user_input = input("是否前往更新？(y/n): ")
        if user_input.lower() == "y":
            webbrowser.open(update_link)
        sys.exit(0)
    elif compareVersionResult >= 0:
        print("当前版本为最新版本。")
    else:
        print("网络异常 或者 无法解析版本号：", local_version, server_version)
        sys.exit(0)


def startAutoReadHelper(fastStartHelper=True):
    print(f"\n======== ▷ 开始启动微信自动过检测助手 ◁ ========\n")
    # print("mian模式 - 当前程序运行目录：", os.getcwd())
    # 杀死某个进程
    # find_and_kill_process_using_port(hook_port)

    # 阻止电脑休眠
    ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)

    # 阻止电脑熄屏
    ctypes.windll.kernel32.SetThreadExecutionState(0x00000002)

    # 创建调度器
    global dispatcher
    dispatcher = Dispatcher([])
    global skipSamePostIn1Second
    skipSelectedResult = True
    if fastStartHelper == False:
        skipSelectedResult = str(input("是否开启 跳过1秒内推送的重复阅读链接（直接回车表示不跳过），请输入 y/n："))
    skipSamePostIn1Second = skipSelectedResult == "y"
    global socketServerThread
    socketServerThread = {"id": None, "port": 10808}
    startListen(fastStartHelper)
    # print("停止机器人 >>> ")
    # stopListen()


# 处理消息的机器人类
class Worker:
    def __init__(self, processRobot):
        # 取微信账号ID作为机器人唯一标识，便于进程重启时恢复记录
        workWechat = processRobot["instance"].GetSelfInfo()
        # print(
        #     f"初始化 机器人中 微信账号 ",
        #     workWechat.get("wxNumber"),
        #     " 微信手机号 ",
        #     workWechat.get("PhoneNumber"),
        #     " 微信昵称 ",
        #     workWechat.get("wxNickName"),
        # )
        self.id = (
            str(workWechat.get("wxNumber"))
            or str(workWechat.get("wxId"))
            or str(uuid.uuid4())
        )
        self.processRobot = processRobot
        self.busy_until = 0
        self.enabled = True
        self.wxNickName = str(processRobot["instance"].GetSelfInfo().get("wxNickName"))

    def process(self, decodePostUrl):
        if not self.enabled:
            print(f"{self.processRobot['index']}号机器人[{self.wxNickName}] 已经被禁用，无法阅读，跳过")
            return False
        try:
            print(f"{self.processRobot['index']}号机器人[{self.wxNickName}] 正在阅读文章")
            comtypes.CoInitialize()
            robot = comtypes.client.CreateObject("WeChatRobot.CWeChatRobot")
            event = comtypes.client.CreateObject("WeChatRobot.RobotEvent")
            wx = WeChatRobot(self.processRobot["pid"], robot, event)
            if wx.IsWxLogin():
                wx.OpenBrowser(decodePostUrl)
                comtypes.CoUninitialize()
            else:
                print(
                    f"{self.processRobot['index']}号机器人[{self.wxNickName}] 已掉线，无法执行任务，开始禁用 ▷▷▷ "
                )
                self.disable()
                comtypes.CoUninitialize()
                return False
            # 直接调用是不行的，因为跨线程调用会报错：失败: (-2147417842, '应用程序调用一个已为另一线程整理的接口。'
            # self.processRobot["instance"].OpenBrowser(decodePostUrl)
            processing_time = random.uniform(4, 6)
            with lock:
                self.busy_until = time.time() + processing_time
            time.sleep(processing_time)
            print(f"{self.processRobot['index']}号机器人[{self.wxNickName}] 已阅读完毕")
            return True
        except Exception as e:
            print(
                f"{self.processRobot['index']}号机器人[{self.wxNickName}] 阅读文章 {decodePostUrl} 失败: {e}"
            )
            return False

    def disable(self):
        self.enabled = False

    def enable(self):
        self.enabled = True


# 分发任务的派送类
class Dispatcher:
    def __init__(self, workers, max_queue_wait_time=8):
        self.workers = workers
        self.queue = queue.Queue()
        self.worker_records = {}
        self.records_file = "records.pkl"
        self.max_queue_wait_time = max_queue_wait_time
        if os.path.exists(self.records_file):
            with open(self.records_file, "rb") as f:
                self.worker_records = pickle.load(f)

    def add_item(self, postUrl):
        self.queue.put((postUrl, time.time()))
        print(f"添加文章到队列成功，当前添加的文章链接为 {postUrl}，当前队列长度为 {self.queue.qsize()}")

    def dispatch(self):
        print(f"开始派发任务，当前任务数量 {self.queue.qsize()}")
        while not self.queue.empty():
            postUrl, added_time = self.queue.get()
            if time.time() - added_time > self.max_queue_wait_time:
                print(
                    f"文章 {postUrl} 已超过消息处理最长等待时间 {self.max_queue_wait_time}秒，已过期，将其移除队列 >>> "
                )
                continue
            sorted_workers = self.get_sorted_workers(postUrl)
            # print("获取到的可用机器人列表：", sorted_workers)
            if (len(sorted_workers) > 0) and (type(sorted_workers) == list):
                print(
                    f"获取到可用的机器人 ",
                    "、".join([item.wxNickName for item in sorted_workers]) or "无",
                )
            else:
                print("当前无可用的机器人，等待是否有可用的机器人")
                self.queue.put((postUrl, added_time))
                time.sleep(0.1)
            if len(sorted_workers) > 0:
                worker = sorted_workers[0]
                print(f"分配任务队列给机器人中 工作微信 {worker.wxNickName} 处理文章为 {postUrl}")
                print(
                    f"{worker.processRobot['index']}号机器人[{worker.wxNickName}] 被分派文章 {postUrl}，今日已阅读该链接 {self.worker_records.get(worker.id, {}).get('count', {}).get(postUrl, 0)}次，今日总共过检测 {self.worker_records.get(worker.id, {}).get('total',  0)}次"
                )
                if worker.process(postUrl):
                    self.record(worker, postUrl)
            self.queue.task_done()

    def record(self, worker, postUrl):
        # 记录机器人处理的消息
        today = datetime.date.today()
        if (
            worker.id not in self.worker_records
            or self.worker_records[worker.id]["date"] != today
        ):
            # 如果机器人的记录不存在，或者记录的日期不是今天，那么创建一个新的记录
            self.worker_records[worker.id] = {"date": today, "count": {}, "total": 0}
        # 如果记录的日期是今天，那么增加计数
        self.worker_records[worker.id]["count"][postUrl] = (
            self.worker_records[worker.id]["count"].get(postUrl, 0) + 1
        )
        self.worker_records[worker.id]["total"] += 1
        # 删除非今日的数据
        self.worker_records = {
            k: v for k, v in self.worker_records.items() if v["date"] == today
        }
        # 持久化记录
        with open(self.records_file, "wb") as f:
            pickle.dump(self.worker_records, f)

    def get_sorted_workers(self, postUrl):
        # 先按照 busy_until 属性排序
        self.workers.sort(key=lambda x: x.busy_until)
        # 然后在可用的机器人中，优先选择处理该消息里的链接次数最少的机器人
        available_workers = [
            worker
            for worker in self.workers
            if worker.busy_until <= time.time() and worker.enabled
        ]
        try:
            available_workers.sort(
                key=lambda x: (
                    self.worker_records.get(x.id, {}).get("count", {}).get(postUrl, 0),
                    self.worker_records.get(x.id, {}).get("total", 0),
                )
            )
        except Exception as e:
            print(f"检测到老版本记录数据，抹除原有记录")
            self.worker_records = {}
            with open(self.records_file, "wb") as f:
                pickle.dump(self.worker_records, f)
            return []
        return available_workers

    def add_worker(self, worker):
        self.workers.append(worker)

    def remove_worker(self, worker_id):
        self.workers = [worker for worker in self.workers if worker.id != worker_id]
