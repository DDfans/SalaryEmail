#coding=utf-8

import threading
import datetime
import base64

import tkinter as tk
from tkinter import ttk
from .db_instance import set_db, SalaryEmail
from .setting_box import AccountPasswordWin, SMTPPortWin, InfoWin, SysSettingWin
from .parse_execl import ParseExcel
from .send_email import SendEmail
import tkinter.messagebox
import tkinter.filedialog
import os

DEFAULT_COUNT = 4

class MainWin(tk.Tk):

    def __init__(self):
        super(MainWin, self).__init__()
        self.title('开心发工资')

        x = (self.winfo_screenwidth() // 2) - 300
        y = (self.winfo_screenheight() // 2) - 300
        self.geometry('600x600+{}+{}'.format(x, y))
        self.resizable(width=False, height=False)  # 禁制拉伸大小
        self.label_width = 55  # 标签长度

        self.db = set_db()
        self.subject = tk.StringVar() # 邮件标题
        self.mail_content = tk.StringVar()
        self.salary_file_path = tk.StringVar()
        self.send_date = tk.StringVar() # 发件时间
        self.sender_text = tk.StringVar() # 发件邮箱
        self.sender_name_text = tk.StringVar() #发件人
        self.sign_text = tk.StringVar()
        self.smtp_text = tk.StringVar() #发件服务器地址
        self.port_text = tk.StringVar() #发件端口
        self.__password = tk.StringVar() # base64邮件密码

        self.done_count = 0  # 完成的邮件数量
        self.P_count = tk.IntVar()  # 进度条量
        self.lock = threading.Lock()  # 计数增量线程锁
        self.gen_lock = threading.Lock()  # 行数据生成器线程锁

        self.thread_count = tk.IntVar() # 发送邮件进程数

        self.show_percent = tk.StringVar()  # 显示百分百
        self.show_percent.set('完成百分比：0%')

        self.show_percent_th = threading.Thread(target=self.show_percent_run)

        self.excel_file = None
        self.setupUI()
        self.set_default_info()

    def setupUI(self):
        '''主界面'''

        # 设置菜单栏
        menubar = self.set_menubar()
        self.config(menu=menubar)
        self.show_base_info()

    def show_account_box(self):
        account_box = AccountPasswordWin(parent=self)
        self.wait_window(account_box)

    def show_smtp_port_box(self):
        smtp_box = SMTPPortWin(parent=self)
        self.wait_window(smtp_box)

    def show_info_box(self):
        info_box = InfoWin(parent=self)
        self.wait_window(info_box)

    def set_menubar(self):
        '''菜单栏'''
        menubar = tk.Menu(self)
        filemenu = tk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="文件", menu=filemenu)

        filemenu.add_command(label='退出', command=self.quit)
        settingmenu = tk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="设置", menu=settingmenu)
        settingmenu.add_command(label="账号/密码", command=self.show_account_box)
        settingmenu.add_command(label="SMTP域名/端口", command=self.show_smtp_port_box)
        settingmenu.add_command(label="邮件信息设置", command=self.show_info_box)
        # settingmenu.add_command(label="系统设置", command=self.show_sys_setting_box)
        return menubar

    def show_sys_setting_box(self):
        '''显示系统设置菜单'''
        sys_setting_box = SysSettingWin(parent=self)
        self.wait_window(sys_setting_box)

    def show_base_info(self):
        '''显示基本信息'''

        row1 = tk.Frame(self)
        row1.pack(fill='x', padx=1, pady=5)
        tk.Label(row1, text="发件邮箱：", width=15).pack(side=tk.LEFT)
        tk.Label(row1, textvariable=self.sender_text, width=25, justify=tk.CENTER).pack(side=tk.LEFT)
        tk.Label(row1, text="发件人：", width=10).pack(side=tk.LEFT)
        tk.Label(row1, textvariable=self.sender_name_text, width=25, justify=tk.CENTER).pack(side=tk.LEFT)

        row7 = tk.Frame(self)
        row7.pack(fill='x', padx=1, pady=5)
        tk.Label(row7, text="SMTP服务器：", width=15).pack(side=tk.LEFT)
        tk.Label(row7, textvariable=self.smtp_text, width=25, justify=tk.CENTER).pack(side=tk.LEFT)
        tk.Label(row7, text="PORT：", width=10).pack(side=tk.LEFT)
        tk.Label(row7, textvariable=self.port_text, width=25, justify=tk.CENTER).pack(side=tk.LEFT)

        # row2 = tk.Frame(self)
        # row2.pack(fill='x', padx=1, pady=5)
        # tk.Label(row2, text="发件人：", width=15).pack(side=tk.LEFT)
        # tk.Label(row2, textvariable=self.sender_name_text, width=self.label_width).pack(side=tk.LEFT)

        row4 = tk.Frame(self)
        row4.pack(fill='x', padx=1, pady=5)
        tk.Label(row4, text="工资条文件：", width=15).pack(side=tk.LEFT)
        tk.Entry(row4, textvariable=self.salary_file_path, width=42, justify=tk.CENTER).pack(side=tk.LEFT)
        tk.Button(row4, text='选择文件', command=self.get_salary_file_path, width=10).pack(side=tk.LEFT)

        row3 = tk.Frame(self)
        row3.pack(fill='x', padx=1, pady=5)
        tk.Label(row3, text="邮件标题：", width=15).pack(side=tk.LEFT)
        tk.Entry(row3, textvariable=self.subject, width=self.label_width, justify=tk.CENTER).pack(side=tk.LEFT)

        row8 = tk.Frame(self)
        row8.pack(fill='x', padx=1, pady=5)
        tk.Label(row8, text="邮件内容：", width=15).pack(side=tk.LEFT)
        self.mail_text = tk.Text(row8, width=63, height=5, relief="solid")
        self.mail_text.pack(side=tk.LEFT)
        self.mail_text.insert(tk.END,"您好，附件为  月工资条，请查收。\n由于薪酬保密，此邮件内容请勿向他人透露，如有问题请回复此邮件或与人力资源部联系")

        # row5 = tk.Frame(self)
        # row5.pack(fill='x', padx=1, pady=5)
        # tk.Label(row5, text="邮件签名/落款：", width=15).pack(side=tk.LEFT)
        # tk.Label(row5, textvariable=self.sign_text, width=self.label_width).pack(side=tk.LEFT)

        # row6 = tk.Frame(self)
        # row6.pack(fill='x', padx=1, pady=5)
        # tk.Label(row6, text="邮件日期：", width=15).pack(side=tk.LEFT)
        # tk.Entry(row6, textvariable=self.send_date, width=self.label_width, justify=tk.CENTER).pack(side=tk.LEFT)





        tk.Button(self, command=self.send_email, text='发送', width=20).pack(padx=1, pady=5)
        tk.Label(self, textvariable=self.show_percent ).pack(padx=1, pady=5)

        # 进度条
        row8 = tk.Frame(self, padx=20, pady=1)
        row8.pack(fill='x', padx=1, pady=5)
        self.progressbar = ttk.Progressbar(row8, orient='horizontal', length=545, mode="determinate", variable=self.P_count)
        self.progressbar.pack(side=tk.LEFT)



        # 发送结果显示框
        self.result_box = tk.Frame(self, borderwidth=1, padx=20, pady=0)
        self.result_box.pack(fill='x',padx=1, pady=5)
        self.result_list = ttk.Treeview(self.result_box, height=12, show="headings")
        self.result_list['columns'] = ('姓名', '员工号',"邮箱", '发送结果')
        self.result_list.column('姓名', width=90, anchor=tk.CENTER)  # 表示列,不显示
        self.result_list.column('员工号', width=90, anchor=tk.CENTER)  # 表示列,不显示
        self.result_list.column("邮箱", width=240, anchor=tk.CENTER)
        self.result_list.column('发送结果', width=120, anchor=tk.CENTER)
        self.result_list.heading('姓名', text="姓名")
        self.result_list.heading('员工号', text="员工号")
        self.result_list.heading("邮箱", text="邮箱")
        self.result_list.heading('发送结果', text='发送结果')
        self.result_list.grid(row=0, column=0, sticky=tk.NSEW)
        vbar = ttk.Scrollbar(self.result_box, orient=tk.VERTICAL, command=self.result_list.yview)
        self.result_list.configure(yscrollcommand=vbar.set)
        vbar.grid(row=0, column=1, sticky=tk.NS)

    def count_done_row(self):
        '''计算完成的任务数'''
        # with self.lock:
        self.done_count += 1
        self.P_count.set(self.done_count)

    def show_percent_run(self):
        total_count = self.excel_file.avaRows
        current_count = self.done_count
        percent = "%0.1f" % (current_count/float(total_count) * 100)
        self.show_percent.set("完成百分百：{}%".format(percent))

    def get_salary_file_path(self):
        '''获取工资条文件路径'''
        path = tk.filedialog.askopenfilename(title='选择文件', filetypes=[("Excel File", "*.xls *.xlsx")])
        self.salary_file_path.set(path)

    def set_default_info(self):
        '''设置默认初始值'''
        self.subject.set('工资条')
        try:
            sender = self.db.session.query(SalaryEmail).filter(SalaryEmail.field_name=='sender').first()
            sender_text = sender.field_value if sender else ""
            sender_name = self.db.session.query(SalaryEmail).filter(SalaryEmail.field_name=='sender_name').first()
            sender_name_text = sender_name.field_value if sender_name else ""
            sign = self.db.session.query(SalaryEmail).filter(SalaryEmail.field_name=='sign').first()
            sign_text = sign.field_value if sign else ""
            smtp = self.db.session.query(SalaryEmail).filter(SalaryEmail.field_name=='smtp_server').first()
            smtp_text = smtp.field_value if smtp else ""
            port = self.db.session.query(SalaryEmail).filter(SalaryEmail.field_name=='port').first()
            port_text = port.field_value if port else ""
            thread = self.db.session.query(SalaryEmail).filter(SalaryEmail.field_name=='thread_count').first()
            thread_count = int(thread.field_value) if thread else DEFAULT_COUNT

        except Exception as e:
            tk.messagebox.showerror(title='错误', message='数据库错误！\n{}'.format(e))
            sender_text = ''
            sender_name_text = ''
            sign_text = ''
            smtp_text = ''
            port_text = ''
            thread_count = DEFAULT_COUNT
        self.sender_text.set(sender_text)
        self.sender_name_text.set(sender_name_text)
        self.sign_text.set(sign_text)
        self.send_date.set(datetime.datetime.now().strftime("%Y-%m-%d"))
        self.smtp_text.set(smtp_text)
        self.port_text.set(port_text)
        self.thread_count.set(thread_count)

    def send_email(self):
        '''发送邮件'''
        try:
            file_name = self.salary_file_path.get()
            if file_name.rsplit('.', 1)[1].lower() not in ('xlsx', 'xls'):
                tk.messagebox.showerror(title='文件错误', message='请选择正确的excel文件！')
                return
            self.excel_file = ParseExcel(file_name=file_name)
            # 初始化进度条
            self.progressbar['maximum'] = self.excel_file.avaRows
            self.P_count.set(0)
        except Exception as e:
            tk.messagebox.showerror(title='文件错误', message='请选择正确的excel文件！\n{}'.format(e))
            return
        try:
            password = self.db.session.query(SalaryEmail).filter(SalaryEmail.field_name == 'password').first()
            self.__password.set(base64.decodebytes(password.field_value).decode('utf-8'))
        except Exception as e:
            tk.messagebox.showerror(title='错误', message='数据库错误！\n{}'.format(e))
            return

        self.done_count = 0  # 重置计数
        self.mail_content = self.mail_text.get("0.0","end");

        root_path = os.path.dirname(file_name)
        excel_path = root_path + '/' + self.send_date.get()
        if not os.path.exists(excel_path):
            os.makedirs(excel_path)


        # ob = SendEmail(self, self.__password.get(), self.excel_file.sheetTitle, self.excel_file.allHeaders, self.excel_file.allUserData,excel_path)
        # ob.run()

        # gen = self.excel_file.iter_salary_line()
        # thread_count = 1
        # for i in range(thread_count):
        #     print(i)
        send_thread = threading.Thread(target=self._send_email, args=(self.excel_file.sheetTitle,self.excel_file.allHeaders,self.excel_file.allUserData, excel_path))  # 子线程发送邮件
        send_thread.setDaemon(True)
        send_thread.start()

    def _send_email(self, t,h,u,p):
        ob = SendEmail(self, self.__password.get(),t,h,u,p)
        ob.run()

    @staticmethod
    def _get_year_month():
        today = datetime.datetime.now()
        year = today.year
        month = today.month
        if month == 1:
            year -= 1
            month = 12
        else:
            month -= 1
        return year, month

    def get_center(self):
        px = self.winfo_x()
        py = self.winfo_y()
        pw = self.winfo_width()
        ph = self.winfo_height()

        return (int(px + pw/2), int(py + ph/2))


if __name__ == '__main__':
    win = MainWin()
    win.mainloop()
