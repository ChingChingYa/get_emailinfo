# https://blog.csdn.net/u010103202/article/details/78045523
# https://blog.csdn.net/tcl415829566/article/details/78481932
# https://zhuanlan.zhihu.com/p/77682457
import poplib
import datetime
import time
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from bs4 import BeautifulSoup


class Email():
    def __init__(self, From, To, Subject, Date, finaltext):
        self.From = From
        self.To = To
        self.Subject = Subject
        self.Date = Date
        self.finaltext = finaltext


# 取得信件資訊(From, To, Subject, Date)
def get_info(msg, keysubject):
    email = Email("", "", "", "", "")
    for header in ['From', 'To', 'Subject', 'Date']:
        value = msg.get(header, '')
        if value:
            if header == 'Subject':
                value = decode_str(value)
                email.Subject = value
                # 判斷主旨
                if (keysubject in value):
                    continue
                else:
                    return
            elif header == 'Date':
                value = decode_str(value)
                email.Date = value
            elif header == 'From':
                hdr, addr = parseaddr(value)
                name = decode_str(hdr)
                value = '%s <%s>' % (name, addr)
                email.From = value
            else:
                hdr, addr = parseaddr(value)
                name = decode_str(hdr)
                value = '%s <%s>' % (name, addr)
                email.To = value

    return email


# 取得信件內容
def context_info(msg, keyword, indent=1):
    if (msg.is_multipart()):
        parts = msg.get_payload()
        # list = []
        for n, part in enumerate(parts):
            print('%spart %s' % ('  ' * indent, n))
            print('%s--------------------' % ('  ' * indent))
            context = context_info(part, keyword, indent + 1)
            # list.append(context)
    else:
        # list = []
        content_type = msg.get_content_type()
        if content_type == 'text/plain' or content_type == 'text/html':
            content = msg.get_payload(decode=True)
            charset = guess_charset(msg)
            if charset == '"gb2312"':
                charset = '"gb18030"'
            if charset:
                content = content.decode(charset)

            soup = BeautifulSoup(content, "html.parser")
            text = soup.get_text().strip().split("Best\xa0 regards,")
            text = text[0].split("This message contains")
            # 判斷信件內容關鍵字
            if (keyword in text[0]):
                context = text[0]
                # print("信件內容:"+text[0])
                
            else:
                return
        # else:
        #     print('%sAttachment: %s' % ('  ' * indent, content_type))
    return context


def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset == 'gb2312':
        charset = 'gb18030'
    if charset:
        value = value.decode(charset)
    return value


def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset


if __name__ == '__main__': 
    # 输入邮件地址, 口令和POP3服务器地址:
    # email = input('Email: ')
    # password = input('Password: ')
    # pop3_server = input('POP3 server: ')
    # 日期赋值
    day = datetime.date.today()
    str_day = str(day).replace('-', '')
    pop3_server = 'casarray.systex.tw'
    email = '1900782@systex.com.tw'
    password = 'a5a1a1a3a2'
    keysubject = input('請輸入指定特定主旨: ')
    keyword = input('請輸入信件內容關鍵字: ')

    
    # 连接到POP3服务器:
    server = poplib.POP3_SSL(pop3_server)
    # 可以打开或关闭调试信息:
    server.set_debuglevel(1)
    # 可选:打印POP3服务器的欢迎文字:
    # print(server.getwelcome().decode('utf-8'))
    
    # 身份认证:
    server.user(email)
    server.pass_(password)
    
    # stat()返回邮件数量和占用空间:
    # print('Messages: %s. Size: %s' % server.stat())

    # list()返回所有邮件的编号:
    resp, mails, octets = server.list()
    # 可以查看返回的列表类似[b'1 82923', b'2 2184', ...]
    # print(mails)
    
    # 获取最新一封邮件, 注意索引号从1开始:
    index = len(mails)
    finalresult = []
    for i in reversed(range(index)):
        resp, lines, octets = server.retr(i+1)

        # lines存储了邮件的原始文本的每一行,
        # 可以获得整个邮件的原始文本:
        msg_content = b'\r\n'.join(lines).decode('utf-8')
        # 稍后解析出邮件:
        msg = Parser().parsestr(msg_content)
        date1 = time.strptime(msg.get("Date")[0:24], '%a, %d %b %Y %H:%M:%S')
        date2 = time.strftime("%Y%m%d", date1)
        # 設定為3日內信件
        if abs(int(date2) - int(str_day)) <= 3:
            email_info = get_info(msg, keysubject)
            if email_info is None:
                continue
            email_info.finaltext = context_info(msg, keyword)
            if email_info.finaltext is None:
                continue
            finalresult.append(email_info)
        else:
            break
    # 可以根据邮件索引号直接从服务器删除邮件:
    # server.dele(index)
    # 关闭连接:
    server.quit()

    # 輸出到txt
    for i, info in enumerate(finalresult):
        i_str = str(i+1)
        file_name = i_str + '.txt'
        f = open(file_name, 'w', encoding="utf-8")
        allcontext = ""
        for j in info.finaltext:
            allcontext += j
        f.write("信件主旨:"+info.Subject+"\n信件內容:"+allcontext)
        f.close()
