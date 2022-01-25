import os,telnetlib,poplib,traceback,re,smtplib,shutil,cv2,time
from PIL import Image
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.header import Header
Image.MAX_IMAGE_PIXELS = None

# 登陆邮箱检查邮件列表
def email_login(email_user,password,pop3_server):
        global mail_id_list_new # 使用全局变量：邮件ID列表
        telnetlib.Telnet(pop3_server,995) # 连接到POP3服务器，进行身份认证
        server=poplib.POP3_SSL(pop3_server,995,timeout=10)
        server.user(email_user) # 发送用户名
        server.pass_(password) # 发送密码
        print('\n邮件总数: %s. 大小: %s'%server.stat()) # 返回邮件数量和占用空间
        resp,mails,octets = server.list() # list()返回所有邮件的编号
        #print('邮件编号及大小：'+str(mails))
        for i in range(0,len(mails)): # 将读取到的邮件ID进行处理，并存入邮件ID列表
            mailid = mail_list_text_process(str(mails[i]))
            mail_id_list_new = mail_id_list_new + [mailid]
        return server

# 登陆邮箱但不检查邮件列表
def email_connect(email_user,password,pop3_server):
    telnetlib.Telnet(pop3_server,995) # 连接到POP3服务器，进行身份认证
    server=poplib.POP3_SSL(pop3_server,995,timeout=10)
    server.user(email_user) # 发送用户名
    server.pass_(password) # 发送密码
    return server

def email_analyze(msg_content): # 解析邮件
    msg = Parser().parsestr(msg_content)
    From = parseaddr(msg.get('from'))[1] # 发件人（用户）
    To = parseaddr(msg.get('To'))[1]
    Subject = decode_str(parseaddr(msg.get('Subject'))[1]) # 主题
    return msg,From,To,Subject

def image_download(msg_content): # 下载邮件附件
    global attachment_path_list,PATH_recv,From,attachment_name_list # 使用全局变量：接受到的文件的路径列表，存储文件夹路径，发件人（用户）
    msg = Parser().parsestr(msg_content)
    for part in msg.walk():
        try: # 增强稳定性
            file_name = part.get_filename() # 获取文件名
            if file_name:
                file_name = str(From)+'！'+decode_str(file_name) # 重命名下载的文件
                attachment_name_list.append(file_name) # 将该文件的名称添加到文件名列表
                data = part.get_payload(decode=True)
                att_file = open(os.path.join(PATH_recv,file_name), 'wb') # 合成完整文件路径
                att_file.write(data) # 写入数据
                att_file.close() # 关闭文件
                attachment_path_list.append(os.path.join(PATH_recv,file_name)) # 将该文件的路径添加到路径列表
                print(f'附件 {file_name} 下载完成，位于：'+str(PATH_recv))
        except:
            print(f'附件下载出错！')
            write_id_txt()

def mail_send(subject_content,body_content,image_path_list,image_name_list): # 发送邮件（主题，正文，附件路径列表）
    global email_user,password,pop3_server,From 
    mm = MIMEMultipart('related')
    mm['From'] = 'Kokomi<' + email_user + '>' # 设置发送者
    mm['To'] = '<'+From+'>' # 设置接受者
    mm['Subject'] = Header(subject_content,'utf-8') # 设置邮件主题
    message_text = MIMEText(body_content,'plain','utf-8') # 构造正文文本
    mm.attach(message_text) # 添加正文文本
    for pic_number in range(0,len(image_path_list)): # 依次添加附件
        file_path = image_path_list[pic_number] # 从列表中提取某个附件路径
        file_name = image_name_list[pic_number] # 从列表中提取某个附件名称
        #print(file_path)
        #print(file_name)
        imageApart = MIMEImage(open(file_path, 'rb').read(), file_path.split('.')[-1])
        imageApart.add_header('Content-Disposition', 'attachment', filename=file_name)
        mm.attach(imageApart) # 添加附件
    stp = smtplib.SMTP() # 创建SMTP对象
    stp.connect(pop3_server, 25) # 设置发件人邮箱的域名和端口
    stp.set_debuglevel(0) # set_debuglevel(1)可以打印出和SMTP服务器交互的所有信息
    stp.login(email_user,password) # 登录邮箱
    stp.sendmail(email_user, From, mm.as_string()) # 发送邮件
    print('邮件发送成功')
    stp.quit() # 关闭SMTP对象

def pic_process(mode): # 图像超分辨率处理（需使用模型编号），处理后将原图移动到归档文件夹
    global PATH_shell
    model_list = ['Real-ESRGAN.sh','Real-ESRGAN-anime.sh','RealSR.sh'] # 可用超分辨率模型的列表
    print('本次放大模型：' + model_list[mode])
    sr = os.popen(os.path.join(PATH_shell,model_list[mode])) # 使用模型对应的的shell
    sr.read()
    remove_file_folder(PATH_recv,PATH_arch) # 完成后，将待处理文件夹中的图片移动至归档文件夹

def pic_compression(imageFile_path,image_name,quality): # 图像压缩
    img=Image.open(imageFile_path)
    imageFile_name_list = imageFile_path
    imageFile_name_list = imageFile_name_list.split(".")
    image_name_list = image_name.split(".")
    if imageFile_name_list[-1] == "png":
        imageFile_name_list[-1] = "jpg"
        image_name_list[-1] = "jpg"
        imageFile_name_list = str.join(".", imageFile_name_list)
        image_name_list = str.join(".", image_name_list)
        try:
            r,g,b,a=img.split()              
            img=Image.merge("RGB",(r,g,b))  
        except:
            print('无alpha通道')
        print(str(imageFile_name_list))
        imageFile_path = imageFile_name_list
        image_name = image_name_list
        img.save(imageFile_path, quality=quality)
        print('图片过大，已转化'+imageFile_path)
        return imageFile_path, image_name

    if imageFile_name_list[-1] == "jpg":
        img.save(imageFile_path, quality=quality)
        print('图片过大，已转化'+imageFile_path)
        return imageFile_path, image_name


def remove_file_folder(old_path, new_path): # 移动文件夹中的文件
    filelist = os.listdir(old_path) # 列出该目录下的所有文件,listdir返回的文件列表是不包含路径的。
    for file in filelist:
        src = os.path.join(old_path, file)
        dst = os.path.join(new_path, file)
        print('src:', src)
        print('dst:', dst)
        shutil.move(src, dst)

def remove_file(file_name, new_path): # 移动文件
    try:
        shutil.move(file_name, new_path) # 将该文件移至归档文件夹
    except:
        os.remove(file_name)
        print('已存在相同名称的文件，新文件被删除:'+file_name)


def mail_list_text_process(text): # 处理邮件ID的格式
    #text = str(re.findall(r' (.+?)\'',text)) # 去除头尾
    text = str(re.findall(r'\'(.+?) ',text))
    text = text.replace('[','') # 去除[号
    text = text.replace(']','') # 去除]号
    text = text.replace('\'','') # 去除\号
    return text

def to_png(name): # 处理文件后缀名
    name_list = name.split(".")
    name_list[-1] = "png"
    name = str.join(".", name_list)
    #name = name.replace(".jpg", ".png")
    #name = name.replace(".JPG", ".png")
    #name = name.replace(".jpeg", ".png")    
    #name = name.replace(".JPEG", ".png")
    #name = name.replace(".PNG", ".png")
    return name

def write_id_txt(): # 保存已处理邮件ID列表到文本文件
    global mail_id_list_old,IDPATH # 全局变量
    file = open(IDPATH, 'w') # 打开ID列表文件
    file.write(str(mail_id_list_old)) # 将已处理ID列表写入
    file.close() # 将文件关闭



def read_id_text():
    global mail_id_list_old,IDPATH
    f = open(IDPATH,'r') # 设置文件对象，存储已处理邮件ID列表的文本文件
    mail_id_list_old_text = f.read() # 将txt文件的所有内容读入到字符串str中
    f.close() # 将文件关闭
    mail_id_list_old = re.findall(r'\'(.+?)\'', mail_id_list_old_text) # 将文本形式的全部ID信息切片为列表

def decode_str(str_in): # 字符编码转换
    value, charset = decode_header(str_in)[0]
    if charset:
        value = value.decode(charset)
    return value
   
def vip_verify(email): # 检查发件人是否为VIP
    if_vip = False
    f = open(VIPPATH,'r') # 设置文件对象
    VIPlist_text = f.read() # 将txt文件的所有内容读入到字符串VIPlist_text中
    f.close() # 将文件关闭
    VIPlist = re.findall(r'\'(.+?)\'', VIPlist_text) # 将文本形式的全部VIP邮箱切片为列表
    if (email in VIPlist):
        if_vip = True
    return if_vip

def blacklist_verify(email): # 检查发件人是否在黑名单中
    if_blacklist = False
    f = open(BLACKLISTPATH,'r') # 设置文件对象
    blacklist_text = f.read() # 将txt文件的所有内容读入到字符串VIPlist_text中
    f.close() # 将文件关闭
    blaclist = re.findall(r'\'(.+?)\'', blacklist_text) # 将文本形式的全部VIP邮箱切片为列表
    if (email in blaclist):
        if_blacklist = True
    return if_blacklist

def pixel_verify(file_name): # 检查是否超过分辨率上限,True为未超过上限
    global PATH_recv, pixel_limit, pixel_limit_mandatory, if_max, if_pixel_error
    try:
        img = cv2.imread(os.path.join(PATH_recv,file_name))
        sp = img.shape
        height = sp[0]
        width = sp[1]
        px = height * width
        print('图片名：' + file_name + '；图片分辨率：' + str(px))
        if px > pixel_limit :
            if if_max == True:
                if px > pixel_limit_mandatory:
                    if_not_too_large = False
                    if_pixel_error = True
                    print('图片分辨率超过强制上限限制')
                else:                    
                    if_not_too_large = True
            else:
                if_not_too_large = False
                print('图片分辨率过高')
                if_pixel_error = True
        else:
            if_not_too_large = True
    except:
        if_not_too_large = False
    return if_not_too_large, if_pixel_error

def if_suffix(path): # 检查是否为图片文件
    global suffix_list, if_suffix_error
    suffix_text_list = path.split(".")
    suffix = suffix_text_list[-1]
    if suffix in suffix_list:
        suffix_available = True
    else:
        suffix_available = False
        if_suffix_error = True
    return suffix_available, if_suffix_error

version = 'v2.3' # 版本号
times = 3600 # 循环次数
wait_time = 10 # 每一轮循环的等待时间
pixel_limit = 4000000 # 限制最高像素数
pixel_limit_mandatory = 25000000 # 强制限制最高像素数
suffix_list =['jpg','JPG','Jpg','jpeg','JPEG','Jpeg','png','PNG','Png'] # 可处理的文件后缀

email_user = 'xxxxxxxxxx@qq.com' # 邮箱地址
password = 'xxxxxxxxxxxxxx' # 密码/口令
pop3_server = 'pop.qq.com' # 服务器地址

print('\n\033[1;31memailSR ' + version + '\033[0m by Youngwang\n')
print('设定为执行' + str(times) + '次')
print('执行间隔时间：' + str(wait_time) + '秒')
print('当前服务提供邮箱：' + email_user)

# 文件路径设定（多平台兼容性未验证）
ROOT_PATH = os.path.abspath(os.path.dirname(__file__)) # 当前文件所在路径
PATH_recv = os.path.join(ROOT_PATH,'Pic_recv','') # 存储接受邮件的路径
PATH_send = os.path.join(ROOT_PATH,'Pic_send','') # 存储发送图片的路径
PATH_arch = os.path.join(ROOT_PATH,'Pic_arch','') # 存储原图归档的路径
PATH_shell = os.path.join(ROOT_PATH,'shell','') # 存储shell的路径
pathlist = [ROOT_PATH] + [PATH_recv] + [PATH_send] + [PATH_send] # 为创建文件夹提供列表
for i in range(0, len(pathlist)):
    if not os.path.exists(pathlist[i]): # 如果文件夹不存在时，创建文件夹
        os.mkdir(pathlist[i])
IDPATH = os.path.join(ROOT_PATH,'ID.txt') # 存储已处理邮件ID列表的文本文件的路径
VIPPATH = os.path.join(ROOT_PATH,'VIP-list.txt') # 存储VIP邮箱列表文件的路径
BLACKLISTPATH = os.path.join(ROOT_PATH,'blacklist.txt') # 存储黑名单邮箱列表文件的路径

# 主要部分
for counter in range(1, times + 1): # 循环执行times次
    print('\n\033[1;31m执行第'+str(counter)+'次循环\033[0m')
    try: # 防止错误终止进程
        mail_id_list_new = []
        mail_id_list_process = []
        mail_id_list_old = []
        mail_id_list_old_text = []
        unprocessed_email_number = 0 # 清空未处理邮件数计数器
        mails_send = 0 # 清空已发送邮件数计数器

        read_id_text() # 将文件中的邮件ID读取至列表中
        #print('old:'+str(mail_id_list_old))

        try:
            server = email_login(email_user,password,pop3_server) # 登录邮箱并检查所有邮件
        except Exception as e:
            print('\n服务器连接错误，检查邮箱失败')
            time.sleep(wait_time)
            traceback.print_exc()
            time.sleep(1)
            continue

        #print('new:'+str(mail_id_list_new))

        for i in range(0,len(mail_id_list_new)): # 对新读取到的邮件ID列表中的每一项进行处理
            try: # 防止列表下标越界导致终止进程
                if mail_id_list_new[i] == mail_id_list_old[i]: # 如果新读取的ID与原始ID一致，则跳过
                    unprocessed_email_number = unprocessed_email_number # 没啥用，纯浪费算力
            except:
                unprocessed_email_number = unprocessed_email_number + 1 # 未处理邮件数计数器+1
                print('发现未处理邮件，ID：'+mail_id_list_new[i])

        #if len(mail_id_list_new) > len(mail_id_list_old):
            #mail_id_list_old = mail_id_list_new

        if unprocessed_email_number > 0: # 如果未处理邮件数不为0
            print('未处理邮件总数：'+str(unprocessed_email_number)+'\n开始依次处理')

            for i in range(len(mail_id_list_new)-unprocessed_email_number+1,len(mail_id_list_new)+1): # 依次处理每一份未处理邮件
                file_name = '' # 文件名
                if_max = False # 用户是否请求绕过分辨率限制
                if_suffix_error = False # 后缀名检查
                if_pixel_error = False # 像素数限制检查
                attachment_name_list = [] # 下载到的附件名称列表
                attachment_path_list = [] # 下载到的附件路径列表
                image_name_list = [] # 图片名称列表
                image_path_list = [] # 图片路径列表
                image_size_total = 0 # 附件文件总大小
                model = 0 # 所使用的超分辨率模型序号
                image_number_total = 0 # 暂时存储可处理图片总数
                print('\n下载邮件中，ID：'+ mail_id_list_new[i-1])

                try: # 读取邮件
                    resp,lines,octets = server.retr(i)
                    msg_content=b'\r\n'.join(lines).decode(encoding='utf-8',errors='ignore')
                except:
                    try:
                        print('邮件ID'+mail_id_list_new[i-1]+'获取出错！重试连接服务器')
                        server = email_connect(email_user,password,pop3_server) # 登录邮箱并检查所有邮件
                        resp,lines,octets = server.retr(i)
                        msg_content=b'\r\n'.join(lines).decode(encoding='utf-8',errors='ignore')
                    except Exception as e:
                        print('邮件再次获取出错！跳过该邮件：' + mail_id_list_new[i-1])
                        mail_id_list_old.append(mail_id_list_new[i-1])
                        write_id_txt()
                        traceback.print_exc()
                        time.sleep(1)
                        continue

                msg,From,To,Subject = email_analyze(msg_content)
                print(f'发件人：{From}；主题：{Subject}')

                if_vip = vip_verify(From) # 检查是否在VIP邮箱列表中
                print(From + '为VIP：' + str(if_vip))

                if_blacklist = blacklist_verify(From)
                print(From + '为黑名单用户：' + str(if_blacklist))

                if if_blacklist == True:
                    print('跳过黑名单用户' + From + '的请求')
                    mail_id_list_old.append(mail_id_list_new[i-1])
                    write_id_txt()
                    continue

                if ('max' in Subject): # 检查用户是否请求绕过绕过分辨率限制
                    if_max = True
                    print('用户请求绕过分辨率限制')

                if ('model' in Subject): # 检查用户是否请求自定义处理模型
                    if ('model0' in Subject):
                        model = 0
                    elif ('model1' in Subject):
                        model = 1
                    elif ('model2' in Subject):
                        model = 2

                # 主要应答部分
                if ('放大' in Subject): 
                    image_download(msg_content) # 下载附件

                    for file_number in range(0,len(attachment_name_list)): # 依次处理附件列表中的每一项

                        file_name = attachment_name_list[file_number]
                        if_processable, if_suffix_error = if_suffix(file_name) # 依次检查后缀
                        if_processable, if_pixel_error = pixel_verify(file_name) # 依次检查图片分辨率

                        if if_processable == True: # 如果文件为可处理的图片
                            image_name_list = image_name_list + [file_name] # 将该文件名称移添加到图片名称列表
                            image_path_list = image_path_list + [os.path.join(PATH_send,file_name)] # 将该文件路径移添加到图片路径列表                            
                        else:
                            print('文件不可处理：'+str(file_name))
                            remove_file((os.path.join(PATH_recv,file_name)),PATH_arch)

                    if len(image_name_list) > 0:

                        for image_name in image_name_list: # 将曾经已处理的文件直接移至归档文件夹
                            if os.path.exists(os.path.join(PATH_send, to_png(image_name))):
                                remove_file(os.path.join(PATH_recv, image_name),PATH_arch)
                                print(image_name+'已存在，跳过处理')

                        print('开始超分辨率处理')
                        pic_process(model) # 超分辨率处理，模型：0:Real-ESRGAN；1:Real-ESRGAN-anime；2:RealSR
                        print('附件中的文件：'+str(image_name_list))

                        image_number_total = len(image_name_list)

                        print('一共'+str(image_number_total)+'张！')

                        pic_number = 0

                        image_processed_total = 0

                        while pic_number != image_number_total:
                            print('pic-number'+str(pic_number))
                            quality = 0 # 使用JPG压缩时的质量，等于0时不压缩
                            image_name_list[pic_number]=to_png(image_name_list[pic_number]) # 将名称后缀改为.png
                            image_path_list[pic_number]=to_png(image_path_list[pic_number]) # 将路径后缀改为.png
                            image_size = os.path.getsize(image_path_list[pic_number]) # 获取图片文件大小
                            print(str(image_path_list)+str(image_name_list)+str(pic_number)+str(image_size))

                            if image_size > 200000000: # 当PNG文件大于约200MB时
                                quality = 75
                            elif image_size > 100000000: # 当PNG文件大于约100MB时
                                quality = 85
                            elif image_size > 50000000: # 当PNG文件大于约50MB时
                                quality = 90
                            elif image_size > 20000000: # 当PNG文件大于约20MB时
                                quality = 95
                            elif image_size > 10000000: # 当PNG文件大于约10MB时
                                quality = 98

                            print('压缩率为'+str(quality))

                            if quality != 0:
                                image_path_list[pic_number], image_name_list[pic_number] = pic_compression(os.path.join(PATH_send,image_name_list[pic_number]),image_name_list[pic_number],quality) # 压缩文件，替换文件名

                            image_size = os.path.getsize(image_path_list[pic_number]) # 重新检查压缩后的图片大小

                            if image_size > 48000000: # 如果单张图片的大小仍超出限制，则再次压缩这张图片
                                pic_number = pic_number - 1 # 重复本次处理
                                continue # 提前进入下一循环

                            image_size_total = image_size_total + image_size

                            print('当前累计图片大小：'+str(image_size_total))

                            if image_size_total > 48000000:

                                image_name_list_send = image_name_list[0:pic_number] # 待发送的附件为图片列表的第一项到本次处理的上一项，复制名称
                                image_path_list_send = image_path_list[0:pic_number] # 同上，复制路径

                                image_processed_total = image_processed_total + pic_number +1

                                subject_content = '已完成您的部分请求（这份邮件中未包含全部附件）'
                                body_content = '处理后的图片位于附件中。由于附件大小限制，图片将分批发送给您，请查收其余邮件。感谢您参与emailSR ' + version + '测试！'
                                print('\n发送邮件预览：主题：'+subject_content+'\n正文：'+body_content+'\n附件：'+str(image_name_list_send))
                                mail_send(subject_content,body_content,image_path_list_send,image_name_list_send) #发送邮件

                                del image_name_list[0:pic_number]
                                del image_path_list[0:pic_number]

                                image_number_total = image_number_total - pic_number

                                pic_number = -1 # 需要仔细想想233
 
                                image_size_total = 0 # 清零总文件大小                         

                            pic_number = pic_number + 1

                            print('结束'+str(pic_number))

                        #for pic_number in range(0,len(image_number_total)): # 依次检查文件是否超过所限制的大小

                        image_processed_total = image_processed_total + pic_number

                        if len(image_name_list) != 0:
                            print('当前附件中的图片大小：'+str(image_size_total))
                            subject_content = '已完成您的本次请求'
                            body_content = '处理后的图片位于附件中，共收到' + str(len(attachment_name_list)) + '个附件，其中' + str(image_processed_total) + '张图片已被处理。感谢您参与emailSR ' + version + '测试！'
                            print('\n发送邮件预览：主题：'+subject_content+'\n正文：'+body_content+'\n附件：'+str(image_name_list))
                            mail_send(subject_content,body_content,image_path_list,image_name_list) #发送邮件
                        else:
                            print('没有可发送的附件，未对邮件进行回复')
                    else:
                        print('未能完成请求，ID：' + mail_id_list_new[i-1])
                        subject_content = '未能完成您的请求'
                        if len(attachment_name_list) > 0:
                            if if_pixel_error == True:
                                body_content = '共收到' + str(len(attachment_name_list)) + '个附件，但由于图像分辨率过高，无法完成您的请求。请检查发件附件内容，或查看emailSR的使用规则。如需更多帮助请联系酷安@珊瑚宫心海。感谢您参与emailSR ' + version + '测试！'
                            elif if_suffix_error == True:
                                body_content = '共收到' + str(len(attachment_name_list)) + '个附件，但由于附件后缀不符合要求，无法完成您的请求。请检查发件附件内容，或查看emailSR的使用规则。如需更多帮助请联系酷安@珊瑚宫心海。感谢您参与emailSR ' + version + '测试！'
                        else:
                            body_content = '抱歉，没有收到附件。请检查发件附件内容，或查看emailSR的使用规则。如需更多帮助请联系酷安@珊瑚宫心海。感谢您参与emailSR ' + version + '测试！'
                        print('\n发送邮件预览：\n主题：'+subject_content+'\n正文：'+body_content+'\n附件：'+str(image_name_list))
                        #mail_send(subject_content,body_content,image_path_list,image_name_list) #发送邮件
                    mails_send = mails_send + 1
                else:
                    print('无关邮件，跳过')
                print('处理完成')
                mail_id_list_old.append(mail_id_list_new[i-1])
                write_id_txt()
                time.sleep(1)

            print('\n本次应答的邮件数:'+str(mails_send))

        else:
            print('没有发现未处理邮件')
            mail_id_list_old = mail_id_list_new
            write_id_txt()

        try:
            server.quit() # 关闭POP3连接
        except:
            print('未能执行POP3连接关闭操作，连接可能已被提前中止')
        time.sleep(wait_time)

    except Exception as e:
        traceback.print_exc()
        time.sleep(wait_time)

print('\n进程已完成\n')
















