"""
config.py
邮箱账号，密码配置文件，在该文件中填写需要导出邮件的用户账号和密码，以及对应的代理proxy
如果需要新增，复制粘贴下方的格式，再填写即可
格式如下：
user_profile = [{
    'username':'admin',
    'password':'Admin@123',
    'proxy':'http://ip:port'
},{
    'username':'test',
    'password':'test123456',
    'proxy':'http://ip:port'
},{
    'username':'填入新增的账号',
    'password':'新增账号的密码',
    'proxy':'socks5://ip:port'
}]

"""
# 在下方进行修改, 如果需要proxy 代理，则添加如上面的proxy 字段

user_profile = [{
    'username': 'xxxa',
    'password': 'xxx',
    'proxy':'http://8.1xx.xx.x:1080'
}, 
{   'username': 'xxxb',
    'password': 'xxx',
    'proxy':'http://x.1xx.xxx.x32:1080'
}
]

download_path = '.\\download\\'
