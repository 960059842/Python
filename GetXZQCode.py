"""
获取县以上行政区划代码
"""
# encoding=utf8

import requests
from lxml import etree
from Tool import MSSQL

"""
获取网页源代码
"""
def getHtml(url):
    headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.62 Safari/537.36'
    }
    print('URL：',url)
    response1=requests.get(url,headers=headers)
    html=response1.text
    #html = str(html).replace('&nbsp;', '')
    html = str(html).replace("<span style='mso-spacerun:yes'>   </span>", '')
    html = str(html).replace("<span style='mso-spacerun:yes'> </span>", '')

    #print(html)
    return html

"""
获取父节点
"""
def getParentID(list,code):
    #根节点
    if code[2:]=='0000':
        return 0

    #list = [{'id': 1, 'code': '110000', 'name': '北京市'}, {'id': 2, 'code': '110101', 'name': '东城区'}, {'id': 3, 'code': '110102', 'name': '西城区'}, {'id': 4, 'code': '110105', 'name': '朝阳区'}]
    #查询不到会抛出异常StopIteration
    try:
        parentcode=code[0:4]+'00'
        return list[next(index for (index, d) in enumerate(list) if d["code"] == parentcode)]['id']
    except StopIteration:
        try:
            parentcode = code[0:2] + '0000'
            return list[next(index for (index, d) in enumerate(list) if d["code"] == parentcode)]['id']
        except StopIteration:
            return -1

def insertData(list):
    ms = MSSQL(host=".", user="sa", pwd="feng", db="DncZeus")
    # 清除数据
    sql = "DELETE FROM dbo.City"
    ms.ExecNonQuery(sql)
    # 插入Sql
    insertsql = "INSERT INTO dbo.City( ID ,Code ,Name ,ParentID) VALUES(%s,%s,%s,%s) "
    # 插入list
    listinsert = []
    # Tuple 是不可变 list，一旦创建了一个 tuple 就不能以任何方式改变它。
    inserttuple = ["0", "1", "2", "3"]
    for city in list:
        inserttuple[0]=city["id"]
        inserttuple[1] = city["code"]
        inserttuple[2] = city["name"]
        inserttuple[3] = city["parentid"]
        listinsert.append(tuple(inserttuple))

    # 插入数据
    ms.ExecuteMany(
        insertsql,
        listinsert
    )
    print('数据初始化成功！！！')

def main():
    url = 'http://www.mca.gov.cn/article/sj/xzqh/2019/201901-06/201906211421.html'
    html = getHtml(url)
    root = etree.HTML(html, etree.HTMLParser())
    # print(root.items())
    nodes = root.xpath('//table/tr/td')
    index=0
    id=1
    code=''
    list=[]
    for node in nodes:
        print(node.text)
        if node.text!=None and node.text!='行政区划代码':
            if index%2==0:
                code=node.text
            else:
                list.append({
                    'id':id,
                    'code':code,
                    'name':node.text,
                    'parentid':getParentID(list,code)
                })
                id=id+1
            index=index+1

    print(list)
    insertData(list)

if __name__=='__main__':
    print('获取县以上行政区划代码')
    main()
