import re
 
# content = "abc:ab   cabc"
# rex = re.search(r'[:、.．： \t]{1,}', content)
# rex2 = re.split(r'[:、.．： \t]{1,}', content)
# if rex == None:
#     print(rex)
# else:
#     print(rex.group())
# print(rex2)

s = "s"
# print(s.encode('UTF-8').isalnum())
# numColumnList = []
# numColumnList = list(set(numColumnList)) # 删除重复的列数
# numColumnList.sort()
# print(numColumnList)
if u'\u4e00' <= s <= u'\u9fa5':
    print("hanzi")
else:
    print("bushi hanzi")