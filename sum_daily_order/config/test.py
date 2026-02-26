import re
str1 = "google runoob taoba\n\r"
regx = r'^goo'
result = re.findall(regx, str1)
print(result)