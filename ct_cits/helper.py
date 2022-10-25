from collections import Counter
import re
L_string = ['УУ 0775 86','86 УМ 8475']
L_digits = []
for i in L_string:
    L_digits.append(re.findall('\d+', i))
    
print(L_digits)