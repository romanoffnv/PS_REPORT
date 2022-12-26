def helprinter(txt):
    txt = txt.split()
    return txt
    # if ' ' in txt:
    #     ind = txt.index(' ')
    #     return txt[:ind]
    # else:
    #     return txt
    # txt = txt.capitalize()
    # res = f'{txt} world'
    # return res

L = ['hello', 
     'good bye', 
     'fuck you', 
     'sucks', 
     'kicks ass', 
     'fucks you']
L = [helprinter(x) for x in L]
print(L)
# check if the first word has 's' in the end
def listmerger(L1, L2):
    L_plates = []
    for x, y in zip(L1, L2):
        if x != '':
            L_plates.append(x)
        if x == '' or len(x) < len(y):
            L_plates.append(y)    
    return L_plates

L1 = ['1', '2', '3', '', '', '', '']
L2 = ['', '', '', 'a', '/N 0015', '', '']
L3 = ['', '', '', '', 'S/N 0015286', 'c', 'd']


LL = listmerger(L1, L2)
LL = listmerger(LL, L3)
print(LL)

# for x, y in zip(L1, L2):
#         if x != '':
#             L_plates.append(x)
#         elif x == '':
#             L_plates.append(y)    
#         return L_plates