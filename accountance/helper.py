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