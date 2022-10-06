# JWT token has to be put manually for now from the postman to headers 'Authorization'
import requests
from pprint import pprint
import json
import pandas as pd
import itertools


def main():
    url = "https://online.omnicomm.ru/ls/api/v2/tree/vehicle"

    payload={}
    headers = {
    'Content-Type': 'application/json',
    'Authorization': 'JWT eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6OTc2LCJsb2dpbiI6InBha2Vyc2VydmljZSIsInV1aWQiOiIyYWE1ZjgwZC05NWQ1LTM4YTAtOWVmZC1jMDMzNGMwOGNhNzYiLCJhdXRvY2hlY2tfaWQiOjgwOCwicGVybWlzc2lvbnMiOlsiI1VzZXIiLCIjdXNlcl9hZGQiLCJhc2UucmVwb3J0cy5tYXAyIiwiYXNlLnJlcG9ydHMubG9jYXRpb24iLCJhc2UucmVwb3J0cy5yZWZyaWdlcmF0b3J3b3JrIiwiYXNlLnJlcG9ydHMudGlyZXByZXNzdXJlIiwiYXNlLnJlZnJpZ2VyYXRvciIsImFzZS50cG1zIiwiYXNlLm1vZGJ1c2dlbiIsImFzZS5tb2RidXNsbHMiLCJhc2UuZ3JvdXBzLmRyaXZlci52aWV3IiwiYXNlLmdyb3Vwcy5yb3V0ZS52aWV3IiwiYXNlLnVzZXJyZXBvcnRzY29udHJvbCIsImFzZS5yZXBvcnRzLnBlcmlvZGljc2VydmljZSIsIm1hcHMueWFuZGV4Lm1hcCIsIm1hcHMud2lraS5tYXAiLCJtYXBzLm9zbS5tYXAiLCJtYXBzLm9tbmljb21tLm1hcCIsIm1hcHMub21uaWNvbW0uamFtcyIsImFzZS5yZXBvcnRzLmZ1ZWxsZXZlbHMiLCJhc2UuZ3JvdXBzLmdlb3pvbmUuY3VzdG9tIiwiYXNlLmdyb3Vwcy5nZW96b25lLnZpZXciLCJhc2UuZ3JvdXBzLmdlb3pvbmUuZWRpdHByb2ZpbGUiLCJzZXJ2aWNlLmZ1ZWxiYWxhbmNlIiwic2VydmljZS5yZXBvcnRzLmZ1ZWxzaGVldCIsImFzZS5yZXBvcnRzLmdyb3Vwc3RhdCIsImFzZS5yZXBvcnRzLnVuaXZlcnNhbCIsImlkZW50aWZpY2F0aW9uIiwibWFpbnBhZ2UubmV3IiwibWFwcy5nb29nbGUubWFwIiwibWFwcy5nb29nbGUuc2F0ZWxpdGUiLCJtYXBzLmdvb2dsZS5oeWJyaWQiLCJhc2UucmVwb3J0cy5ldmVudHMiLCJhc2UuZ3JvdXBzLnZlaGljbGUuY3VzdG9tIiwiYXNlLmdyb3Vwcy52ZWhpY2xlLnZpZXciLCJhc2UuZ3JvdXBzLnZlaGljbGUuZWRpdHRyZWUiLCJhc2UuZ3JvdXBzLnZlaGljbGUudmlld3Byb2ZpbGUiLCJhc2UuZ3JvdXBzLnZlaGljbGUuZWRpdGluZm8iLCJhc2UuZ3JvdXBzLnZlaGljbGUuZWRpdHByb2ZpbGUiLCJhc2UuZ3JvdXBzLnZlaGljbGUuc2VydmljZW1uZ21udCIsImFzZS5ncm91cHMudmVoaWNsZS5yZWNhbGMiLCJhc2UucmVwb3J0cy5nZW96b25lc3JlcG9ydCIsImFzZS5ncm91cHMuZ2Vvem9uZS5kZWxldGUiLCJhdXRoLmNsaWNrcmVwb3J0cyIsImF1dG9jaGVjay5hY2Nlc3MiLCJzZXJ2aWNlLmJpbGxpbmcubmV3IiwibWFwcy53aWtpLm1hcCIsIm1hcHMub3NtLm1hcCIsIm1hcHMuc3B1dG5pay5tYXAiLCJzZXJ2aWNlLmNvcHMucmVhZCIsInNlcnZpY2UuY29wcy51cGRhdGUiLCJzZXJ2aWNlLmNvcHMuYWRkIiwic2VydmljZS5DQU5fYnlfZW1kZCIsInNlcnZpY2Uub3ZtcyIsInNlcnZpY2UuaGZtcy5tZXNzYWdlcyIsInNlcnZpY2UucmVwb3J0cy5oZm1zX3N0YXR1cyIsInNlcnZpY2UucmVwb3J0cy5oZm1zX21lc3NhZ2VzIiwidXNlcnMubW9iaWxlLnB1c2giLCJzZXJ2aWNlLmNvcHMucmVhZCIsInNlcnZpY2UuY29wcy51cGRhdGUiLCJzZXJ2aWNlLmNvcHMuYWRkIiwic2VydmljZS5iaWxsaW5nLnVzZXIuZWRpdG9yIiwic2VydmljZS5zYWZlZHJpdmluZ3JlcG9ydCIsInNlcnZpY2UuYmlsbGluZy5uZXciLCJhc2UuZnVlbG1hc3MiLCJmdWVsbWFzcyIsImZ1ZWwuYmFsYW5jZS5uZXciLCJmdWVsLmJhbGFuY2UuZmVzX2FwaSIsImxsczUuYWNjZXNzIiwiYXV0aC5jbGlja3JlcG9ydHMiLCJzZXJ2aWNlLkNBTl9ieV9lbWRkIl0sImRyaXZlcmdyb3VwX2lkIjoiN2QxZDlkZGItNDE2ZC0zOWNkLThmOGQtYWY2NDNlNTQxZTE4IiwidmVoaWNsZWdyb3VwX2lkIjoiZWMyMjJkZWYtODU3Zi0zYzljLWE2N2EtOTAzZGRiOTEwYTJhIiwidXNlcmdyb3VwX2lkIjpudWxsLCJnZW96b25lZ3JvdXBfaWQiOiI1YTJkZjA0Yi04MmQ5LTMxYzAtODRkZS00N2EwZGZlNDcwNzEiLCJyb3V0ZWdyb3VwX2lkIjoiODQ1MGU0YzktOWQ1Ny0zNDI1LTk1YWQtZDQzYWMxNjNiNDliIiwic2VydmVyX25hbWUiOiJncmVlbnNjYWxlIiwibWF4X3JlcG9ydF9wZXJpb2QiOjYwLCJzZXJ2ZXIiOnsiaWQiOjYsImhvc3QiOiJodHRwOi8vMTAuNTUuNi4xMTYiLCJwb3J0Ijo4MDgwLCJzbHVnIjoiZ3JlZW5zY2FsZSIsImZxZG4iOiIxMC41NS42LjExNiIsInZvbHVtZV91bml0cyI6IkwiLCJsaWdodCI6MH0sInJvbGVzIjpbeyJpZCI6NCwibmFtZSI6IlVzZXIiLCJzbHVnIjoidXNlciIsIndlaWdodCI6MTB9LHsiaWQiOjY4LCJuYW1lIjoidXNlcl9hZGQiLCJzbHVnIjoidXNlcl9hZGQiLCJ3ZWlnaHQiOjB9XSwid2wiOnt9LCJpYXQiOjE2NjMxOTM4NjEsImV4cCI6MTY2MzE5NzQ2MSwiYXVkIjoiYXNlIn0.hr3662Cr5XcARoPY1d28XOgd7v8OS7b1lal0I95hbjE'
    }

    response = requests.request("GET", url, headers=headers, data=payload)

    data = response.json()
    # the func returns 
    def list_maker(x):
        group = data['children'][x]['objects']
        L = [x['name'] for x in group]
        return L 

    L_group_access = [0, 1, 2, 3, 4, 5, 6, 7]
    L_units = []
    for i in L_group_access:
        L_units.append(list_maker(i))
    

    # Get the list of the group names by length multiplication
    L_group_len = []
    for i in L_units:
        L_group_len.append(len(i))

    L_group_name = ['TOYOTA', 'ГРП 1', 'ГРП 2', 'ГРП 3', 'ГНКТ 1', 'ГНКТ 2', 'ГНКТ 3', 'ГНКТ 4']
    
    L_group_name_total = [(i + '**').split('**') * j for i, j in (zip(L_group_name, L_group_len))]
    L_group_name_total = list(itertools.chain.from_iterable(L_group_name_total))
    L_group_name_total = list(filter(None, L_group_name_total))

    L_units = list(itertools.chain.from_iterable(L_units))

    # Build the df
    pd.set_option('display.max_rows', None)
    df = pd.DataFrame(zip(L_group_name_total, L_units), columns=['Group', 'Units'])

    print(df)
    
    # L_units = list(itertools.chain.from_iterable(L_units))
    # print(L_units) 
    # 
    
    # df = pd.DataFrame(zip(L_group, L))
    # print(df)
    # frac1 = data['children'][1]['objects']
    # L_frac1_units = [x['name'] for x in frac1]
    # L_frac1_group = ['ГРП 1' for x in L_frac1_units]
    
    # print(L_frac1_units)
    # print(L_frac1_group)

if __name__ == '__main__':
    main()