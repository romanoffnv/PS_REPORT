import json

D_om_diesels = {
                '100230350001дэсwt(инв.№эл-)': 'ДЭС WT10023035 (инв. №ЭЛ-0001)',
                '1961000004235дэсchh(инв.№)': 'ДЭС CHH1961 (инв. №000004235)',
                '445141206023000003798дэсcumminscd(s)№(инв.№)': 'ДЭС Cummins C44D5(S) №141206023 (инв. №000003798)',
                '442440000004312дэсhgpcchk(инв.№)': 'ДЭС HG44PC CHK2440 (инв. №000004312)',
                '39606007дэсинв.(f)': ' ДЭС инв.396 (F06007)',
                '39906010дэсинв.(f)': 'ДЭС инв.399 (F06010)',
                '441945000004326дэсhgpccnn(инв.№)': 'ДЭС HG44PC CNN1945 (инв. №000004326)',
                '30400727816000002219дэсад-с-т(инв.№)': 'ДЭС АД-30С-Т400 727816 (инв. №000002219)',
                '5064000004236дэсchv(инв.№)': 'ДЭС CHV5064 (инв. №000004236)',
                '5000003346дэс№(инв.№)': 'ДЭС № 5 (инв. №000003346)',
                '304000013дэсад--т/(инв.№ю/нб-)': 'ДЭС АД-30-Т/400 (инв. № Ю/НБ-0013)',
                '448805000005567дэсalfa(инв.№)': 'ДЭС ALFA 448805 (инв. №000005567)',
                '696602000005568дэсalfa(инв.№)': ' ДЭС ALFA 696602 (инв. №000005568)',
                '2439000004408дэсchк(инв.№)': 'ДЭС CHК2439 (инв. №000004408)',
                '665дэсcumminscd': 'ДЭС Cummins C66D5',
                '902788000003647дэсkubota(инв.№)': 'ДЭС Kubota 902788 (инв. №000003647)',
                '5767130002дэс(инв.№эл-)': 'ДЭС 576713 (инв. №ЭЛ-0002)',
                '441907010дэсhggl№': ' ДЭС HG44GL № 1907010',
                '397дэсинв.': ' ДЭС инв.397',
                '398дэсинв.': ' ДЭС инв.398',
                '1дэс': ' дэс1',
            }        
json.dump(D_om_diesels, open("D_om_diesels.json", 'w'))