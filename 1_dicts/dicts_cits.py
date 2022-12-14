
import json
D_ct_plates = {
        '567': 'Е567НС 186',
        '010': 'а010ук 186',
        '2467': '2467ат 86',
        '4235': '000004235',
        '6241': '6241АУ86',
        '2219': '000002219',
        '072': '072нива',
        '766': 'Е766ОВ 186',
        '232': 'о232уо 86',
        '894': 'а894ое 186',
        '403': 'в403та 186',
        '890': '890нива',
        '233': 'у233вк 186',
        '885': 'Е885МХ 186',
        '600': 'т600ак 186',
        '192': 'у192вм 186',
        '725': 'у725вм 186',
        '666': 'т666ак 186',
        '342': 'а342ср 186',
        '555': 'т555ак 186',
        '031': 'а031на 186',
        '834': 'а834нм 186',
        '297': 'р297ам 186',
        '856': 'а856нм 186',
        '347': 'в347кт 186',
        '689': 'а689ке 186',
        '060': 'х060ас 186',
        '494': 'Е494ОВ 186',
        '953': 'а953вв 186',
        '109': 'в109ек 186',
        '097': 'н097вк 186',
        '081': 'в081ек 186',
        '627': 'а627ка 186',
        '946': 'а946те 186',
        '712': 'а712те 186',
        '035': 'а035на 186',
        '374': 'е374уе 86',
        '952': 'а952вв 186',
        '771': 'р771ае 186',
        '334': 'а334рр 186',
        '893': 'а893ев 186',
        '628': 'а628ка 186',
        '680': 'а680на 186',
        '557': 'в557ат 186',
        '968': 'а968те 186',
        '499': 'а499на 186',
        '153': 'а153км 186',
        '049': '049нива',
        '652': 'в652нт186',
        '469': 'р469ав 186',
        '723': 'у723вм 186',
        '094': 'в094ек 186',
        '445': 'Е445НС 186',
        '877': 'Е877СВ 186',
        '197': 'а197хо 186',
        '905': 'Е905МУ 186',
        '029': 'а029на 186',
        '596': 'а596см 186',
        '004': 'в004еу 186',
        '845': 'Е845МХ 186',
        '694': 'а694мв 186',
        '975': 'а975те 186',
        '319': 'а319рр 186',
        '033': 'а033на 186',
        '626': 'а626ав 186',
        '371': 'а371ср 186',
        '149': 'а149ср 186',
        '679': 'а679тв 186',
        '052': 'а052ск 186',
        '699': 'в699ев 186',
        '032': 'а032на 186',
        '681': 'а681ко 186',
        '889': 'Е889ОВ 186',
        '158': 'а158сн 186',
        '427': 'в427тр 186',
        '837': 'Е837МХ 186',
        '179': 'Е179СУ 186',
        '3792': 'АС 3792 86',
        '8091': '8091 АХ 86',
        '3805': 'АТ 3805 86',
        '728': 'у728вм 186',
        '872': 'а872тв 186',
        '2458': '2458 НТ 77',
        '246': 'АТ 2467 86',
        '063': 'в063ек 186',
        '100': 'в100ек 186',
        '562': 'А562НК 186',
        '7824': 'АН 782482',
        '3806': 'АТ 3806 86',
        '86 УМ 8475': '8475 УМ 86',
        '822': 'Е822МХ 186',
        '358': 'А358РР186',
        '340': 'Е340ОМ 186',
        '8090': '8090 АХ 86',
        'ДЭС АД 30 Т 4002шт': 'ДЭС АД-30С-Т400 727816 инв 000002219',
        '№0002': 'ДЭС 576713 инв ЭЛ-0002',
        '№2219': 'ДЭС АД-30С-Т400 727816 инв 000002219',
        '6566': 'АС 6566 86',
        'ДЭС4235  30КВт1шт': 'ДЭС CHH1961 инв 000004235',
        '0291': 'ТА 0291 86',


        # Crutches for bitten plates:
        'Н764082': 'АН 7640 82',
}

json.dump(D_ct_plates, open( "D_ct_plates.json", 'w' ))
