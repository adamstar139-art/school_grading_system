import streamlit as st
import pandas as pd
import sqlite3
import openpyxl
import urllib.parse
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import streamlit.components.v1 as components

# ==============================================================================
# 1. تهيئة قاعدة البيانات والبيانات المعتمدة للطلاب
# ==============================================================================
DB_NAME = "school_grading_system.db"

STUDENTS_INIT = [
    # الصف الأول المتوسط - فصل 1
    {"id": "1167628468", "name": "ابراهيم بن محمد بن علي الوهيبي", "grade": "الأول المتوسط", "class": 1, "phone": "966504158122"},
    {"id": "2395664317", "name": "بلال عبدالرزاق عيسى العيسى", "grade": "الأول المتوسط", "class": 1, "phone": "966507448712"},
    {"id": "1170582165", "name": "حسام بن محمد بن علي ال رايان البارقي", "grade": "الأول المتوسط", "class": 1, "phone": "966504445699"},
    {"id": "1169004353", "name": "ريان عبدالله جابر الاسمري", "grade": "الأول المتوسط", "class": 1, "phone": "966554260960"},
    {"id": "2446713998", "name": "زيد زياد عبد اللطيف ابو قبع", "grade": "الأول المتوسط", "class": 1, "phone": "966590123455"},
    {"id": "2527104554", "name": "سامي سعد عباس حمد", "grade": "الأول المتوسط", "class": 1, "phone": "966591781701"},
    {"id": "1170111759", "name": "سعد ناصر سعد السيف", "grade": "الأول المتوسط", "class": 1, "phone": "966503219351"},
    {"id": "1195559479", "name": "عبدالعزيز عبدالله عبدالعزيز العمار", "grade": "الأول المتوسط", "class": 1, "phone": "966555838394"},
    {"id": "1153310501", "name": "عبدالله بن سليمان بن عبدالله الراجحي", "grade": "الأول المتوسط", "class": 1, "phone": "0551418881"},
    {"id": "1170836520", "name": "عبدالله سعد بن محمد العيشان", "grade": "الأول المتوسط", "class": 1, "phone": "966504217660"},
    {"id": "1171448515", "name": "علي احمد علي كريري", "grade": "الأول المتوسط", "class": 1, "phone": "966558885481"},
    {"id": "1170853053", "name": "علي سعد علي القحطاني", "grade": "الأول المتوسط", "class": 1, "phone": "966505466546"},
    {"id": "1172018036", "name": "عمر عبدالله سعد الجبرين", "grade": "الأول المتوسط", "class": 1, "phone": "966555249420"},
    {"id": "2552851368", "name": "مازن اسلام احمد ابراهيم موسى", "grade": "الأول المتوسط", "class": 1, "phone": "966550490495"},
    {"id": "013609321", "name": "محمد أحمد علي عقيل", "grade": "الأول المتوسط", "class": 1, "phone": "966546000184"},
    {"id": "2394606749", "name": "محمد اشرف مسعود ابوخاطر", "grade": "الأول المتوسط", "class": 1, "phone": "966501276888"},
    {"id": "1170042046", "name": "محمد بن فيصل بن مصلح الشمراني", "grade": "الأول المتوسط", "class": 1, "phone": "966555832145"},
    {"id": "1169174164", "name": "محمد نايف فراج الدعجاني", "grade": "الأول المتوسط", "class": 1, "phone": "966554444782"},
    {"id": "2380890976", "name": "وائل بولعيش", "grade": "الأول المتوسط", "class": 1, "phone": "966591534495"},

    # الصف الأول المتوسط - فصل 2
    {"id": "1170348286", "name": "الوليد ابن خالد بن فهد العتيبي", "grade": "الأول المتوسط", "class": 2, "phone": "966558522229"},
    {"id": "1172433185", "name": "باسل محمد فرج الدوسري", "grade": "الأول المتوسط", "class": 2, "phone": "966537589781"},
    {"id": "1173391556", "name": "بسام بن عبدالكريم بن عبدالله الحرقان الدوسري", "grade": "الأول المتوسط", "class": 2, "phone": "966534467820"},
    {"id": "1169185053", "name": "تركي عبدالله مسفر الدوسري", "grade": "الأول المتوسط", "class": 2, "phone": "966505258369"},
    {"id": "1170108078", "name": "تميم فهد عبدالعزيز العزاز", "grade": "الأول المتوسط", "class": 2, "phone": "966554435692"},
    {"id": "1170970741", "name": "جاسر بن عبدالله بن منصور المطاطحة الحارثي", "grade": "الأول المتوسط", "class": 2, "phone": "966559455545"},
    {"id": "1168982427", "name": "راكان عبدالله يحي كريري", "grade": "الأول المتوسط", "class": 2, "phone": "966581727444"},
    {"id": "1172590968", "name": "ريان عبدالله منصور السبر", "grade": "الأول المتوسط", "class": 2, "phone": "966566959670"},
    {"id": "2392863888", "name": "ريان وليد حلاق", "grade": "الأول المتوسط", "class": 2, "phone": "966530527662"},
    {"id": "1170420473", "name": "سيف عبدالكريم بريك العصيمي", "grade": "الأول المتوسط", "class": 2, "phone": "966504277904"},
    {"id": "1168942108", "name": "صالح حسن فتحي سندي", "grade": "الأول المتوسط", "class": 2, "phone": "966557553922"},
    {"id": "1173182138", "name": "عبدالرحمن ابراهيم عبدالله الحضيف", "grade": "الأول المتوسط", "class": 2, "phone": "966558822674"},
    {"id": "1172448548", "name": "عبدالله صالح حمد الصفيان", "grade": "الأول المتوسط", "class": 2, "phone": "966558889978"},
    {"id": "1170000945", "name": "فهد ابن احمد بن فهد العثمان", "grade": "الأول المتوسط", "class": 2, "phone": "966504484898"},
    {"id": "1167092616", "name": "فهد عويض ثعيل المطيري", "grade": "الأول المتوسط", "class": 2, "phone": "966508271056"},
    {"id": "1170413171", "name": "فهد نايف فهد الحسينان", "grade": "الأول المتوسط", "class": 2, "phone": "966504140616"},
    {"id": "1170294118", "name": "فيصل موينع عبدالله بن موينع", "grade": "الأول المتوسط", "class": 2, "phone": "966541600918"},
    {"id": "1171524604", "name": "فيصل ناصر سيف العريفي", "grade": "الأول المتوسط", "class": 2, "phone": "966505474606"},
    {"id": "2502333707", "name": "محمد اسلام محمد دراز", "grade": "الأول المتوسط", "class": 2, "phone": "966556124553"},
    {"id": "1170374993", "name": "مشاري عثمان سعد ناصر السعد", "grade": "الأول المتوسط", "class": 2, "phone": "966500330693"},
    {"id": "1170884165", "name": "يزن محمد علي اليحيا", "grade": "الأول المتوسط", "class": 2, "phone": "966557072133"},
    {"id": "1170548737", "name": "يوسف محمد عبدالله الدوسري", "grade": "الأول المتوسط", "class": 2, "phone": "966556666176"},

    # الصف الثاني المتوسط - فصل 1
    {"id": "1163760935", "name": "احمد سامي بن احمد العمران", "grade": "الثاني المتوسط", "class": 1, "phone": "966551501503"},
    {"id": "1153756612", "name": "الوليد عبدالله بن ابراهيم المبدل", "grade": "الثاني المتوسط", "class": 1, "phone": "966505241627"},
    {"id": "1164269209", "name": "ذياب بن محمد بن ذياب بن محمد ال مربط القحطاني", "grade": "الثاني المتوسط", "class": 1, "phone": "966561169999"},
    {"id": "1163187972", "name": "راكان سالم بن محمد بن مسفر القحطاني", "grade": "الثاني المتوسط", "class": 1, "phone": "966556609291"},
    {"id": "1171617069", "name": "سعود خالد عبدالله الحمد", "grade": "الثاني المتوسط", "class": 1, "phone": "966555242944"},
    {"id": "1163458878", "name": "سعود مشعل بن ابراهيم الشثري", "grade": "الثاني المتوسط", "class": 1, "phone": "966598887996"},
    {"id": "1167623758", "name": "سلطان عبدالله حسن القحطاني", "grade": "الثاني المتوسط", "class": 1, "phone": "966563484825"},
    {"id": "1164769430", "name": "عبدالرحمن حمد بن محمد العريفي", "grade": "الثاني المتوسط", "class": 1, "phone": "966555556856"},
    {"id": "1167893740", "name": "عبدالرحمن ربيع جابر خبراني", "grade": "الثاني المتوسط", "class": 1, "phone": "966535924655"},
    {"id": "1159740032", "name": "عبدالعزيز سعود بن فهد العتيبي", "grade": "الثاني المتوسط", "class": 1, "phone": "966544155592"},
    {"id": "1164747436", "name": "عبدالمجيد بن محمد بن مسعود ال عايض القحطاني", "grade": "الثاني المتوسط", "class": 1, "phone": "966555275591"},
    {"id": "1160901128", "name": "فيصل بن عبدالله بن سعود بن عبدالعزيز الجميعة", "grade": "الثاني المتوسط", "class": 1, "phone": "966554949948"},
    {"id": "1162168627", "name": "مبارك صالح مبارك هليل", "grade": "الثاني المتوسط", "class": 1, "phone": "966553663819"},
    {"id": "1163212978", "name": "محمد بن عبدالله بن حمد بن ناصر بن عمران", "grade": "الثاني المتوسط", "class": 1, "phone": "966544779170"},
    {"id": "1161858301", "name": "محمد عبدالمحسن ناصر الحزام", "grade": "الثاني المتوسط", "class": 1, "phone": "966505264075"},
    {"id": "1175902442", "name": "محمد فايز عبدالرحمن بن يوسف", "grade": "الثاني المتوسط", "class": 1, "phone": "966505482728"},
    {"id": "1165686179", "name": "مشاري سلطان سالم الشمراني", "grade": "الثاني المتوسط", "class": 1, "phone": "966553908888"},
    {"id": "1166040053", "name": "معاذ عبدالله سعود العريفي", "grade": "الثاني المتوسط", "class": 1, "phone": "966505473192"},
    {"id": "1167081981", "name": "ناصر حسين محمد ال جبران", "grade": "الثاني المتوسط", "class": 1, "phone": "966550004952"},
    {"id": "1171868639", "name": "وائل بن عبدالله بن عامر علي ال عبيد الغامدي", "grade": "الثاني المتوسط", "class": 1, "phone": "966548888663"},
    {"id": "1163191222", "name": "يزيد بن طارق بن علي الحديثي", "grade": "الثاني المتوسط", "class": 1, "phone": "966554084040"},

    # الصف الثاني المتوسط - فصل 2
    {"id": "1166753291", "name": "ابراهيم بن مبارك بن راشد بن عبدالرحمن السبعان آل موينع", "grade": "الثاني المتوسط", "class": 2, "phone": "966555212896"},
    {"id": "1163613795", "name": "ابراهيم ياسر ابراهيم الحلوي", "grade": "الثاني المتوسط", "class": 2, "phone": "966502220990"},
    {"id": "1167148251", "name": "حامد بن محمد بن حامد شباط", "grade": "الثاني المتوسط", "class": 2, "phone": "966595001616"},
    {"id": "1164599977", "name": "حسام حسن محمد الشهري", "grade": "الثاني المتوسط", "class": 2, "phone": "966557775278"},
    {"id": "1169057351", "name": "خالد تركي عايض القحطاني", "grade": "الثاني المتوسط", "class": 2, "phone": "966536201378"},
    {"id": "1164120600", "name": "خالد داود بن عابد الحارثي", "grade": "الثاني المتوسط", "class": 2, "phone": "966501076244"},
    {"id": "1165839455", "name": "سطام عبدالعزيز عبدالله العريفي", "grade": "الثاني المتوسط", "class": 2, "phone": "966599791658"},
    {"id": "1163778960", "name": "سعود سلطان بن هليل العتيبي", "grade": "الثاني المتوسط", "class": 2, "phone": "966554820082"},
    {"id": "1166582989", "name": "طلال محمد منير المهدرس", "grade": "الثاني المتوسط", "class": 2, "phone": "966531167666"},
    {"id": "1165143783", "name": "عبدالكريم مساعد عبدالعزيز الهزاع", "grade": "الثاني المتوسط", "class": 2, "phone": "966503210252"},
    {"id": "1164277830", "name": "عبداللطيف ابراهيم محمد الطمرة", "grade": "الثاني المتوسط", "class": 2, "phone": "966505404365"},
    {"id": "1165495258", "name": "عبدالله سامي سعد الحوشاني", "grade": "الثاني المتوسط", "class": 2, "phone": "966555219086"},
    {"id": "013609088", "name": "علي احمد علي عقيل", "grade": "الثاني المتوسط", "class": 2, "phone": "966546000184"},
    {"id": "1164825802", "name": "عمر بن سعد بن هلال الشبانات", "grade": "الثاني المتوسط", "class": 2, "phone": "966505213725"},
    {"id": "1163537838", "name": "فارس مشعل عبدالله بن موينع", "grade": "الثاني المتوسط", "class": 2, "phone": "966555200719"},
    {"id": "1162761306", "name": "فهد عيسى محمد العيسى", "grade": "الثاني المتوسط", "class": 2, "phone": "966554499908"},
    {"id": "1164997858", "name": "مازن خالد دخيل المطيري", "grade": "الثاني المتوسط", "class": 2, "phone": "966501110052"},
    {"id": "2348937422", "name": "مازن رفعت محمد حاج النيل", "grade": "الثاني المتوسط", "class": 2, "phone": "966501331089"},
    {"id": "1172720045", "name": "محمد بن علي محسن العثيميني", "grade": "الثاني المتوسط", "class": 2, "phone": "966506256254"},
    {"id": "1166803245", "name": "نايف بن بندر بن خلفان العلوي", "grade": "الثاني المتوسط", "class": 2, "phone": "966532225560"},
    {"id": "1165668417", "name": "نواف عبدالعزيز عبدالله المرزوق", "grade": "الثاني المتوسط", "class": 2, "phone": "966501100076"},
    {"id": "1164387977", "name": "هادي سلطان هادي القحطاني", "grade": "الثاني المتوسط", "class": 2, "phone": "966505936192"},
    {"id": "1165002153", "name": "يزيد بن حسين بن متعب بن محمد كعكم", "grade": "الثاني المتوسط", "class": 2, "phone": "966550117805"},

    # الصف الثاني المتوسط - فصل 3
    {"id": "1166911709", "name": "ثامر عمر ابراهيم عثمان", "grade": "الثاني المتوسط", "class": 3, "phone": "966538384444"},
    {"id": "008464815", "name": "جهاد فارس عبدالقادر حناوي", "grade": "الثاني المتوسط", "class": 3, "phone": "966562674178"},
    {"id": "1164830562", "name": "خالد محمد عبدالكريم الخفاجي", "grade": "الثاني المتوسط", "class": 3, "phone": "966533074601"},
    {"id": "1188914319", "name": "سعد ابن مسفر بن سعد القحطاني", "grade": "الثاني المتوسط", "class": 3, "phone": "966508057005"},
    {"id": "1165099498", "name": "سعود بن عبدالله بن سعود السحامي", "grade": "الثاني المتوسط", "class": 3, "phone": "966500650867"},
    {"id": "1167770468", "name": "سعود ناصر سيف العريفي", "grade": "الثاني المتوسط", "class": 3, "phone": "966505474606"},
    {"id": "2344500760", "name": "سعيد محمد باوزير", "grade": "الثاني المتوسط", "class": 3, "phone": "966553435135"},
    {"id": "1164983874", "name": "طلال بن فهد بن عطيه بالحكم الزهراني", "grade": "الثاني المتوسط", "class": 3, "phone": "966567837159"},
    {"id": "2362260263", "name": "عبدالرحمن احمد جاسم الحمدي", "grade": "الثاني المتوسط", "class": 3, "phone": "966503432054"},
    {"id": "1167153434", "name": "عبدالعزيز ماجد راشد الزير", "grade": "الثاني المتوسط", "class": 3, "phone": "966500933390"},
    {"id": "1164512566", "name": "عبدالعزيز وليد ناصر بن سعران", "grade": "الثاني المتوسط", "class": 3, "phone": "966556660555"},
    {"id": "1167267341", "name": "عبدالله بن بندر بن فهد المسيحل", "grade": "الثاني المتوسط", "class": 3, "phone": "966500155334"},
    {"id": "2358022958", "name": "عز الدين احمد محمد سعد", "grade": "الثاني المتوسط", "class": 3, "phone": "966561317507"},
    {"id": "1167515020", "name": "عزام خالد شلهوب بن شلهوب", "grade": "الثاني المتوسط", "class": 3, "phone": "966506404016"},
    {"id": "1164747014", "name": "عزام فهد احمد صلوي", "grade": "الثاني المتوسط", "class": 3, "phone": "966555796951"},
    {"id": "4533080448", "name": "عمر وليد ياسين درويش علي", "grade": "الثاني المتوسط", "class": 3, "phone": "966557790508"},
    {"id": "1163397811", "name": "فارس ابن محمد بن سالم بن نويشي الوهبي الحربي", "grade": "الثاني المتوسط", "class": 3, "phone": "966583228278"},
    {"id": "1166629798", "name": "يزيد بن حمد بن مترك بن محمد ال مسعود القحطاني", "grade": "الثاني المتوسط", "class": 3, "phone": "966505203795"},
    {"id": "1167371093", "name": "يوسف عايد عواد البلوي", "grade": "الثاني المتوسط", "class": 3, "phone": "966531066289"},

    # الصف الثالث المتوسط - فصل 1
    {"id": "1158966166", "name": "أصيل ناصر بن محمد مذكور", "grade": "الثالث المتوسط", "class": 1, "phone": "966552149044"},
    {"id": "1162308223", "name": "خالد محمد مسدف معافا", "grade": "الثالث المتوسط", "class": 1, "phone": "966552680201"},
    {"id": "1161109093", "name": "راكان بن عبدالله بن سالم اليافعي", "grade": "الثالث المتوسط", "class": 1, "phone": "966504234219"},
    {"id": "1160805899", "name": "زياد احمد بن علي اللحيد", "grade": "الثالث المتوسط", "class": 1, "phone": "966504432362"},
    {"id": "1160267124", "name": "سطام محمد سعود الدوسري", "grade": "الثالث المتوسط", "class": 1, "phone": "966555260669"},
    {"id": "1163270869", "name": "سلطان احمد صالح الفنتوخ", "grade": "الثالث المتوسط", "class": 1, "phone": "966555242266"},
    {"id": "1161085236", "name": "ضاري صالح مهنا العازمي", "grade": "الثالث المتوسط", "class": 1, "phone": "966531111140"},
    {"id": "1160585624", "name": "عبدالعزيز عبدالله شراز المالكي", "grade": "الثالث المتوسط", "class": 1, "phone": "966556999627"},
    {"id": "1160050678", "name": "عبدالعزيز عبدالله عايض الاسمري", "grade": "الثالث المتوسط", "class": 1, "phone": "966555992269"},
    {"id": "1161021314", "name": "عبدالله عبيد عبدالله العتيبي", "grade": "الثالث المتوسط", "class": 1, "phone": "966597882020"},
    {"id": "1160857700", "name": "عبدالله فهد جلوي سالم الشرعي", "grade": "الثالث المتوسط", "class": 1, "phone": "966555457732"},
    {"id": "2502333723", "name": "عماد الدين اسلام محمد دراز", "grade": "الثالث المتوسط", "class": 1, "phone": "966556124553"},
    {"id": "1161418593", "name": "فهد عبدالرحمن فهد العتيبي", "grade": "الثالث المتوسط", "class": 1, "phone": "966552270402"},
    {"id": "1163074592", "name": "فيصل بن عبدالمحسن بن عايض العصيمي العتيبي", "grade": "الثالث المتوسط", "class": 1, "phone": "966505552320"},
    {"id": "1171918236", "name": "مازن خالد عبدربه الزهراني", "grade": "الثالث المتوسط", "class": 1, "phone": "966540707365"},
    {"id": "1158815876", "name": "محمد سلطان عبدالعزيز العيد", "grade": "الثالث المتوسط", "class": 1, "phone": "966503167770"},
    {"id": "1166075653", "name": "محمد مقعد ساير العتيبي", "grade": "الثالث المتوسط", "class": 1, "phone": "966536655992"},
    {"id": "1160693949", "name": "مشاري ابراهيم عبداللطيف المغري", "grade": "الثالث المتوسط", "class": 1, "phone": "966542744245"},
    {"id": "1160803878", "name": "مشاري علي موسى عقيلي", "grade": "الثالث المتوسط", "class": 1, "phone": "966502259722"},
    {"id": "1161661846", "name": "مهند عبدالله فهد الزكري", "grade": "الثالث المتوسط", "class": 1, "phone": "966558794720"},
    {"id": "1159404795", "name": "نواف وليد حمد الشعالان", "grade": "الثالث المتوسط", "class": 1, "phone": "966555798074"},
    {"id": "1168385894", "name": "يوسف نايف مقعد العتيبي", "grade": "الثالث المتوسط", "class": 1, "phone": "966505290037"},

    # الصف الثالث المتوسط - فصل 2
    {"id": "1156933093", "name": "تركي عبدالعزيز عبدالله المرزوق", "grade": "الثالث المتوسط", "class": 2, "phone": "966501100076"},
    {"id": "1160223317", "name": "تركي عثمان عبدالعزيز العثمان", "grade": "الثالث المتوسط", "class": 2, "phone": "966505226153"},
    {"id": "1159683497", "name": "راشد احمد فهد ال سعيد", "grade": "الثالث المتوسط", "class": 2, "phone": "966555992829"},
    {"id": "2310646332", "name": "راكان ابراهيم محمد ديوان", "grade": "الثالث المتوسط", "class": 2, "phone": "966500030732"},
    {"id": "1161397599", "name": "ريان ناصر عبدالرحمن المرشود", "grade": "الثالث المتوسط", "class": 2, "phone": "966550666662"},
    {"id": "1163112129", "name": "صالح بن ممدوح بن صالح بن خالد الجويعي", "grade": "الثالث المتوسط", "class": 2, "phone": "966549887719"},
    {"id": "2508581135", "name": "عبد الرحمن محمد صلاح السيد بدر الدين", "grade": "الثالث المتوسط", "class": 2, "phone": "966507652707"},
    {"id": "1162188872", "name": "عبدالعزيز تركي عبدالعزيز اللهيم", "grade": "الثالث المتوسط", "class": 2, "phone": "966505256806"},
    {"id": "1161340763", "name": "عبدالعزيز عبدالمحسن فهد بن بديع", "grade": "الثالث المتوسط", "class": 2, "phone": "966554457163"},
    {"id": "1171845140", "name": "عبدالله متعب بن عبدالرحمن الجبرين", "grade": "الثالث المتوسط", "class": 2, "phone": "966559898559"},
    {"id": "1159200318", "name": "عبدالمحسن طارق بن عبدالرحمن العروان", "grade": "الثالث المتوسط", "class": 2, "phone": "966506291294"},
    {"id": "1162454266", "name": "عمر فهد محمد السقامي", "grade": "الثالث المتوسط", "class": 2, "phone": "966564234552"},
    {"id": "1165152107", "name": "فيصل محمد صالح الفنتوخ", "grade": "الثالث المتوسط", "class": 2, "phone": "966556488802"},
    {"id": "1162325722", "name": "ماجد فهد عبدالعزيز الكثيري", "grade": "الثالث المتوسط", "class": 2, "phone": "966557609015"},
    {"id": "1162461857", "name": "محمد خالد محمد بن مشرف", "grade": "الثالث المتوسط", "class": 2, "phone": "966551777559"},
    {"id": "1161288897", "name": "محمد سعد بن محمد العيشان", "grade": "الثالث المتوسط", "class": 2, "phone": "966504217660"},
    {"id": "1156334813", "name": "محمد عبدالعزيز محمد الخالدي", "grade": "الثالث المتوسط", "class": 2, "phone": "966500091387"},
    {"id": "1162044851", "name": "مهند ماجد علي كعبي", "grade": "الثالث المتوسط", "class": 2, "phone": "966533313738"},
    {"id": "1158021137", "name": "ناصر محمد عبدالله الزريعي", "grade": "الثالث المتوسط", "class": 2, "phone": "966505231121"},
    {"id": "1161363443", "name": "نواف سعد بن علي القاسم", "grade": "الثالث المتوسط", "class": 2, "phone": "966504200199"},
    {"id": "1162274086", "name": "ياسر تركي اسماعيل مسملي", "grade": "الثالث المتوسط", "class": 2, "phone": "966504261855"},

    # الصف الثالث المتوسط - فصل 3
    {"id": "1163525544", "name": "ثامر وليد بن عبدالعزيز الطليحي", "grade": "الثالث المتوسط", "class": 3, "phone": "966504437710"},
    {"id": "1160712996", "name": "خالد بن عبدالرؤوف بن عبدالله الشنيبر", "grade": "الثالث المتوسط", "class": 3, "phone": "966504173163"},
    {"id": "1162560054", "name": "خالد عبدالله خالد الخالدي", "grade": "الثالث المتوسط", "class": 3, "phone": "966558890881"},
    {"id": "1174188647", "name": "خالد محمد بن عبدالله ال درعان", "grade": "الثالث المتوسط", "class": 3, "phone": "966505556029"},
    {"id": "1159155223", "name": "راشد سعيد راشد عبدالسلام", "grade": "الثالث المتوسط", "class": 3, "phone": "966533177877"},
    {"id": "1174226389", "name": "راشد صالح بن عبدالعزيز الحلوان", "grade": "الثالث المتوسط", "class": 3, "phone": "966551112126"},
    {"id": "1167756897", "name": "رواد محمد ابراهيم الخليل", "grade": "الثالث المتوسط", "class": 3, "phone": "966502555411"},
    {"id": "1159394046", "name": "صالح بن محمد بن صالح الميموني المطيري", "grade": "الثالث المتوسط", "class": 3, "phone": "966555097811"},
    {"id": "1158551372", "name": "عبدالرحمن بدر عبدالرحمن الطريقي", "grade": "الثالث المتوسط", "class": 3, "phone": "966507004114"},
    {"id": "1195815558", "name": "عبدالرحمن خالد محمد سعيد", "grade": "الثالث المتوسط", "class": 3, "phone": "966504411393"},
    {"id": "1158561843", "name": "عبدالله تركي عبدالله الأحمد", "grade": "الثالث المتوسط", "class": 3, "phone": "966542800700"},
    {"id": "1159977451", "name": "عبدالله عبدالرحمن عبدالله النجراني", "grade": "الثالث المتوسط", "class": 3, "phone": "966546416395"},
    {"id": "1162387458", "name": "علي بن خالد بن علي العجيري", "grade": "الثالث المتوسط", "class": 3, "phone": "966505199500"},
    {"id": "1158128270", "name": "علي عبدالله علي ال حمود", "grade": "الثالث المتوسط", "class": 3, "phone": "966545555161"},
    {"id": "1161333677", "name": "فارس وليد بن عبدالله الحوطي", "grade": "الثالث المتوسط", "class": 3, "phone": "966552805550"},
    {"id": "1158198604", "name": "فهد بن خالد بن فهد بن عبدالعزيز الزيد", "grade": "الثالث المتوسط", "class": 3, "phone": "966555198633"},
    {"id": "1159551264", "name": "فيصل عبدالرحمن عزيز القحطاني", "grade": "الثالث المتوسط", "class": 3, "phone": "966556444082"},
    {"id": "1186515613", "name": "متعب مطر جمعان الدوسري", "grade": "الثالث المتوسط", "class": 3, "phone": "966530545913"},
    {"id": "1159852746", "name": "نواف فهد بن ناصر القحطاني", "grade": "الثالث المتوسط", "class": 3, "phone": "966556557210"},
    {"id": "1163072392", "name": "يوسف عبدالله عوض العتيبي", "grade": "الثالث المتوسط", "class": 3, "phone": "966506371377"}
]

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS students (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            grade TEXT NOT NULL,
            class INT NOT NULL,
            phone TEXT NOT NULL
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS grades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id TEXT NOT NULL,
            term TEXT NOT NULL,
            week TEXT NOT NULL,
            score REAL,
            is_absent INT DEFAULT 0,
            UNIQUE(student_id, term, week),
            FOREIGN KEY (student_id) REFERENCES students(id)
        )
    ''')
    c.execute("SELECT COUNT(*) FROM students")
    count = c.fetchone()[0]
    if count == 0:
        for st_data in STUDENTS_INIT:
            c.execute("INSERT OR REPLACE INTO students (id, name, grade, class, phone) VALUES (?, ?, ?, ?, ?)",
                      (st_data["id"], st_data["name"], st_data["grade"], st_data["class"], st_data["phone"]))
    conn.commit()
    conn.close()

init_db()

# ==============================================================================
# 2. الدوال المساعدة وصياغة الرسائل
# ==============================================================================

def create_whatsapp_url(phone, text):
    phone_clean = str(phone).strip().replace("+", "").replace(" ", "").replace("-", "")
    if phone_clean.startswith("05"):
        phone_clean = "966" + phone_clean[1:]
    elif phone_clean.startswith("5"):
        phone_clean = "966" + phone_clean
    encoded_text = urllib.parse.quote(text)
    return f"https://api.whatsapp.com/send?phone={phone_clean}&text={encoded_text}"

def generate_parent_message(student_name, score, is_absent):
    if is_absent:
        return (
            f"المكرم ولي أمر الطالب/ {student_name}، نود التنبيه على غياب الطالب هذا الأسبوع، "
            f"ونحثكم على متابعة الانتظام وحضور الاختبارات لتجنب حسم الدرجات والتأثير على مستواه التحصيلي: متوسطة الثغر النموذجية الأهلية."
        )
    
    sc_str = f"{score}%" if score is not None else "أقل من 50%"

    if score is None or score < 50:
        return (
            f"المكرم ولي أمر الطالب/ {student_name}، نفيدكم بأن نسبة إتقان الطالب هذا الأسبوع هي ({sc_str}). "
            f"حرصاً منا على مصلحة ابنكم ومستقبله الدراسي، نود إشعاركم بوجود تراجع ملحوظ في مستواه التحصيلي مؤخراً، "
            f"ونرجو منكم تكثيف المتابعة المنزلية والتواصل معنا للوقوف على أسباب هذا التراجع ووضع خطة لتحسين أدائه. مع تحياتنا متوسطة الثغر النموذجية الأهلية."
        )
    elif score <= 75:
        return (
            f"المكرم ولي أمر الطالب/ {student_name}، نفيدكم بأن نسبة إتقان الطالب هذا الأسبوع هي ({sc_str}). "
            f"نود إحاطتكم علماً بأن المستوى التحصيلي لابنكم جيد ومستقر بشكل عام، ولكنه يمتلك قدرات أعلى تؤهله لتحقيق درجات أفضل. "
            f"نأمل منكم التركيز معه في الفترة القادمة لرفع كفاءته الدراسية. شاكرين لكم تعاونكم الدائم. مع تحياتنا متوسطة الثغر النموذجية الأهلية."
        )
    else:
        return (
            f"المكرم ولي أمر الطالب/ {student_name}، نتقدم بخالص الشكر والتقدير لكم وللطالب على الاهتمام والتفوق بنسبة إتقان ممتازة ({sc_str})، "
            f"يسعدنا إبلاغكم بأن ابنكم قدم أداءً تحصيلياً متميزاً وسلوكاً رائعاً داخل الفصل، وحصل على درجات ممتازة في التقييمات الأخيرة. "
            f"نشكر لكم حسن المتابعة والاهتمام، ونرجو الاستمرار في هذا الدعم المتبادل للحفاظ على هذا المستوى المتفوق. مع تحياتنا متوسطة الثغر النموذجية الأهلية."
        )

def render_printable_html_view(html_content, title="طباعة التقرير"):
    full_html = f'''<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
<meta charset="utf-8">
<title>{title}</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700&display=swap');
html, body {{
    font-family: 'Tajawal', sans-serif !important;
    direction: rtl !important;
    text-align: right !important;
    background-color: #f8fafc;
    margin: 0;
    padding: 15px;
}}
.print-container {{
    background: white;
    padding: 25px;
    border-radius: 12px;
    box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    max-width: 900px;
    margin: 0 auto;
}}
.print-btn {{
    background-color: #1e3c72;
    color: white;
    border: none;
    padding: 12px 24px;
    border-radius: 8px;
    font-size: 16px;
    font-weight: bold;
    cursor: pointer;
    margin-bottom: 20px;
    width: 100%;
}}
.header-table {{
    width: 100%;
    margin-bottom: 25px;
    border-bottom: 2px solid #1e3c72;
    padding-bottom: 15px;
}}
.data-table {{
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 25px;
}}
.data-table th, .data-table td {{
    border: 1px solid #cbd5e1;
    padding: 10px;
    text-align: center;
}}
.data-table th {{
    background-color: #1e3c72;
    color: white;
}}
.signatures {{
    width: 100%;
    margin-top: 30px;
    text-align: center;
}}
.badge-green {{ background-color: #d1fae5; color: #065f46; padding: 4px 8px; border-radius: 4px; font-weight: bold; }}
.badge-blue {{ background-color: #dbeafe; color: #1e40af; padding: 4px 8px; border-radius: 4px; font-weight: bold; }}
.badge-red {{ background-color: #fee2e2; color: #991b1b; padding: 4px 8px; border-radius: 4px; font-weight: bold; }}
.badge-gray {{ background-color: #f3f4f6; color: #374151; padding: 4px 8px; border-radius: 4px; font-weight: bold; }}
@media print {{
    .no-print {{ display: none !important; }}
    body {{ background: white; padding: 0; }}
    .print-container {{ box-shadow: none; padding: 0; }}
}}
</style>
</head>
<body>
<div class="print-container">
    <button class="print-btn no-print" onclick="window.print()">🖨️ اضغط هنا للطباعة المباشرة / التصدير كـ PDF</button>
    {html_content}
</div>
</body>
</html>'''
    components.html(full_html, height=650, scrolling=True)

# ==============================================================================
# 3. واجهة برنامج Streamlit
# ==============================================================================
st.set_page_config(
    page_title="برنامج رصد الدرجات - متوسطة الثغر النموذجية الأهلية",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown('''
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700&display=swap');
html, body, [class*="css"], .stMarkdown, .stText, div[data-baseweb="input"], div[data-baseweb="select"], .stNumberInput input {
    font-family: 'Tajawal', sans-serif !important;
    direction: rtl !important;
    text-align: right !important;
}
.stApp {
    background-color: #f8fafc;
}
.header-box {
    background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
    color: white;
    padding: 22px;
    border-radius: 14px;
    text-align: center !important;
    margin-bottom: 25px;
    box-shadow: 0 4px 18px rgba(0,0,0,0.12);
}
.student-name-box {
    text-align: right !important;
    direction: rtl !important;
    font-weight: 700;
    font-size: 16px;
    color: #1e3c72;
    padding: 6px 10px;
    background-color: #ffffff;
    border-right: 4px solid #2a5298;
    border-radius: 6px;
    margin-bottom: 8px;
}
.stDataFrame table, .stDataFrame td, .stDataFrame th {
    text-align: right !important;
    direction: rtl !important;
}
div[data-testid="stForm"] {
    border-radius: 12px;
    border: 1px solid #e2e8f0;
    padding: 20px;
    background-color: #ffffff;
}
</style>
''', unsafe_allow_html=True)

st.markdown('''
<div class="header-box">
    <h2 style="margin:0;">🏫 نظام رصد الدرجات والتواصل مع أولياء الأمور</h2>
    <p style="margin:5px 0 0 0; font-size:16px;">متوسطة الثغر النموذجية الأهلية (بنين) بالرياض</p>
</div>
''', unsafe_allow_html=True)

st.sidebar.title("📌 القائمة الرئيسية")
page = st.sidebar.radio("اختر الصفحة:", ["📝 صفحة الرصد", "🏫 إدارة المدرسة وتقارير أولياء الأمور"])

# ==============================================================================
# الصفحة الأولى: صفحة الرصد (RECORDING SHEET)
# ==============================================================================
if page == "📝 صفحة الرصد":
    st.subheader("📝 صفحة رصد درجات الإتقان الأسبوعية")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        term = st.selectbox("الفصل الدراسي:", ["الفصل الدراسي الأول", "الفصل الدراسي الثاني"])
    with col2:
        grade = st.selectbox("الصف الدراسي:", ["الأول المتوسط", "الثاني المتوسط", "الثالث المتوسط"])
    with col3:
        class_num = st.selectbox("الفصل / الشعبة:", [1, 2, 3])
    with col4:
        weeks = [f"الأسبوع {i}" for i in range(1, 19)]
        week = st.selectbox("الأسبوع المستهدف:", weeks)

    st.markdown("---")

    conn = get_db_connection()
    query = """
        SELECT s.id, s.name, s.grade, s.class, g.score, g.is_absent 
        FROM students s
        LEFT JOIN grades g ON s.id = g.student_id AND g.term = ? AND g.week = ?
        WHERE s.grade = ? AND s.class = ?
        ORDER BY s.name ASC
    """
    df_students = pd.read_sql_query(query, conn, params=(term, week, grade, class_num))
    conn.close()

    if df_students.empty:
        st.warning("لا يوجد طلاب مسجلين في هذا الصف والشعبة.")
    else:
        st.info(f"📊 عدد الطلاب في {grade} - فصل ({class_num}): **{len(df_students)} طالب** | {term} - {week}")
        
        with st.form("recording_form"):
            st.markdown("##### 📥 أدخل/عدّل درجات الطلاب وحالة الغياب:")
            
            updated_data = []
            for idx, row in df_students.iterrows():
                col_name, col_score, col_absent = st.columns([3, 2, 1])
                with col_name:
                    st.markdown(f'<div class="student-name-box">📌 {idx+1}. {row["name"]} <span style="font-size: 12px; color: #64748b; font-weight: normal;">({row["id"]})</span></div>', unsafe_allow_html=True)
                with col_score:
                    curr_score = float(row["score"]) if (pd.notnull(row["score"]) and row["score"] is not None) else 0.0
                    sc = st.number_input(f"الدرجة (100)", min_value=0.0, max_value=100.0, value=curr_score, step=1.0, key=f"sc_{row['id']}")
                with col_absent:
                    curr_abs = bool(row["is_absent"]) if pd.notnull(row["is_absent"]) else False
                    is_abs = st.checkbox("غائب ⚪", value=curr_abs, key=f"abs_{row['id']}")
                
                final_score = 0.0 if is_abs else sc
                updated_data.append({
                    "student_id": row["id"],
                    "name": row["name"],
                    "score": final_score,
                    "is_absent": 1 if is_abs else 0
                })
            
            save_btn = st.form_submit_button("💾 حفظ البيانات والتحديث")
            
        if save_btn:
            conn = get_db_connection()
            c = conn.cursor()
            for item in updated_data:
                c.execute("""
                    INSERT INTO grades (student_id, term, week, score, is_absent)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(student_id, term, week) DO UPDATE SET
                    score = excluded.score,
                    is_absent = excluded.is_absent
                """, (item["student_id"], term, week, item["score"], item["is_absent"]))
            conn.commit()
            conn.close()
            st.success("✅ تم حفظ البيانات وتحديث جدول الرصد وقاعدة البيانات بنجاح!")
            st.rerun()

        st.markdown("### 📊 جدول نتائج الرصد المنسق بالتلوين الشرطي:")
        
        table_rows = []
        for item in updated_data:
            sc = item["score"]
            is_abs = item["is_absent"]
            if is_abs == 1:
                cat = "غائب ⚪"
                pct = "0% (غائب)"
            elif sc < 50:
                cat = "أقل من 50% (ضعيف) 🔴"
                pct = f"{sc}%"
            elif sc <= 75:
                cat = "50% - 75% (متوسط) 🔵"
                pct = f"{sc}%"
            else:
                cat = "76% - 100% (ممتاز) 🟢"
                pct = f"{sc}%"
                
            table_rows.append({
                "اسم الطالب": item["name"],
                "رقم الهوية": item["student_id"],
                "درجة الإتقان / 100": sc if is_abs == 0 else 0.0,
                "النسبة المئوية": pct,
                "الفئة / الحالة": cat
            })
        
        df_display = pd.DataFrame(table_rows)
        st.dataframe(df_display, use_container_width=True)

        st.markdown("---")
        st.markdown("### 🖨️ طباعة وتصدير كشوف الرصد:")
        
        tab_exp1, tab_exp2 = st.tabs(["🖨️ معاينة وطباعة كشف مفرغ (HTML / PDF)", "📊 تصدير إلى Excel"])
        
        with tab_exp1:
            st.markdown("#### 📝 معاينة وطباعة كشف رصد مفرغ للرصد اليدوي:")
            
            rows_html = ""
            for i, st_item in enumerate(updated_data, 1):
                rows_html += f"""
                <tr>
                    <td>{i}</td>
                    <td>{st_item['student_id']}</td>
                    <td style="text-align: right; padding-right: 15px;">{st_item['name']}</td>
                    <td></td>
                    <td></td>
                    <td></td>
                </tr>
                """
                
            blank_sheet_html = f"""
            <table class="header-table">
                <tr>
                    <td style="text-align: right; width: 35%;">
                        <b>المملكة العربية السعودية</b><br>
                        <b>وزارة التعليم</b><br>
                        <b>الإدارة العامة للتعليم بمنطقة الرياض</b><br>
                        <b>متوسطة الثغر النموذجية الأهلية (بنين)</b>
                    </td>
                    <td style="text-align: center; width: 30%;">
                        <h3 style="margin: 0; color: #1e3c72;">كشف رصد مفرغ للرصد اليدوي</h3>
                        <p style="margin: 5px 0;">درجات الإتقان الأسبوعية</p>
                    </td>
                    <td style="text-align: left; width: 35%;">
                        <b>الصف الدراسي:</b> {grade}<br>
                        <b>الفصل / الشعبة:</b> ({class_num})<br>
                        <b>الفصل الدراسي:</b> {term}<br>
                        <b>الأسبوع:</b> {week}
                    </td>
                </tr>
            </table>

            <table class="data-table">
                <thead>
                    <tr>
                        <th style="width: 5%;">م</th>
                        <th style="width: 20%;">رقم الهوية</th>
                        <th style="width: 40%;">اسم الطالب</th>
                        <th style="width: 12%;">درجة الإتقان / 100</th>
                        <th style="width: 10%;">النسبة %</th>
                        <th style="width: 13%;">ملاحظات</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                </tbody>
            </table>

            <table class="signatures">
                <tr>
                    <td>وكيل الشؤون التعليمية<br><br><b>محمد مبروك السيد</b></td>
                    <td>وكيل شؤون الطلاب<br><br><b>صالح بن عبدالله الدعجاني</b></td>
                    <td>مدير المدرسة<br><br><b>إبراهيم بن موسى التميمي</b></td>
                </tr>
            </table>
            <div style="text-align: center; margin-top: 15px; font-size: 12px; color: #777;">
                تصميم وتطوير: <b>محمد سامي السعيد</b>
            </div>
            """
            
            render_printable_html_view(blank_sheet_html, title="كشف_رصد_مفرغ")

        with tab_exp2:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "كشف_الرصد"
            ws.views.sheetView[0].rightToLeft = True
            
            headers = ["م", "رقم الهوية", "اسم الطالب", "الدرجة / 100", "النسبة المئوية", "الحالة"]
            ws.append(headers)
            
            for i, r in enumerate(table_rows, 1):
                ws.append([i, r["رقم الهوية"], r["اسم الطالب"], r["درجة الإتقان / 100"], r["النسبة المئوية"], r["الفئة / الحالة"]])
                
            excel_io = io.BytesIO()
            wb.save(excel_io)
            excel_io.seek(0)
            
            st.download_button(
                label="📊 تصدير كشف الدرجات (Excel)",
                data=excel_io,
                file_name=f"كشف_درجات_{grade}_{class_num}_{week}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ==============================================================================
# الصفحة الثانية: إدارة المدرسة وتقارير أولياء الأمور (ADMIN & PARENT REPORTS)
# ==============================================================================
elif page == "🏫 إدارة المدرسة وتقارير أولياء الأمور":
    st.subheader("🏫 إدارة المدرسة وإرسال وتقارير أولياء الأمور")
    
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        weeks = [f"الأسبوع {i}" for i in range(1, 19)]
        selected_week = st.selectbox("اختر الأسبوع لعرض التقرير والرسائل:", weeks)
    with col_w2:
        selected_term = st.selectbox("اختر الفصل الدراسي:", ["الفصل الدراسي الأول", "الفصل الدراسي الثاني"])
        
    st.markdown("---")

    conn = get_db_connection()
    query = """
        SELECT s.id, s.name, s.grade, s.class, s.phone, g.score, g.is_absent
        FROM students s
        LEFT JOIN grades g ON s.id = g.student_id AND g.term = ? AND g.week = ?
        ORDER BY s.grade, s.class, s.name
    """
    df_reports = pd.read_sql_query(query, conn, params=(selected_term, selected_week))
    conn.close()

    cat_red, cat_blue, cat_green, cat_gray = [], [], [], []

    for idx, r in df_reports.iterrows():
        sc = r["score"]
        is_abs = r["is_absent"]
        msg = generate_parent_message(r["name"], sc, is_abs)
        
        row_dict = {
            "id": r["id"],
            "name": r["name"],
            "grade": r["grade"],
            "class": r["class"],
            "phone": r["phone"],
            "score": sc if (sc is not None and is_abs == 0) else 0.0,
            "is_absent": is_abs,
            "message": msg
        }
        
        if is_abs == 1:
            cat_gray.append(row_dict)
        elif sc is None or sc < 50:
            cat_red.append(row_dict)
        elif sc <= 75:
            cat_blue.append(row_dict)
        else:
            cat_green.append(row_dict)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("🔴 فئة أقل من 50%", f"{len(cat_red)} طالب")
    m2.metric("🔵 فئة 50% - 75%", f"{len(cat_blue)} طالب")
    m3.metric("🟢 فئة 76% - 100%", f"{len(cat_green)} طالب")
    m4.metric("⚪ فئة الغياب", f"{len(cat_gray)} طالب")

    st.markdown("---")

    st.markdown("##### 🟢 خيارات إرسال الرسائل عبر الواتساب (WhatsApp):")
    col_wa1, col_wa2 = st.columns(2)
    with col_wa1:
        st.success("✅ يتم توليد رابط واتساب مباشر لكل طالب يتضمن رقم جوال ولي الأمر ونص الرسالة المخصص تلقائياً.")
    with col_wa2:
        st.info("💡 اضغط على زر **'إرسال الرسالة الآن عبر الواتساب'** الأخضر أمام كل طالب لفتح محادثة الواتساب فوراً مع الرسالة المجهزة.")

    st.markdown("### 📋 تفاصيل الرسائل النصية الموجهة حسب الفئات:")

    tab1, tab2, tab3, tab4 = st.tabs([
        f"🔴 أقل من 50% ({len(cat_red)})",
        f"🔵 50% - 75% ({len(cat_blue)})",
        f"🟢 76% - 100% ({len(cat_green)})",
        f"⚪ الغياب ({len(cat_gray)})"
    ])

    def show_category_tab(cat_list, cat_name):
        if not cat_list:
            st.info(f"لا يوجد طلاب في {cat_name} بهذا الأسبوع.")
        else:
            st.markdown(f"##### 📲 قائمة رسائل {cat_name} الموجهة لولي الأمر عبر الواتساب:")
            for item in cat_list:
                wa_link = create_whatsapp_url(item['phone'], item['message'])
                with st.expander(f"👤 {item['name']} ({item['grade']} - فصل {item['class']}) | جوال ولي الأمر: {item['phone']}"):
                    st.write(f"**رقم الهوية:** {item['id']}")
                    st.write(f"**النسبة المئوية / الدرجة:** {item['score']}%" if item['is_absent'] == 0 else "**الحالة:** غائب ⚪")
                    st.info(f"💬 **نص الرسالة الموجهة:**\n\n{item['message']}")
                    st.markdown(f'''
                    <a href="{wa_link}" target="_blank" style="text-decoration:none;">
                        <div style="background-color:#25D366; color:white; padding:12px 18px; border-radius:8px; text-align:center; font-weight:bold; font-size:15px; margin-top:8px; display:block;">
                            💬 إرسال الرسالة الآن عبر الواتساب (WhatsApp) إلى ولي الأمر ({item['phone']})
                        </div>
                    </a>
                    ''', unsafe_allow_html=True)

    with tab1: show_category_tab(cat_red, "فئة أقل من 50%")
    with tab2: show_category_tab(cat_blue, "فئة 50% - 75%")
    with tab3: show_category_tab(cat_green, "فئة 76% - 100%")
    with tab4: show_category_tab(cat_gray, "فئة الغياب")

    st.markdown("---")

    # ==============================================================================
    # قسم طباعة وتصدير التقارير الرسمية لأولياء الأمور والإدارة
    # ==============================================================================
    st.markdown("### 🖨️ قسم طباعة وتصدير التقارير الرسمية لأولياء الأمور والإدارة:")

    print_type = st.radio("اختر نوع التقرير المراد طباعته وتصديره:", [
        "👤 طباعة تقرير طالب محدد",
        "🏫 طباعة تقرير صف بالكامل (جميع الفصول)",
        "🎒 طباعة تقرير فصل / شعبة محددة"
    ])

    if print_type == "👤 طباعة تقرير طالب محدد":
        st.markdown("#### 👤 طباعة وتصدير تقرير فردي لطالب:")
        selected_student_name = st.selectbox("اختر اسم الطالب:", df_reports["name"].tolist())
        st_info = df_reports[df_reports["name"] == selected_student_name].iloc[0]

        # 1. احتساب الدرجة وحالة الغياب وتوليد نص الرسالة أولاً لتجنب NameError
        sc = st_info["score"]
        is_abs = st_info["is_absent"]
        msg = generate_parent_message(st_info["name"], sc, is_abs)

        # 2. توليد رابط الواتساب بعد التأكد من وجود المتغير msg
        wa_indiv_url = create_whatsapp_url(st_info['phone'], msg)
        st.markdown(f'''
        <a href="{wa_indiv_url}" target="_blank" style="text-decoration:none;">
            <div style="background-color:#25D366; color:white; padding:14px 20px; border-radius:8px; text-align:center; font-weight:bold; font-size:16px; margin-bottom:15px; display:block;">
                💬 إرسال التقرير والرسالة فوراً إلى ولي الأمر عبر الواتساب (WhatsApp) -> {st_info['phone']}
            </div>
        </a>
        ''', unsafe_allow_html=True)

        if is_abs == 1:
            badge_class = "badge-gray"
            badge_text = "غائب ⚪"
            score_display = "0% (غائب)"
        elif sc is None or sc < 50:
            badge_class = "badge-red"
            badge_text = "أقل من 50% (يحتاج متابعة) 🔴"
            score_display = f"{sc}%"
        elif sc <= 75:
            badge_class = "badge-blue"
            badge_text = "50% - 75% (مستوى متوسط) 🔵"
            score_display = f"{sc}%"
        else:
            badge_class = "badge-green"
            badge_text = "76% - 100% (ممتاز ومتفوق) 🟢"
            score_display = f"{sc}%"

        student_report_html = f"""
        <table class="header-table">
            <tr>
                <td style="text-align: right; width: 35%;">
                    <b>المملكة العربية السعودية</b><br>
                    <b>وزارة التعليم</b><br>
                    <b>الإدارة العامة للتعليم بمنطقة الرياض</b><br>
                    <b>متوسطة الثغر النموذجية الأهلية (بنين)</b>
                </td>
                <td style="text-align: center; width: 30%;">
                    <h3 style="margin: 0; color: #1e3c72;">تقرير ولي الأمر لدرجات الإتقان</h3>
                    <p style="margin: 5px 0;">{selected_term} - {selected_week}</p>
                </td>
                <td style="text-align: left; width: 35%;">
                    <b>التاريخ:</b> 1447/1448 هـ<br>
                    <b>رقم السجل:</b> {st_info['id']}
                </td>
            </tr>
        </table>

        <div style="background: #fdfdfd; border: 1px solid #1e3c72; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            <table style="width: 100%; border-collapse: collapse;">
                <tr>
                    <td style="padding: 8px;"><b>اسم الطالب:</b> {st_info['name']}</td>
                    <td style="padding: 8px;"><b>رقم الهوية:</b> {st_info['id']}</td>
                </tr>
                <tr>
                    <td style="padding: 8px;"><b>الصف الدراسي:</b> {st_info['grade']}</td>
                    <td style="padding: 8px;"><b>الفصل / الشعبة:</b> ({st_info['class']})</td>
                </tr>
                <tr>
                    <td style="padding: 8px;"><b>نسبة الإتقان الأسبوعية:</b> {score_display}</td>
                    <td style="padding: 8px;"><b>مستوى الطالب:</b> <span class="{badge_class}">{badge_text}</span></td>
                </tr>
            </table>
        </div>

        <div style="background: #eef2f5; border-right: 5px solid #1e3c72; padding: 15px; border-radius: 4px; margin-bottom: 25px;">
            <h4 style="margin-top: 0; color: #1e3c72;">💬 الرسالة الموجهة لولي الأمر:</h4>
            <p style="font-size: 15px; line-height: 1.6; margin: 0;">{msg}</p>
        </div>

        <table class="signatures">
            <tr>
                <td>وكيل الشؤون التعليمية<br><br><b>محمد مبروك السيد</b></td>
                <td>وكيل شؤون الطلاب<br><br><b>صالح بن عبدالله الدعجاني</b></td>
                <td>مدير المدرسة<br><br><b>إبراهيم بن موسى التميمي</b></td>
            </tr>
        </table>
        <div style="text-align: center; margin-top: 15px; font-size: 12px; color: #777;">
            تصميم وتطوير: <b>محمد سامي السعيد</b>
        </div>
        """
        
        render_printable_html_view(student_report_html, title=f"تقرير_{st_info['name']}")

    elif print_type == "🏫 طباعة تقرير صف بالكامل (جميع الفصول)":
        st.markdown("#### 🏫 طباعة وتصدير تقرير صف دراسي كامل:")
        selected_grade_print = st.selectbox("اختر الصف الدراسي:", ["الأول المتوسط", "الثاني المتوسط", "الثالث المتوسط"])
        
        df_grade = df_reports[df_reports["grade"] == selected_grade_print]
        
        grade_rows_html = ""
        for idx, r in df_grade.iterrows():
            sc = r["score"]
            is_abs = r["is_absent"]
            if is_abs == 1:
                b_class = "badge-gray"
                b_text = "غائب ⚪"
                sc_str = "0% (غائب)"
            elif sc is None or sc < 50:
                b_class = "badge-red"
                b_text = "أقل من 50% 🔴"
                sc_str = f"{sc}%"
            elif sc <= 75:
                b_class = "badge-blue"
                b_text = "50% - 75% 🔵"
                sc_str = f"{sc}%"
            else:
                b_class = "badge-green"
                b_text = "76% - 100% 🟢"
                sc_str = f"{sc}%"
                
            msg_short = generate_parent_message(r["name"], sc, is_abs)
            
            grade_rows_html += f"""
            <tr>
                <td>{r['id']}</td>
                <td style="text-align: right; padding-right: 10px;"><b>{r['name']}</b></td>
                <td>فصل ({r['class']})</td>
                <td><b>{sc_str}</b></td>
                <td><span class="{b_class}">{b_text}</span></td>
                <td style="text-align: right; font-size: 12px; padding: 6px;">{msg_short}</td>
            </tr>
            """

        grade_report_html = f"""
        <table class="header-table">
            <tr>
                <td style="text-align: right; width: 35%;">
                    <b>المملكة العربية السعودية</b><br>
                    <b>وزارة التعليم</b><br>
                    <b>الإدارة العامة للتعليم بمنطقة الرياض</b><br>
                    <b>متوسطة الثغر النموذجية الأهلية (بنين)</b>
                </td>
                <td style="text-align: center; width: 30%;">
                    <h3 style="margin: 0; color: #1e3c72;">تقرير الإتقان لصف {selected_grade_print} كاملاً</h3>
                    <p style="margin: 5px 0;">{selected_term} - {selected_week}</p>
                </td>
                <td style="text-align: left; width: 35%;">
                    <b>الصف الدراسي:</b> {selected_grade_print}<br>
                    <b>إجمالي الطلاب:</b> {len(df_grade)} طالب
                </td>
            </tr>
        </table>

        <table class="data-table">
            <thead>
                <tr>
                    <th style="width: 15%;">رقم الهوية</th>
                    <th style="width: 25%;">اسم الطالب</th>
                    <th style="width: 10%;">الفصل</th>
                    <th style="width: 10%;">النسبة %</th>
                    <th style="width: 15%;">الحالة</th>
                    <th style="width: 25%;">نص الرسالة الموجهة لولي الأمر</th>
                </tr>
            </thead>
            <tbody>
                {grade_rows_html}
            </tbody>
        </table>

        <table class="signatures">
            <tr>
                <td>وكيل الشؤون التعليمية<br><br><b>محمد مبروك السيد</b></td>
                <td>وكيل شؤون الطلاب<br><br><b>صالح بن عبدالله الدعجاني</b></td>
                <td>مدير المدرسة<br><br><b>إبراهيم بن موسى التميمي</b></td>
            </tr>
        </table>
        <div style="text-align: center; margin-top: 15px; font-size: 12px; color: #777;">
            تصميم وتطوير: <b>محمد سامي السعيد</b>
        </div>
        """
        
        render_printable_html_view(grade_report_html, title=f"تقرير_{selected_grade_print}")

    elif print_type == "🎒 طباعة تقرير فصل / شعبة محددة":
        st.markdown("#### 🎒 طباعة وتصدير تقرير شعبة / فصل محدد:")
        col_g, col_c = st.columns(2)
        with col_g:
            gr_select = st.selectbox("اختر الصف الدراسي:", ["الأول المتوسط", "الثاني المتوسط", "الثالث المتوسط"], key="gr_sec")
        with col_c:
            cl_select = st.selectbox("اختر الفصل / الشعبة:", [1, 2, 3], key="cl_sec")
            
        df_class = df_reports[(df_reports["grade"] == gr_select) & (df_reports["class"] == cl_select)]
        
        class_rows_html = ""
        for idx, r in df_class.iterrows():
            sc = r["score"]
            is_abs = r["is_absent"]
            if is_abs == 1:
                b_class = "badge-gray"
                b_text = "غائب ⚪"
                sc_str = "0% (غائب)"
            elif sc is None or sc < 50:
                b_class = "badge-red"
                b_text = "أقل من 50% 🔴"
                sc_str = f"{sc}%"
            elif sc <= 75:
                b_class = "badge-blue"
                b_text = "50% - 75% 🔵"
                sc_str = f"{sc}%"
            else:
                b_class = "badge-green"
                b_text = "76% - 100% 🟢"
                sc_str = f"{sc}%"
                
            msg_short = generate_parent_message(r["name"], sc, is_abs)
            
            class_rows_html += f"""
            <tr>
                <td>{r['id']}</td>
                <td style="text-align: right; padding-right: 10px;"><b>{r['name']}</b></td>
                <td><b>{sc_str}</b></td>
                <td><span class="{b_class}">{b_text}</span></td>
                <td style="text-align: right; font-size: 12px; padding: 6px;">{msg_short}</td>
            </tr>
            """

        class_report_html = f"""
        <table class="header-table">
            <tr>
                <td style="text-align: right; width: 35%;">
                    <b>المملكة العربية السعودية</b><br>
                    <b>وزارة التعليم</b><br>
                    <b>الإدارة العامة للتعليم بمنطقة الرياض</b><br>
                    <b>متوسطة الثغر النموذجية الأهلية (بنين)</b>
                </td>
                <td style="text-align: center; width: 30%;">
                    <h3 style="margin: 0; color: #1e3c72;">تقرير شعبة {gr_select} - فصل ({cl_select})</h3>
                    <p style="margin: 5px 0;">{selected_term} - {selected_week}</p>
                </td>
                <td style="text-align: left; width: 35%;">
                    <b>الصف الدراسي:</b> {gr_select}<br>
                    <b>الشعبة / الفصل:</b> ({cl_select})<br>
                    <b>عدد الطلاب:</b> {len(df_class)} طالب
                </td>
            </tr>
        </table>

        <table class="data-table">
            <thead>
                <tr>
                    <th style="width: 18%;">رقم الهوية</th>
                    <th style="width: 27%;">اسم الطالب</th>
                    <th style="width: 12%;">النسبة %</th>
                    <th style="width: 15%;">الحالة</th>
                    <th style="width: 28%;">نص الرسالة الموجهة لولي الأمر</th>
                </tr>
            </thead>
            <tbody>
                {class_rows_html}
            </tbody>
        </table>

        <table class="signatures">
            <tr>
                <td>وكيل الشؤون التعليمية<br><br><b>محمد مبروك السيد</b></td>
                <td>وكيل شؤون الطلاب<br><br><b>صالح بن عبدالله الدعجاني</b></td>
                <td>مدير المدرسة<br><br><b>إبراهيم بن موسى التميمي</b></td>
            </tr>
        </table>
        <div style="text-align: center; margin-top: 15px; font-size: 12px; color: #777;">
            تصميم وتطوير: <b>محمد سامي السعيد</b>
        </div>
        """
        
        render_printable_html_view(class_report_html, title=f"تقرير_{gr_select}_فصل_{cl_select}")

    st.markdown("---")
    st.markdown("### 📞 إدارة ورصد أرقام جوالات أولياء الأمور (تحديث وحفظ):")

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st_select = st.selectbox("اختر الطالب لتحديث رقم جوال ولي أمره:", df_reports["name"].tolist())

    selected_st_row = df_reports[df_reports["name"] == st_select].iloc[0]

    with col_p2:
        new_phone = st.text_input("رقم الجوال الجديد:", value=selected_st_row["phone"])
        if st.button("💾 تحديث رقم الجوال"):
            conn = get_db_connection()
            c = conn.cursor()
            c.execute("UPDATE students SET phone = ? WHERE id = ?", (new_phone, selected_st_row["id"]))
            conn.commit()
            conn.close()
            st.success(f"✅ تم تحديث رقم جوال الطالب {st_select} بنجاح!")
            st.rerun()

    st.markdown("---")
    st.markdown("### 🖨️ الاعتماد الرسمي وتوقيعات إدارة المدرسة:")

    st.markdown('''
    <div style="background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.05); text-align: center;">
        <table style="width: 100%; text-align: center; border-collapse: collapse;">
            <tr style="font-weight: bold; background-color: #1e3c72; color: white;">
                <td style="padding: 10px;">وكيل الشؤون التعليمية</td>
                <td style="padding: 10px;">وكيل شؤون الطلاب</td>
                <td style="padding: 10px;">مدير المدرسة</td>
            </tr>
            <tr style="font-size: 16px;">
                <td style="padding: 15px;"><b>محمد مبروك السيد</b></td>
                <td style="padding: 15px;"><b>صالح بن عبدالله الدعجاني</b></td>
                <td style="padding: 15px;"><b>إبراهيم بن موسى التميمي</b></td>
            </tr>
        </table>
        <hr>
        <p style="color: #666; font-size: 14px; margin-top: 10px;"><b>تصميم وتطوير:</b> محمد سامي السعيد</p>
    </div>
    ''', unsafe_allow_html=True)
