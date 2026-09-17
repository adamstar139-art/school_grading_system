import streamlit as st
import pandas as pd
import sqlite3
import io

# محاولة استيراد المكتبات الاختيارية لحمايتها عند الاستضافة على المنصات
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

# ==============================================================================
# 1. تهيئة قاعدة البيانات والبيانات المعتمدة للطلاب
# ==============================================================================
DB_NAME = "school_grading_system.db"

# قائمة طلاب متوسطة الثغر النموذجية الأهلية المسجلين بالصفوف والفصول
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
    {"id": "1158966166", "name": "أاصيل ناصر بن محمد مذكور", "grade": "الثالث المتوسط", "class": 1, "phone": "966552149044"},
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
def generate_parent_message(student_name, score, is_absent):
    if is_absent:
        return f"المكرم ولي أمر الطالب/ {student_name}، نود التنبيه على غياب الطالب هذا الأسبوع، ونحثكم على متابعة انتظام وحضور الاختبارات لتجنب حسم الدرجات والتأثير على مستواه الأكاديمي."
    elif score is None or score < 50:
        sc_str = f"{score}%" if score is not None else "0%"
        return f"المكرم ولي أمر الطالب/ {student_name}، نفيدكم بأن نسبة إتقان الطالب هذا الأسبوع هي ({sc_str}) وهي أقل من 50%. نأمل منكم المتابعة والعمل على رفع مستواه الدراسي."
    elif score <= 75:
        return f"المكرم ولي أمر الطالب/ {student_name}، نشيد بجهود الطالب ونسبة إتقانه ({score}%)، ونحثه على بذل المزيد من الجهد للوصول إلى مستوى أفضل والتميز."
    else:
        return f"المكرم ولي أمر الطالب/ {student_name}، نتقدم بخالص الشكر والتقدير لكم وللطالب على الاهتمام والتفوق بنسبة إتقان ممتازة ({score}%)، مع أطيب دعواتنا له بالتوفيق والاستمرارية."

# ==============================================================================
# 3. واجهة برنامج Streamlit وتحديد التنسيق العربي RTL
# ==============================================================================
st.set_page_config(
    page_title="برنامج رصد الدرجات - متوسطة الثغر النموذجية الأهلية",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown('''
<style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Tajawal', sans-serif;
        direction: rtl;
        text-align: right;
    }
    .stApp {
        background-color: #f4f6f9;
    }
    .header-box {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 20px;
        border-radius: 12px;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
</style>
''', unsafe_allow_html=True)

if not HAS_REPORTLAB or not HAS_OPENPYXL:
    st.info("💡 تنبيه الاستضافة: لتفعيل كافة ميزات تصدير ملفات Excel و PDF على المنصة، يرجى التأكد من إضافة ملف `requirements.txt` في مجلد مشروعك الأساسي.")

# الهيدر الرئيسي للمدرسة
st.markdown('''
<div class="header-box">
    <h2>🏫 المملكة العربية السعودية - وزارة التعليم</h2>
    <h3>الإدارة العامة للتعليم بمنطقة الرياض | متوسطة الثغر النموذجية الأهلية (بنين)</h3>
    <h4>برنامج رصد درجات الإتقان وإدارة تقارير أولياء الأمور</h4>
</div>
''', unsafe_allow_html=True)

# شريط التنقل الجانبي
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
            st.markdown("##### 📥 جدول رصد الدرجات والنسب المئوية للطلاب:")
            st.caption("ملاحظة: تظهر الدرجة الافتراضية (0) أمام كل طالب، وعند اختيار 'غائب' تصفر البيانات وتلغى الدرجة.")
            
            # عناوين الأعمدة في الجدول
            hdr1, hdr2, hdr3, hdr4 = st.columns([3.5, 2, 2, 1.5])
            hdr1.markdown("**اسم الطالب**")
            hdr2.markdown("**درجة الإتقان (100)**")
            hdr3.markdown("**النسبة المئوية (%)**")
            hdr4.markdown("**خيار غائب**")
            st.markdown("<hr style='margin: 5px 0 15px 0;'>", unsafe_allow_html=True)
            
            updated_data = []
            for idx, row in df_students.iterrows():
                # logic for default score: 0 unless previously recorded in DB
                is_db_abs = bool(row["is_absent"]) if pd.notnull(row["is_absent"]) else False
                if is_db_abs:
                    curr_score = 0.0
                    curr_abs = True
                else:
                    curr_score = float(row["score"]) if (pd.notnull(row["score"]) and row["score"] is not None) else 0.0
                    curr_abs = False

                col_name, col_score, col_pct, col_absent = st.columns([3.5, 2, 2, 1.5])
                
                with col_name:
                    st.write(f"**{idx+1}. {row['name']}**")
                with col_score:
                    sc = st.number_input(
                        label=f"sc_lbl_{row['id']}",
                        min_value=0.0,
                        max_value=100.0,
                        value=curr_score,
                        step=1.0,
                        key=f"sc_{row['id']}",
                        label_visibility="collapsed"
                    )
                with col_absent:
                    is_abs = st.checkbox("غائب ⚪", value=curr_abs, key=f"abs_{row['id']}")
                
                # Zero out data if absent is selected
                if is_abs:
                    final_score = 0.0
                    pct_str = "0%"
                else:
                    final_score = sc
                    pct_str = f"{final_score:.0f}%" if final_score.is_integer() else f"{final_score}%"

                with col_pct:
                    if is_abs:
                        st.markdown("<span style='color: #777; font-weight: bold;'>0% (غائب)</span>", unsafe_allow_html=True)
                    else:
                        st.markdown(f"**{pct_str}**")

                updated_data.append({
                    "student_id": row["id"],
                    "name": row["name"],
                    "score": final_score,
                    "is_absent": 1 if is_abs else 0,
                    "pct_display": pct_str
                })
            
            st.markdown("<br>", unsafe_allow_html=True)
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
            st.success("✅ تم حفظ البيانات في قاعدة البيانات بنجاح!")
            st.rerun()

        st.markdown("### 📊 جدول نتائج الرصد المنسق بالتلوين الشرطي:")
        
        table_rows = []
        for item in updated_data:
            sc = item["score"]
            is_abs = item["is_absent"]
            
            if is_abs:
                cat = "غائب ⚪ (0%)"
                pct = "0%"
                sc_disp = 0
            else:
                pct = item["pct_display"]
                sc_disp = sc
                if sc < 50:
                    cat = "أقل من 50% (ضعيف) 🔴"
                elif sc <= 75:
                    cat = "50% - 75% (متوسط) 🔵"
                else:
                    cat = "76% - 100% (ممتاز) 🟢"
                
            table_rows.append({
                "اسم الطالب": item["name"],
                "رقم الهوية": item["student_id"],
                "درجة الإتقان / 100": sc_disp,
                "النسبة المئوية": pct,
                "الفئة / الحالة": cat
            })
        
        df_display = pd.DataFrame(table_rows)
        st.dataframe(df_display, use_container_width=True)

        st.markdown("---")
        st.markdown("### 📥 تصدير وطباعة كشوف الدرجات:")
        
        col_exp1, col_exp2 = st.columns(2)
        
        # تصدير ملف Excel
        with col_exp1:
            if HAS_OPENPYXL:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "كشف_الرصد"
                ws.sheet_view.rightToLeft = True
                
                headers = ["م", "رقم الهوية", "اسم الطالب", "درجة الإتقان / 100", "النسبة المئوية", "الحالة"]
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
            else:
                st.warning("تصدير Excel يتطلب تثبيت مكتبة `openpyxl` في `requirements.txt`")

        # تصدير كشف مفرغ PDF
        with col_exp2:
            if HAS_REPORTLAB:
                pdf_io = io.BytesIO()
                doc = SimpleDocTemplate(pdf_io, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                elements = []
                styles = getSampleStyleSheet()
                
                title_style = ParagraphStyle("Title", parent=styles["Heading1"], alignment=1, fontSize=14)
                elements.append(Paragraph(f"<b>متوسطة الثغر النموذجية الأهلية - كشف رصد مفرغ</b>", title_style))
                elements.append(Paragraph(f"<b>الصف: {grade} | الفصل: ({class_num}) | {term} | {week}</b>", title_style))
                elements.append(Spacer(1, 15))
                
                table_data = [["م", "رقم الهوية", "اسم الطالب", "درجة الإتقان / 100", "ملاحظات"]]
                for i, item in enumerate(updated_data, 1):
                    table_data.append([str(i), item["student_id"], item["name"], "0", ""])
                    
                t = Table(table_data, colWidths=[30, 90, 200, 100, 100])
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1e3c72')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('GRID', (0,0), (-1,-1), 1, colors.black),
                    ('FONTSIZE', (0,0), (-1,-1), 10),
                ]))
                elements.append(t)
                doc.build(elements)
                pdf_io.seek(0)
                
                st.download_button(
                    label="📝 طباعة كشف مفرغ للرصد اليدوي (PDF)",
                    data=pdf_io,
                    file_name=f"كشف_مفرغ_{grade}_{class_num}.pdf",
                    mime="application/pdf"
                )
            else:
                st.warning("تصدير PDF يتطلب تثبيت مكتبة `reportlab` في `requirements.txt`")

# ==============================================================================
# الصفحة الثانية: إدارة المدرسة وتقارير أولياء الأمور (ADMIN & PARENT REPORTS)
# ==============================================================================
elif page == "🏫 إدارة المدرسة وتقارير أولياء الأمور":
    st.subheader("🏫 إدارة المدرسة وإرسال تقارير أولياء الأمور")
    
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
            "score": sc,
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
    
    if st.button("📲 إرسال الرسائل النصية الجماعية لجميع الأولياء الأمور (مرة واحدة)"):
        st.success(f"🚀 تم إرسال الرسائل النصية بنجاح إلى {len(df_reports)} ولي أمر مقسمين على الفئات الأربع!")

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
            for item in cat_list:
                with st.expander(f"👤 {item['name']} ({item['grade']} - فصل {item['class']}) | جوال ولي الأمر: {item['phone']}"):
                    st.write(f"**رقم الهوية:** {item['id']}")
                    st.write(f"**النسبة المئوية / الدرجة:** {item['score'] if not item['is_absent'] else 'غائب (0%)'}")
                    st.info(f"💬 **نص الرسالة الموجهة:**\n{item['message']}")

    with tab1: show_category_tab(cat_red, "فئة أقل من 50%")
    with tab2: show_category_tab(cat_blue, "فئة 50% - 75%")
    with tab3: show_category_tab(cat_green, "فئة 76% - 100%")
    with tab4: show_category_tab(cat_gray, "فئة الغياب")

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
