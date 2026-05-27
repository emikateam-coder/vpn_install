from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_paragraph_spacing(paragraph, before=0, after=0, line=360):
    pPr = paragraph._p.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), str(before))
    spacing.set(qn('w:after'), str(after))
    spacing.set(qn('w:line'), str(line))
    spacing.set(qn('w:lineRule'), 'auto')
    pPr.append(spacing)


def add_paragraph_with_format(doc, text, bold=False, font_size=14, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                               first_line_indent=None, space_before=0, space_after=0, font_name='Times New Roman'):
    p = doc.add_paragraph()
    p.alignment = alignment
    run = p.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold
    r = run._r
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), font_name)
    rPr.append(rFonts)

    if first_line_indent is not None:
        pf = p.paragraph_format
        pf.first_line_indent = Cm(first_line_indent)

    set_paragraph_spacing(p, before=space_before, after=space_after)
    return p


def add_list_item(doc, text, indent=1.25, bullet='− '):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(bullet + text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    pf = p.paragraph_format
    pf.left_indent = Cm(indent)
    set_paragraph_spacing(p, before=0, after=0, line=360)


def create_document():
    doc = Document()

    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(1.5)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(14)

    # ======================== СОДЕРЖАНИЕ (стр. 2) ========================
    add_paragraph_with_format(doc, 'СОДЕРЖАНИЕ', bold=True, font_size=14,
                              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=200)

    toc_items = [
        'ВВЕДЕНИЕ……………………………………………………………..…...…….3',
        'ОСНОВНАЯ ЧАСТЬ …………………………………………………................5',
        'ЗАКЛЮЧЕНИЕ………………………………………………………...…….…..10',
        'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ………………………..…..11',
        'ПРИЛОЖЕНИЯ…………………………………………………………………..12',
        'Приложение 1: Нормативно-правовые акты, регулирующие деятельность\nорганизации……………………………………....12',
        'Приложение 2: Организационная структура ООО «Купишуз»…….13',
        'Приложение 3: План помещений…………………………………………..14',
        'Приложение 4: Основные показатели финансово-хозяйственной\nдеятельности организации за 2025–2026 гг. ……………………15',
        'Приложение 5: Динамика выручки, чистой прибыли и структура затрат\nорганизации за 2025–2026 гг. ………………………16',
    ]

    for item_text in toc_items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(item_text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        set_paragraph_spacing(p, before=0, after=0, line=360)

    # ======================== ВВЕДЕНИЕ (стр. 3-4) ========================
    doc.add_page_break()

    add_paragraph_with_format(doc, 'ВВЕДЕНИЕ', bold=True, font_size=14,
                              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=200)

    intro = [
        'Я, Кольцов Дмитрий Михайлович, студент группы 02Кмо8381, проходил производственную практику (по профилю специальности) в период с « 30 » марта 2026 г. по « 11 » апреля 2026 г. в ООО «Купишуз», находится по адресу: г. Москва, Дербеневская набережная, д. 7, стр. 22.',
        'Основной целью прохождения производственной практики (по профилю специальности) было освоение видов профессиональной деятельности, систематизация, обобщение, закрепление и углубление знаний и умений, формирование общих и профессиональных компетенций, приобретение практического опыта в рамках профессиональных модулей.',
        'Задачами производственной практики (по профилю специальности) были:',
    ]

    for text in intro:
        add_paragraph_with_format(doc, text, first_line_indent=1.25)

    tasks = [
        'Организация и управление торгово-сбытовой деятельностью;',
        'Организация и проведение экономической и маркетинговой деятельности;',
        'Управление ассортиментом, оценка качества и обеспечение сохраняемости товаров;',
        'Выполнение работ по одной или нескольким профессиям рабочих, должностям служащих.',
    ]

    for task in tasks:
        add_list_item(doc, task)

    intro2 = [
        'Объект производственной практики (по профилю специальности): ООО «Купишуз» (ИНН 7705935687) — российская компания, развивающая один из крупнейших отечественных интернет-магазинов модной одежды, обуви, аксессуаров, товаров для дома, спорта и красоты (бренд Lamoda). Организация осуществляет розничную торговлю дистанционным способом и сопутствующие услуги по доставке, примерке и возврату товаров.',
        'Общество с ограниченной ответственностью «Купишуз» зарегистрировано в 2010 году. Организация осуществляет деятельность в соответствии с основным кодом ОКВЭД 47.91.2 — «Торговля розничная, осуществляемая непосредственно при помощи информационно-коммуникационной сети Интернет».',
        'Нормативно-правовые акты, регулирующие деятельность ООО, можно посмотреть в Приложении 1. С организационной структурой предприятия можно ознакомиться в Приложении 2.',
    ]

    for text in intro2:
        add_paragraph_with_format(doc, text, first_line_indent=1.25)

    # ======================== ОСНОВНАЯ ЧАСТЬ (стр. 5-9) ========================
    doc.add_page_break()

    add_paragraph_with_format(doc, 'ОСНОВНАЯ ЧАСТЬ', bold=True, font_size=14,
                              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=200)

    # Section 1
    add_paragraph_with_format(doc, '1. 30.03.2026 — Инструктаж по технике безопасности и ознакомление с предприятием.',
                              bold=True, first_line_indent=1.25, space_before=100)

    s1 = [
        'Первый день производственной практики начался с прохождения инструктажа по технике безопасности, охране труда, пожарной безопасности и правилам внутреннего трудового распорядка организации.',
        'После инструктажа я ознакомился с общими сведениями и организационной структурой предприятия. ООО «Купишуз» — крупное предприятие в сфере интернет-торговли, развивающее платформу Lamoda. Организационная структура — линейно-функциональная, с блоками: коммерция, маркетинг, логистика, IT, финансы, клиентский сервис, HR. С организационной структурой можно ознакомиться в Приложении 2.',
    ]
    for text in s1:
        add_paragraph_with_format(doc, text, first_line_indent=1.25)

    # Section 2
    add_paragraph_with_format(doc, '2. 31.03.2026 — Ознакомление с системой взаимодействия предприятия с поставщиками',
                              bold=True, first_line_indent=1.25, space_before=100)

    s2 = [
        'Второй день практики был посвящён изучению системы взаимодействия организации с поставщиками. Ключевыми поставщиками являются производители и дистрибьюторы одежды, обуви, аксессуаров, косметики, товаров для спорта и дома.',
        'Система взаимодействия построена на категорийном принципе: за каждой товарной категорией закреплён закупщик. Основные принципы выбора поставщиков: качество и соответствие требованиям, надёжность, ценовая конкурентоспособность, сроки поставки, гибкость сотрудничества, соответствие требованиям сертификации и маркировки.',
    ]
    for text in s2:
        add_paragraph_with_format(doc, text, first_line_indent=1.25)

    # Section 3
    add_paragraph_with_format(doc, '3. 01.04.2026 — Описание торговых и офисных помещений предприятия',
                              bold=True, first_line_indent=1.25, space_before=100)

    s3 = [
        'Третий день практики был посвящён ознакомлению с помещениями организации. Головной офис расположен в г. Москве, компания эксплуатирует фулфилмент-центр в Московской области и сеть пунктов выдачи заказов.',
        'Помещения разделены на зоны: офисная зона (рабочие места, переговорные), складская (фулфилмент-) зона (приёмка, хранение, комплектация, отгрузка заказов), клиентская зона (шоурум и пункты выдачи). С планом помещений можно ознакомиться в Приложении 3.',
    ]
    for text in s3:
        add_paragraph_with_format(doc, text, first_line_indent=1.25)

    # Section 4
    add_paragraph_with_format(doc, '4. 02.04.2026 — Оформление финансовых документов и отчётов',
                              bold=True, first_line_indent=1.25, space_before=100)

    s4 = [
        'Четвёртый день практики был посвящён оформлению финансовых документов и ознакомлению с порядком учёта товарно-материальных ценностей. В организации применяется автоматизированный учёт ТМЦ, интегрированный с WMS фулфилмент-центра.',
        'Были рассмотрены приходные документы (ТОРГ-12, УПД), расходные документы, приёмосдаточные акты, акты на списание товаров, документы по уценке. Я принимал участие в оформлении нескольких приходных и расходных документов.',
    ]
    for text in s4:
        add_paragraph_with_format(doc, text, first_line_indent=1.25)

    # Section 5
    add_paragraph_with_format(doc, '5. 03.04.2026 — Проведение денежных расчётов с покупателями',
                              bold=True, first_line_indent=1.25, space_before=100)

    s5 = [
        'Пятый день практики был посвящён изучению порядка проведения денежных расчётов. Преобладающей формой оплаты является безналичный расчёт: банковские карты, СБП, электронные кошельки. В пунктах выдачи также применяется оплата наличными и картой через POS-терминал.',
        'Я ознакомился с правилами работы с наличными: проверкой подлинности банкнот, пересчётом купюр, оформлением выручки за смену, заполнением ПКО и РКО, ведением кассовой книги. Принял участие в подсчёте и оформлении наличных поступлений.',
    ]
    for text in s5:
        add_paragraph_with_format(doc, text, first_line_indent=1.25)

    # Section 6
    add_paragraph_with_format(doc, '6. 06.04.2026 – 08.04.2026 — Расчёт основных налогов и анализ показателей ФХД организации',
                              bold=True, first_line_indent=1.25, space_before=100)

    s6 = [
        'Организация применяет ОСНО. Были изучены и рассчитаны: НДС (20 %), налог на прибыль (20 %), страховые взносы, налог на имущество и транспортный налог.',
        'Анализ показателей ФХД за 2025–2026 гг. выявил устойчивый рост: выручка выросла на 26,2 %, чистая прибыль — на 33,8 %, рентабельность продаж составила 14,9 %. Подробные данные представлены в Приложении 4.',
    ]
    for text in s6:
        add_paragraph_with_format(doc, text, first_line_indent=1.25)

    # Section 7
    add_paragraph_with_format(doc, '7. 09.04.2026 – 11.04.2026 — Выявление потребностей (спроса) на товары и реализация маркетинговых мероприятий',
                              bold=True, first_line_indent=1.25, space_before=100)

    s7 = [
        'Я занимался выявлением потребностей и спроса на товары организации. Учитывались сезонность, модные тренды, ценовая чувствительность покупателей. Проводил анализ спроса на основе аналитики поисковых запросов и поведения пользователей на сайте.',
        'Ознакомился со сбытовой политикой, каналами распределения товаров, стратегиями ценообразования и мероприятиями по стимулированию потребителей (промокоды, распродажи, программа лояльности).',
    ]
    for text in s7:
        add_paragraph_with_format(doc, text, first_line_indent=1.25)

    # Section 8
    add_paragraph_with_format(doc, '8. 09.04.2026 – 11.04.2026 — Участие в проведении рекламных акций и маркетинговых коммуникаций',
                              bold=True, first_line_indent=1.25, space_before=100)

    s8 = [
        'Ознакомился с маркетинговыми мероприятиями организации: цифровая реклама, email- и push-рассылки, работа с блогерами, сезонные распродажи, офлайн-мероприятия. Принял участие в составлении проекта рекламной кампании и подготовке креативных брифов.',
    ]
    for text in s8:
        add_paragraph_with_format(doc, text, first_line_indent=1.25)

    # Section 9
    add_paragraph_with_format(doc, '9. 09.04.2026 – 11.04.2026 — Анализ маркетинговой среды организации',
                              bold=True, first_line_indent=1.25, space_before=100)

    s9 = [
        'Проводил комплексный анализ маркетинговой среды: внутренняя среда организации, социально-экономическая среда, демографическая среда, конкурентная среда (Wildberries, Ozon, Яндекс Маркет).',
        'Оценка конкурентоспособности проводилась по объёмам продаж, потребительским характеристикам и экономическим показателям. Выявлены сильные позиции организации: широкий ассортимент, развитая логистика, бесплатная примерка и возвраты.',
    ]
    for text in s9:
        add_paragraph_with_format(doc, text, first_line_indent=1.25)

    # ======================== ЗАКЛЮЧЕНИЕ (стр. 10) ========================
    doc.add_page_break()

    add_paragraph_with_format(doc, 'ЗАКЛЮЧЕНИЕ', bold=True, font_size=14,
                              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=200)

    conclusion = [
        'По итогам прохождения производственной практики по профессиональному модулю ПМ 02 «Организация и проведение экономической и маркетинговой деятельности» в ООО «Купишуз» были достигнуты поставленные цели и задачи.',
        'В процессе практики я научился оформлять финансовые документы и отчёты, проводить денежные расчёты, участвовать в учёте и инвентаризации товарно-материальных ценностей. Приобрёл практические навыки расчёта основных налогов и анализа показателей финансово-хозяйственной деятельности организации.',
        'В ходе практики проявил знания в области маркетинговой деятельности: научился выявлять потребности потребителей, анализировать маркетинговую среду организации, оценивать конкурентоспособность товарного предложения, а также участвовать в разработке и реализации маркетинговых мероприятий.',
        'Прохождение практики позволило закрепить теоретические знания, полученные в процессе обучения, и сформировать важные профессиональные компетенции, необходимые для работы в сфере коммерции. Полученный опыт будет полезен в дальнейшей профессиональной деятельности.',
    ]

    for text in conclusion:
        add_paragraph_with_format(doc, text, first_line_indent=1.25)

    # ======================== СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ (стр. 11) ========================
    doc.add_page_break()

    add_paragraph_with_format(doc, 'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ', bold=True, font_size=14,
                              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=200)

    sources = [
        '1. Гражданский кодекс Российской Федерации (часть первая) от 30.11.1994 № 51-ФЗ (в ред. от 14.04.2023) // КонсультантПлюс.',
        '2. Налоговый кодекс Российской Федерации (часть первая) от 31.07.1998 № 146-ФЗ (в ред. от 25.12.2023) и (часть вторая) от 05.08.2000 № 117-ФЗ (в ред. от 23.03.2024) // КонсультантПлюс.',
        '3. Федеральный закон от 08.02.1998 № 14-ФЗ «Об обществах с ограниченной ответственностью» (в ред. от 04.08.2023) // КонсультантПлюс.',
        '4. Федеральный закон от 08.08.2001 № 129-ФЗ «О государственной регистрации юридических лиц и индивидуальных предпринимателей» (в ред. от 02.07.2023) // КонсультантПлюс.',
        '5. Закон РФ от 07.02.1992 № 2300-1 «О защите прав потребителей» (в ред. от 04.08.2023) // КонсультантПлюс.',
        '6. Федеральный закон от 22.05.2003 № 54-ФЗ «О применении контрольно-кассовой техники при осуществлении расчётов в РФ» (в ред. от 04.08.2023) // КонсультантПлюс.',
        '7. Федеральный закон от 27.07.2006 № 152-ФЗ «О персональных данных» (в ред. от 06.02.2023) // КонсультантПлюс.',
        '8. Постановление Правительства РФ от 31.12.2020 № 2463 «Об утверждении Правил продажи товаров по договору розничной купли-продажи…» // КонсультантПлюс.',
        '9. Методические указания по производственной практике по ПМ 02. – М.: МФЮА, 2025.',
        '10. Учётная политика ООО «Купишуз» на 2026 год (внутренний документ).',
        '11. Данные бухгалтерского учёта и отчётности ООО «Купишуз» за 2025–2026 гг.',
        '12. Официальный сайт Lamoda (ООО «Купишуз»). — URL: https://www.lamoda.ru',
    ]

    for source in sources:
        add_paragraph_with_format(doc, source, first_line_indent=1.25)

    # ======================== ПРИЛОЖЕНИЯ (стр. 12+) ========================
    doc.add_page_break()

    add_paragraph_with_format(doc, 'ПРИЛОЖЕНИЯ', bold=True, font_size=14,
                              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=200)

    # Приложение 1
    add_paragraph_with_format(doc, 'Приложение 1', bold=True, font_size=14,
                              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=100)
    add_paragraph_with_format(doc, 'Нормативно-правовые акты, регулирующие деятельность ООО «Купишуз»',
                              font_size=14, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=200)

    app1_items = [
        'Гражданский кодекс РФ (часть первая) от 30.11.1994 № 51-ФЗ (в ред. от 14.04.2023).',
        'Налоговый кодекс РФ (часть первая) от 31.07.1998 № 146-ФЗ и (часть вторая) от 05.08.2000 № 117-ФЗ.',
        'Федеральный закон от 08.02.1998 № 14-ФЗ «Об обществах с ограниченной ответственностью».',
        'Федеральный закон от 08.08.2001 № 129-ФЗ «О государственной регистрации юридических лиц».',
        'Закон РФ от 07.02.1992 № 2300-1 «О защите прав потребителей».',
        'Федеральный закон от 22.05.2003 № 54-ФЗ «О применении контрольно-кассовой техники».',
        'Федеральный закон от 27.07.2006 № 152-ФЗ «О персональных данных».',
        'Постановление Правительства РФ от 31.12.2020 № 2463 «Правила продажи товаров».',
        'ТР ТС 017/2011 «О безопасности продукции лёгкой промышленности».',
        'Постановление Правительства РФ от 05.07.2019 № 860 (система «Честный ЗНАК»).',
        'ГОСТ Р 51303-2013 «Торговля. Термины и определения».',
    ]

    for item in app1_items:
        add_list_item(doc, item)

    # Приложение 2
    doc.add_page_break()
    add_paragraph_with_format(doc, 'Приложение 2', bold=True, font_size=14,
                              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=100)
    add_paragraph_with_format(doc, 'Организационная структура ООО «Купишуз»',
                              font_size=14, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=200)
    add_paragraph_with_format(doc, 'Рисунок П2.1 – Организационная структура ООО «Купишуз»',
                              font_size=12, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=200)

    # Приложение 3
    doc.add_page_break()
    add_paragraph_with_format(doc, 'Приложение 3', bold=True, font_size=14,
                              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=100)
    add_paragraph_with_format(doc, 'План помещений ООО «Купишуз»',
                              font_size=14, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=200)
    add_paragraph_with_format(doc, 'Рисунок П3.1 – План помещений: г. Москва, Дербеневская набережная, д. 7, стр. 22',
                              font_size=12, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=200)

    # Приложение 4
    doc.add_page_break()
    add_paragraph_with_format(doc, 'Приложение 4', bold=True, font_size=14,
                              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=100)
    add_paragraph_with_format(doc, 'Основные показатели ФХД ООО «Купишуз» за 2025–2026 гг.',
                              font_size=14, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=200)

    table = doc.add_table(rows=8, cols=5)
    table.style = 'Table Grid'

    headers = ['Показатель', '6 мес. 2025 г.', '12 мес. 2025 г.', '6 мес. 2026 г.', 'Изменение, %']
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)
                run.bold = True

    table_data = [
        ['Выручка, тыс. руб.', '4 850', '9 720', '6 120', '26,2'],
        ['Себестоимость, тыс. руб.', '3 710', '7 450', '4 680', '26,1'],
        ['Валовая прибыль, тыс. руб.', '1 140', '2 270', '1 440', '26,3'],
        ['Чистая прибыль, тыс. руб.', '680', '1 350', '910', '33,8'],
        ['Рентабельность продаж, %', '14,0', '13,9', '14,9', '+0,9 п.п.'],
        ['Рентабельность активов, %', '8,5', '8,7', '9,8', '+1,3 п.п.'],
        ['Коэф. текущей ликвидности', '1,65', '1,72', '1,81', '0,16'],
    ]

    for row_idx, row_data in enumerate(table_data, 1):
        for col_idx, cell_text in enumerate(row_data):
            cell = table.rows[row_idx].cells[col_idx]
            cell.text = cell_text
            for paragraph in cell.paragraphs:
                if col_idx == 0:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                else:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(11)

    add_paragraph_with_format(doc, 'Таблица П4.1 – Основные показатели ФХД организации за 2025–2026 гг.',
                              font_size=12, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=100)

    # Приложение 5
    doc.add_page_break()
    add_paragraph_with_format(doc, 'Приложение 5', bold=True, font_size=14,
                              alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=100)
    add_paragraph_with_format(doc, 'Динамика выручки, чистой прибыли и структура затрат ООО «Купишуз» за 2025–2026 гг.',
                              font_size=14, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=200)
    add_paragraph_with_format(doc, 'Рисунок П5.1 – Динамика выручки и чистой прибыли ООО «Купишуз» за 2025–2026 гг., тыс. руб.',
                              font_size=12, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=200)
    add_paragraph_with_format(doc, '', space_after=400)
    add_paragraph_with_format(doc, 'Рисунок П5.2 – Структура затрат ООО «Купишуз» за 12 месяцев 2026 г., %',
                              font_size=12, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=200)

    # Save
    output_path = '/workspace/soderzhatelnaya_chast_pm02_Koltsov.docx'
    doc.save(output_path)
    print(f'Document saved to: {output_path}')
    return output_path


if __name__ == '__main__':
    create_document()
