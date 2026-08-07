from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def fix_settings_theme_lang(doc):
    """
    Удаляет <w:themeFontLang ...> из settings.xml, если есть.
    Работает через старый API part_related_by, без get_rel.
    """
    reltype_settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    try:
        settings_part = doc.part.package.part_related_by(reltype_settings)
    except KeyError:
        # у документа ещё нет settings.xml
        return

    settings = settings_part._element  # <w:settings>
    theme_lang = settings.find(qn('w:themeFontLang'))
    if theme_lang is not None:
        settings.remove(theme_lang)


def set_docdefaults_language_ru(doc):
    """
    Добавляет/обновляет <w:docDefaults>/<w:rPrDefault>/<w:rPr>/<w:lang w:val="ru"/>
    в styles.xml.
    """
    reltype_styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    try:
        styles_part = doc.part.package.part_related_by(reltype_styles)
    except KeyError:
        # нет styles.xml (маловероятно, но на всякий случай)
        return

    styles = styles_part._element  # <w:styles>

    docDefaults = styles.find(qn('w:docDefaults'))
    if docDefaults is None:
        docDefaults = OxmlElement('w:docDefaults')
        styles.insert(0, docDefaults)

    rPrDefault = docDefaults.find(qn('w:rPrDefault'))
    if rPrDefault is None:
        rPrDefault = OxmlElement('w:rPrDefault')
        docDefaults.append(rPrDefault)

    rPr = rPrDefault.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        rPrDefault.append(rPr)

    lang = rPr.find(qn('w:lang'))
    if lang is None:
        lang = OxmlElement('w:lang')
        rPr.append(lang)

    lang.set(qn('w:val'), 'ru')


def set_normal_style_language_ru(doc):
    """
    Дополнительно прописывает язык ru в стиле Normal:
    <w:style w:styleId="Normal"> ... <w:rPr><w:lang w:val="ru"/></w:rPr>
    """
    try:
        normal = doc.styles['Normal']
    except KeyError:
        return

    rPr = normal.element.get_or_add_rPr()
    lang = rPr.find(qn('w:lang'))
    if lang is None:
        lang = OxmlElement('w:lang')
        rPr.append(lang)

    lang.set(qn('w:val'), 'ru')
