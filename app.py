import json
import re
import zipfile
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, unquote, urlparse
from xml.sax.saxutils import escape

ROOT = Path(__file__).parent
GENERATED = ROOT / "generated"
GENERATED.mkdir(exist_ok=True)


def value(data, key, fallback=""):
    return str(data.get(key, "")).strip() or fallback


def paragraphs_for(doc_type, d):
    if doc_type == "student":
        name = value(d, "student_name", "обучающийся")
        p = [
            f"Характеристика на {name}",
            f"{value(d, 'student_name', 'Обучающийся')} обучается в {value(d, 'organization', 'образовательной организации')}, {value(d, 'class_name', 'указанном классе')}."
        ]
        if d.get("birth_date"):
            p.append(f"Дата рождения: {d['birth_date']}.")
        if d.get("study_period"):
            p.append(f"Период обучения: {d['study_period']}.")
        for key, label in [("academic", "Успеваемость"), ("attitude", "Отношение к учебной деятельности"), ("motivation", "Учебная мотивация")]:
            if d.get(key):
                p.append(f"{label}: {d[key]}.")
        social = [("class_life", "В жизни класса"), ("school_life", "В жизни школы"), ("peers", "Взаимоотношения со сверстниками"), ("teachers", "Взаимоотношения с педагогами"), ("qualities", "Личностные качества"), ("extracurricular", "Внеурочная деятельность")]
        details = [f"{label.lower()}: {d[key]}" for key, label in social if d.get(key)]
        if details:
            p.append("В социально-общественной и внеурочной сфере отмечается следующее: " + "; ".join(details) + ".")
        if d.get("additional"):
            p.append(f"Дополнительная информация: {d['additional']}.")
        p.append("Характеристика составлена для предоставления по месту требования.")
        return p
    if doc_type == "meeting":
        return [
            "ПРОТОКОЛ № 1",
            f"родительского собрания {value(d, 'class_name', '_____')} класса",
            f"Дата проведения: {value(d, 'meeting_date', '__________')}. Место проведения: {value(d, 'place', '__________')}.",
            f"Председатель собрания: классный руководитель {value(d, 'teacher', '__________')}. Присутствовали: {value(d, 'parents_count', '__________')} родителей.",
            f"Повестка дня: {value(d, 'agenda', 'вопросы текущей учебно-воспитательной работы')}.",
            f"По первому вопросу: {value(d, 'questions', 'обсуждены актуальные вопросы обучения и воспитания обучающихся')}.",
            f"Выступили: {value(d, 'speeches', 'классный руководитель и родители обучающихся')}.",
            f"Обсуждение: {value(d, 'discussion', 'участники обменялись мнениями и предложениями')}.",
            f"Решили: {value(d, 'decisions', 'принять предложенные решения и продолжить взаимодействие семьи и школы')}.",
            *( [f"Дополнительные сведения: {d['additional']}"] if d.get("additional") else [] ),
            "Классный руководитель: ____________________ / " + value(d, "teacher", "____________________")
        ]
    p = [
        f"Анализ работы классного руководителя за {value(d, 'school_year', 'учебный год')}",
        f"Образовательная организация: {value(d, 'organization', 'не указана')}. Класс: {value(d, 'class_name', 'не указан')}. Классный руководитель: {value(d, 'teacher', 'не указан')}.",
        f"В классе обучается {value(d, 'students_count', 'не указано')} обучающихся.",
        f"Основные направления воспитательной работы: {value(d, 'directions', 'гражданско-патриотическое, духовно-нравственное, физическое и социальное развитие')}.",
        f"Проведённые мероприятия: {value(d, 'events', 'мероприятия проводились в соответствии с планом воспитательной работы')}.",
        f"Работа с родителями: {value(d, 'parents_work', 'осуществлялась посредством родительских собраний и индивидуальных консультаций')}.",
        f"Индивидуальная работа с обучающимися: {value(d, 'individual_work', 'проводилась с учётом индивидуальных образовательных и воспитательных потребностей')}.",
        f"Работа с активом класса: {value(d, 'class_asset', 'актив класса привлекался к планированию и проведению мероприятий')}.",
        f"Участие обучающихся в школьных мероприятиях: {value(d, 'school_events', 'обучающиеся принимали участие в общешкольных мероприятиях')}.",
        f"Результаты работы: {value(d, 'results', 'поставленные задачи в основном выполнены')}.",
        f"Выявленные проблемы: {value(d, 'problems', 'существенных проблем не выявлено')}.",
        f"Достигнутые результаты: {value(d, 'achievements', 'созданы условия для развития коллектива и повышения активности обучающихся')}.",
        f"Выводы: {value(d, 'conclusions', 'работа классного руководителя осуществлялась системно и последовательно')}.",
        f"Задачи на следующий учебный год: {value(d, 'next_tasks', 'продолжить работу по развитию классного коллектива и взаимодействию с родителями')}.",
    ]
    if d.get("additional"):
        p.append(f"Дополнительная информация: {d['additional']}.")
    p.append("Классный руководитель: ____________________ / " + value(d, "teacher", "____________________"))
    return p


def safe_filename(doc_type, d):
    raw = {"student": "Характеристика_" + value(d, "student_name", "обучающийся"), "meeting": "Протокол_родительского_собрания_" + value(d, "meeting_date", "дата"), "analysis": "Анализ_работы_классного_руководителя_" + value(d, "school_year", "год")}[doc_type]
    return re.sub(r"[^\wА-Яа-яЁё ._-]", "", raw).replace(" ", "_") + ".docx"


def docx_bytes(title, paragraphs):
    def para(text, bold=False, center=False):
        align = '<w:jc w:val="center"/>' if center else ''
        weight = '<w:b/>' if bold else ''
        return f'<w:p><w:pPr>{align}<w:spacing w:after="160"/></w:pPr><w:r><w:rPr>{weight}<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'
    if paragraphs and paragraphs[0] == title:
        body = para(paragraphs[0], True, True) + ''.join(para(x) for x in paragraphs[1:])
    else:
        body = para(title, True, True) + ''.join(para(x) for x in paragraphs)
    body += '<w:p><w:r><w:t></w:t></w:r></w:p>'
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>'''
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
    out = __import__('io').BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/document.xml', document)
    return out.getvalue()


class Handler(BaseHTTPRequestHandler):
    def send_json(self, payload, status=200):
        raw = json.dumps(payload, ensure_ascii=False).encode()
        self.send_response(status); self.send_header('Content-Type', 'application/json; charset=utf-8'); self.send_header('Content-Length', str(len(raw))); self.end_headers(); self.wfile.write(raw)

    def do_GET(self):
        path = urlparse(self.path).path
        if path.startswith('/files/'):
            filename = Path(unquote(path.removeprefix('/files/'))).name
            file = GENERATED / filename
            if file.exists() and file.suffix == '.docx':
                raw = file.read_bytes()
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                self.send_header('Content-Disposition', f"attachment; filename*=UTF-8''{quote(filename)}")
                self.send_header('Content-Length', str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)
                return
            self.send_error(404)
            return
        if path == '/' or path == '/index.html':
            raw = (ROOT / 'templates' / 'index.html').read_bytes(); self.send_response(200); self.send_header('Content-Type', 'text/html; charset=utf-8'); self.send_header('Content-Length', str(len(raw))); self.end_headers(); self.wfile.write(raw); return
        if path.startswith('/static/'):
            file = ROOT / path.lstrip('/')
            if file.exists():
                raw = file.read_bytes(); self.send_response(200); self.send_header('Content-Type', 'text/css' if file.suffix == '.css' else 'application/javascript'); self.send_header('Content-Length', str(len(raw))); self.end_headers(); self.wfile.write(raw); return
        self.send_error(404)

    def do_POST(self):
        if self.path != '/api/generate': self.send_error(404); return
        try:
            data = json.loads(self.rfile.read(int(self.headers.get('Content-Length', 0))))
            doc_type = data.get('type'); fields = data.get('fields', {})
            titles = {'student': 'Характеристика на обучающегося', 'meeting': 'Протокол родительского собрания', 'analysis': 'Анализ работы классного руководителя за учебный год'}
            if doc_type not in titles: raise ValueError('Неизвестный тип документа')
            paragraphs = paragraphs_for(doc_type, fields); filename = safe_filename(doc_type, fields); content = docx_bytes(titles[doc_type], paragraphs)
            target = GENERATED / filename; target.write_bytes(content)
            self.send_json({'title': titles[doc_type], 'paragraphs': paragraphs, 'filename': filename, 'download': '/files/' + filename})
        except Exception as exc:
            self.send_json({'error': str(exc)}, 400)

    def do_HEAD(self):
        self.send_error(404)

    def log_message(self, fmt, *args):
        print(f"{self.address_string()} - {fmt % args}")


if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 8000))
    print(f'Цифровой помощник запущен на порту {port}')
    ThreadingHTTPServer(('0.0.0.0', port), Handler).serve_forever()
