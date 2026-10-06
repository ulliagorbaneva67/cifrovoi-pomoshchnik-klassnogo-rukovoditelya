import io
import json
import urllib.parse
import urllib.request
import zipfile
from xml.etree import ElementTree


BASE = "http://127.0.0.1:8000"
CASES = [
    ("student", {"student_name": "Проверка 2026 А.А.", "class_name": "9Б", "organization": "Тестовая школа", "academic": "хорошая успеваемость", "qualities": "ответственный"}),
    ("meeting", {"organization": "Тестовая школа", "class_name": "9Б", "meeting_date": "07.10.2026", "teacher": "Проверка Анна Сергеевна", "agenda": "Итоги четверти", "decisions": "продолжить работу"}),
    ("analysis", {"organization": "Тестовая школа", "class_name": "9Б", "school_year": "2026-check", "teacher": "Проверка Анна Сергеевна", "directions": "гражданское воспитание", "events": "классные часы", "conclusions": "работа выполнена", "next_tasks": "продолжить работу"}),
]

page = urllib.request.urlopen(BASE).read()
assert b"Digital" not in page
assert "Цифровой помощник".encode() in page
print("interface: HTTP 200")

for doc_type, fields in CASES:
    request = urllib.request.Request(BASE + "/api/generate", data=json.dumps({"type": doc_type, "fields": fields}, ensure_ascii=False).encode(), headers={"Content-Type": "application/json"})
    result = json.load(urllib.request.urlopen(request))
    assert result["paragraphs"] and result["filename"].endswith(".docx")
    raw = urllib.request.urlopen(BASE + "/files/" + urllib.parse.quote(result["filename"])).read()
    with zipfile.ZipFile(io.BytesIO(raw)) as archive:
        assert "word/document.xml" in archive.namelist()
        ElementTree.fromstring(archive.read("word/document.xml"))
        assert "[Content_Types].xml" in archive.namelist()
    print(f"{doc_type}: generated, downloaded, XML valid, {len(raw)} bytes")
