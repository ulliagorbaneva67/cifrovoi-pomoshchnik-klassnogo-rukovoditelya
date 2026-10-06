import io
import json
import urllib.parse
import urllib.request
import zipfile


BASE = "http://127.0.0.1:8000"
CASES = [
    ("student", {"student_name": "Иванов Иван Иванович", "class_name": "7А", "organization": "МБОУ Школа №1", "academic": "успевает на хорошем уровне", "attitude": "проявляет ответственность", "qualities": "внимательный и доброжелательный"}),
    ("meeting", {"organization": "МБОУ Школа №1", "class_name": "7А", "meeting_date": "06.10.2026", "teacher": "Петрова Анна Сергеевна", "agenda": "Итоги первой четверти", "decisions": "продолжить взаимодействие семьи и школы"}),
    ("analysis", {"organization": "МБОУ Школа №1", "class_name": "7А", "school_year": "2025-2026", "teacher": "Петрова Анна Сергеевна", "directions": "гражданско-патриотическое воспитание", "events": "проведены классные часы и экскурсии", "conclusions": "работа выполнена системно", "next_tasks": "развивать ученическое самоуправление"}),
]


for doc_type, fields in CASES:
    request = urllib.request.Request(
        BASE + "/api/generate",
        data=json.dumps({"type": doc_type, "fields": fields}, ensure_ascii=False).encode(),
        headers={"Content-Type": "application/json"},
    )
    result = json.load(urllib.request.urlopen(request))
    assert result["paragraphs"]
    download = urllib.request.urlopen(BASE + "/files/" + urllib.parse.quote(result["filename"])).read()
    archive = zipfile.ZipFile(io.BytesIO(download))
    assert download[:2] == b"PK"
    assert "word/document.xml" in archive.namelist()
    print(doc_type, result["filename"], len(download), "bytes: OK")
