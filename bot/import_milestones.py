import sqlite3
import csv
import urllib.request
import io
import hashlib

DB_PATH = '/root/tg_manager/data.db'
CSV_URL = 'https://docs.google.com/spreadsheets/d/15aFnbsYDqFSFx8zf9Er60RIb6s7kTAXqgdoX_uqs-ac/export?format=csv'

def init_milestones_table():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    c.execute('''CREATE TABLE IF NOT EXISTS objects 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  code TEXT UNIQUE,
                  name TEXT,
                  description TEXT)''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS milestones
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  short_code TEXT,
                  name TEXT,
                  object_id INTEGER,
                  object_code TEXT,
                  deadline_date TEXT,
                  status TEXT,
                  responsible TEXT,
                  responsible_email TEXT,
                  role TEXT,
                  depends_on_uuid TEXT,
                  blocks_uuid TEXT,
                  uuid TEXT UNIQUE,
                  last_updated TEXT,
                  FOREIGN KEY(object_id) REFERENCES objects(id))''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS milestone_dependencies
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  milestone_id INTEGER,
                  depends_on_id INTEGER,
                  FOREIGN KEY(milestone_id) REFERENCES milestones(id),
                  FOREIGN KEY(depends_on_id) REFERENCES milestones(id))''')
    
    conn.commit()
    conn.close()

def parse_date(date_str):
    if not date_str or date_str.strip() == '':
        return None
    try:
        date_str = date_str.strip().strip('"')
        parts = date_str.split('.')
        if len(parts) == 3:
            return f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}"
    except:
        pass
    return None

def generate_uuid(short_code, name, object_code):
    """Генерируем UUID на основе данных если его нет"""
    data = f"{short_code}|{name}|{object_code}"
    return hashlib.md5(data.encode()).hexdigest()

def import_milestones():
    print("Загружаем данные из Google Sheets...")
    
    response = urllib.request.urlopen(CSV_URL)
    content = response.read().decode('utf-8')
    
    reader = csv.reader(io.StringIO(content))
    rows = list(reader)
    
    if len(rows) < 2:
        print("Ошибка: нет данных в таблице")
        return
    
    headers = rows[0]
    print(f"Колонки: {headers[:10]}...")
    print(f"Всего строк: {len(rows) - 1}")
    
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    c.execute("DELETE FROM milestone_dependencies")
    c.execute("DELETE FROM milestones")
    c.execute("DELETE FROM objects")
    
    uuid_to_id = {}
    object_codes = set()
    
    # Первый проход - собираем объекты
    for row in rows[1:]:
        if len(row) < 15:
            continue
        
        corps = row[3].strip() if len(row) > 3 else ''
        
        if '10ф' in corps or 'ф10' in corps:
            obj_code = 'ПБ_ф10_к5'
            obj_name = 'ПБ корпус 5, фазенда 10'
        elif '16ф' in corps or 'ф16' in corps:
            obj_code = 'ПБ_ф16_к4'
            obj_name = 'ПБ корпус 4, фазенда 16'
        else:
            continue  # Пропускаем строки без объекта
        
        object_codes.add((obj_code, obj_name))
    
    for code, name in object_codes:
        c.execute("INSERT INTO objects (code, name) VALUES (?, ?)", (code, name))
    
    conn.commit()
    
    object_map = {}
    for code, name in object_codes:
        obj_id = c.execute("SELECT id FROM objects WHERE code = ?", (code,)).fetchone()[0]
        object_map[code] = obj_id
    
    imported = 0
    skipped = 0
    
    # Второй проход - импорт вех
    for row in rows[1:]:
        try:
            if len(row) < 15:
                skipped += 1
                continue
            
            # Пропускаем заголовки и техническую строку
            short_code = row[1].strip() if len(row) > 1 else ''
            first_col = row[0].strip() if len(row) > 0 else ''
            
            # Пропускаем заголовки
            if first_col == '№' or short_code == 'КОРОТКОЕ НАЗВАНИЕ':
                skipped += 1
                continue
            
            # Пропускаем техническую строку с цифрами (row 1)
            if first_col.isdigit() or short_code.isdigit():
                skipped += 1
                continue
            
            if not short_code or not short_code.strip():
                skipped += 1
                continue
            
            name = row[2].strip() if len(row) > 2 else ''
            corps = row[3].strip() if len(row) > 3 else ''
            deadline = row[4].strip() if len(row) > 4 else ''
            status = row[7].strip() if len(row) > 7 else ''
            responsible = row[8].strip() if len(row) > 8 else ''
            email = row[9].strip() if len(row) > 9 else ''
            role = row[10].strip() if len(row) > 10 else ''
            last_updated = row[11].strip() if len(row) > 11 else ''
            blocks_uuid = row[12].strip() if len(row) > 12 else ''  # У кого зависят сроки
            depends_on_uuid = row[13].strip() if len(row) > 13 else ''  # От какой вехи зависят
            uuid = row[14].strip() if len(row) > 14 else ''
            
            # Определяем объект
            if '10ф' in corps or 'ф10' in corps:
                obj_code = 'ПБ_ф10_к5'
            elif '16ф' in corps or 'ф16' in corps:
                obj_code = 'ПБ_ф16_к4'
            else:
                skipped += 1
                continue
            
            obj_id = object_map.get(obj_code, 1)
            
            # Если нет UUID, генерируем
            if not uuid:
                uuid = generate_uuid(short_code, name, obj_code)
            
            deadline_db = parse_date(deadline)
            
            status_map = {
                'Выполнено': 'done',
                'На исполнении': 'progress',
                'Новая': 'new',
                'В работе': 'progress',
                'Отложено': 'delayed'
            }
            status_db = status_map.get(status, 'new')
            
            if short_code and name:
                # Проверяем есть ли уже такая веха
                existing = c.execute("SELECT id FROM milestones WHERE uuid = ?", (uuid,)).fetchone()
                if existing:
                    skipped += 1
                    continue
                
                c.execute('''INSERT INTO milestones 
                            (short_code, name, object_id, object_code, deadline_date, status, 
                             responsible, responsible_email, role, depends_on_uuid, blocks_uuid, uuid, last_updated)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                         (short_code, name, obj_id, obj_code, deadline_db, status_db,
                          responsible, email, role, depends_on_uuid, blocks_uuid, uuid, last_updated))
                
                uuid_to_id[uuid] = c.lastrowid
                imported += 1
                
        except Exception as e:
            print(f"Ошибка в строке: {e}")
            skipped += 1
            continue
    
    conn.commit()
    
    # Создаём связи между вехами
    links_created = 0
    for uuid, mile_id in uuid_to_id.items():
        mile = c.execute("SELECT depends_on_uuid FROM milestones WHERE id = ?", (mile_id,)).fetchone()
        if mile and mile[0]:
            depends_on_uuid = mile[0]
            if depends_on_uuid in uuid_to_id:
                c.execute("INSERT INTO milestone_dependencies (milestone_id, depends_on_id) VALUES (?, ?)",
                         (mile_id, uuid_to_id[depends_on_uuid]))
                links_created += 1
    
    conn.commit()
    conn.close()
    
    print(f"\n✅ Импортировано вех: {imported}")
    print(f"✅ Пропущено строк: {skipped}")
    print(f"✅ Создано связей зависимостей: {links_created}")
    return imported

if __name__ == '__main__':
    init_milestones_table()
    import_milestones()