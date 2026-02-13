import os
from datetime import datetime

def create_project_dump():
    output_file = "project_dump.txt"
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("# ============================================\n")
        f.write("# ДАМП ПРОЕКТА: Бот для заказа обедов\n")
        f.write(f"# Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n")
        f.write("# ============================================\n\n")
        
        # Собираем все .py файлы
        py_files = []
        for root, dirs, files in os.walk('.'):
            if 'venv' in root or '__pycache__' in root or '.git' in root:
                continue
            for file in files:
                if file.endswith('.py'):
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path)
                    py_files.append((rel_path, full_path))
        
        # Сортируем файлы
        py_files.sort()
        
        # Записываем каждый файл
        for rel_path, full_path in py_files:
            f.write(f"# ============================================\n")
            f.write(f"# ФАЙЛ: {rel_path}\n")
            f.write(f"# ============================================\n\n")
            
            try:
                with open(full_path, 'r', encoding='utf-8') as py_file:
                    content = py_file.read()
                    f.write(content)
            except Exception as e:
                f.write(f"# ОШИБКА ЧТЕНИЯ: {e}\n")
            
            f.write("\n\n")
        
        # Добавляем requirements.txt
        if os.path.exists('requirements.txt'):
            f.write(f"# ============================================\n")
            f.write(f"# ФАЙЛ: requirements.txt\n")
            f.write(f"# ============================================\n\n")
            with open('requirements.txt', 'r', encoding='utf-8') as req:
                f.write(req.read())
            f.write("\n\n")
        
        # Добавляем структуру папок
        f.write(f"# ============================================\n")
        f.write(f"# СТРУКТУРА ПРОЕКТА\n")
        f.write(f"# ============================================\n\n")
        
        for root, dirs, files in os.walk('.'):
            if 'venv' in root or '__pycache__' in root or '.git' in root:
                continue
            level = root.replace('.', '').count(os.sep)
            indent = '    ' * level
            f.write(f"{indent}📁 {os.path.basename(root)}/\n")
            subindent = '    ' * (level + 1)
            for file in sorted(files):
                if file.endswith('.py') or file == 'requirements.txt':
                    f.write(f"{subindent}📄 {file}\n")
    
    size = os.path.getsize(output_file) / 1024
    print(f"✅ Дамп создан: {output_file}")
    print(f"📁 Размер: {size:.1f} KB")
    print(f"📊 Строк: {sum(1 for _ in open(output_file, 'r', encoding='utf-8'))}")

if __name__ == "__main__":
    create_project_dump()
    input("\nНажми Enter для выхода...")