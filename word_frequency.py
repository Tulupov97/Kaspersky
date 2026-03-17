import aiofiles
import pymorphy3
import openpyxl
from openpyxl.styles import Font
from typing import Dict, List
import asyncio

morph = pymorphy3.MorphAnalyzer(lang='ru')

# Ограничение на количество одновременных обработок файлов
MAX_CONCURRENT_TASKS = 3
semaphore = asyncio.Semaphore(MAX_CONCURRENT_TASKS)

class WordFrequencyProcessor:
    def __init__(self):
        self.normal_forms: Dict[str, List[int]] = {}  # словоформа -> список количеств в строках

    async def process_file(self, file_path: str):
        """Обработка файла с подсчетом частоты слов по строкам."""
        self.normal_forms.clear()
        line_index = 0

        async with aiofiles.open(file_path, mode='r', encoding='utf-8') as f:
            async for line in f:
                line = line.strip().lstrip('\ufeff')
                if not line:
                    line_index += 1
                    continue

                words = line.split()
                line_counts = {}

                for word in words:
                    clean_word = word.strip('.,!?"\'()[]{};:-')
                    if not clean_word:
                        continue
                    parsed = morph.parse(clean_word)[0]
                    normal_form = parsed.normal_form

                    line_counts[normal_form] = line_counts.get(normal_form, 0) + 1

                # Обновляем статистику по всем строкам
                for normal_form, count in line_counts.items():
                    if normal_form not in self.normal_forms:
                        self.normal_forms[normal_form] = [0] * line_index  # заполняем нулями до текущей строки
                    elif len(self.normal_forms[normal_form]) < line_index:
                        # расширяем список до текущей строки
                        self.normal_forms[normal_form].extend([0] * (line_index - len(self.normal_forms[normal_form])))

                    self.normal_forms[normal_form].append(count)

                line_index += 1

        # Заполняем оставшиеся позиции нулями для всех слов
        for counts in self.normal_forms.values():
            if len(counts) < line_index:
                counts.extend([0] * (line_index - len(counts)))

    def generate_report(self, output_path: str):
        """Создание XLSX отчета."""
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Frequency Report"

        # Заголовки
        sheet["A1"] = "Словоформа"
        sheet["B1"] = "Общее количество"
        sheet["C1"] = "Количество по строкам"

        # Жирный шрифт для заголовков
        for cell in ["A1", "B1", "C1"]:
            sheet[cell].font = Font(bold=True)

        row = 2
        for normal_form, counts in self.normal_forms.items():
            total_count = sum(counts)
            counts_str = ",".join(map(str, counts))
            sheet[f"A{row}"] = normal_form
            sheet[f"B{row}"] = total_count
            sheet[f"C{row}"] = counts_str
            row += 1

        workbook.save(output_path)