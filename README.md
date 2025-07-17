# nap_stay_declaration

# E_STAY_GEN

**Генератор на XML файлове за НАП (Електронни транспортни декларации за престой - ЕДДП)**

`E_STAY_GEN` е настолно Python приложение с графичен интерфейс, което автоматично генерира XML файлове в съответствие с изискванията на НАП за престой на транспортно средство с гориво. Приложението извлича информация от входен XML файл (ЕДП), допълва я с адресни данни от потребителя и създава нов файл тип `stayTransportDeclaration`.

---

## 🔧 Основни функции

- Зареждане на съществуващ ЕДП XML файл
- Попълване на:
  - Област, община, населено място (с автоматично попълване на кодовете)
  - Адрес, номер и дата
- Извличане на информация за гориво, превозвач, шофьор и МПС от ЕДП
- Генериране на stayTransportDeclaration XML файл, готов за подаване към НАП

---

## 🖥️ Технологии

- Python 3.x
- `tkinter` — GUI
- `tkcalendar` — за избор на дата
- `openpyxl` — за четене на Excel файловете с кодове
- `xml.etree.ElementTree` и `minidom` — за създаване и форматиране на XML

---

## 📦 Инсталация

1. **Клонирай репото:**

```bash
git clone https://github.com/plamensve/E_STAY_GEN.git
cd E_STAY_GEN
pip install -r requirements.txt
pip install tkcalendar openpyxl

python eStayGen.py

📁 Вход и изход
Вход: XML файл тип ЕДП, предоставен от НАП (или генериран от друг модул)
Изход: XML файл тип stayTransportDeclaration, запазен на избрано място

📝 Автор
Plamen Svetoslavov
© 2025 - All rights reserved


