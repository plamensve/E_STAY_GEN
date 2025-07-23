import tkinter as tk
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os
import openpyxl
# pip install tkcalendar
from tkcalendar import DateEntry
import sys  # <-- добави този ред, ако още не съществува
import json


def resource_path(relative_path):
    """Намира пътя до файл при работа с PyInstaller и при обикновен скрипт."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


selected_file = None  # Глобална променлива за избрания входен файл
output_file_path = None  # Глобална променлива за избрания изходен файл


# --- Зареждане на данни от xlsx файловете ---
def load_dict_from_xlsx(filepath):
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    result = {}
    for row in ws.iter_rows(min_row=1, max_col=2):
        if row[1].value and isinstance(row[1].value, str):
            key = row[1].value.strip()
            code = str(row[0].value).strip() if row[0].value is not None else ''
            result[key] = code
    return result


DOMAIN_DICT = load_dict_from_xlsx(resource_path('data/domain.xlsx'))
MUNICIPALITY_DICT = load_dict_from_xlsx(resource_path('data/municipality.xlsx'))
CITY_DICT = load_dict_from_xlsx(resource_path('data/city.xlsx'))

DOMAIN_LIST = list(DOMAIN_DICT.keys())
MUNICIPALITY_LIST = list(MUNICIPALITY_DICT.keys())
CITY_LIST = list(CITY_DICT.keys())


# --- AutocompleteEntry widget ---
class AutocompleteEntry(tk.Entry):
    def __init__(self, autocomplete_list, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.autocomplete_list = sorted(autocomplete_list, key=str.lower)
        self.var = self["textvariable"] = tk.StringVar()
        self.var.trace('w', self.changed)
        self.bind("<Down>", self.move_down)
        self.bind("<Up>", self.move_up)
        self.bind("<Return>", self.selection)
        self.bind("<FocusOut>", lambda e: self.hide_listbox())
        self.listbox = None
        self.lb_index = 0
        self.root = self.winfo_toplevel()
        self.on_select_callback = None
        print(f"AutocompleteEntry loaded with {len(self.autocomplete_list)} items.")

    def set_on_select(self, callback):
        self.on_select_callback = callback

    def changed(self, *args):
        print(f"Entry changed: '{self.var.get()}'")
        if self.var.get() == '':
            self.hide_listbox()
            if self.on_select_callback:
                self.on_select_callback('')
        else:
            words = self.comparison()
            print(f"Suggestions: {words}")
            if words:
                self.show_listbox()
                self.listbox.delete(0, tk.END)
                for w in words:
                    self.listbox.insert(tk.END, w)
                self.lb_index = 0
                self.listbox.select_set(self.lb_index)
            else:
                self.hide_listbox()

    def selection(self, event):
        if self.listbox and self.listbox.size() > 0:
            value = self.listbox.get(tk.ACTIVE)
            self.var.set(value)
            self.icursor(tk.END)
            self.hide_listbox()
            if self.on_select_callback:
                self.on_select_callback(value)
        return 'break'

    def move_down(self, event):
        if self.listbox:
            if self.lb_index < self.listbox.size() - 1:
                self.lb_index += 1
                self.listbox.select_clear(0, tk.END)
                self.listbox.select_set(self.lb_index)
                self.listbox.activate(self.lb_index)
        return 'break'

    def move_up(self, event):
        if self.listbox:
            if self.lb_index > 0:
                self.lb_index -= 1
                self.listbox.select_clear(0, tk.END)
                self.listbox.select_set(self.lb_index)
                self.listbox.activate(self.lb_index)
        return 'break'

    def show_listbox(self):
        if not self.listbox:
            self.listbox = tk.Listbox(self.root, width=self["width"])
            x = self.winfo_rootx() - self.root.winfo_rootx()
            y = self.winfo_rooty() - self.root.winfo_rooty() + self.winfo_height()
            self.listbox.place(x=x, y=y)
            self.listbox.bind("<Button-1>", self.selection)
            self.listbox.bind("<Return>", self.selection)
        else:
            x = self.winfo_rootx() - self.root.winfo_rootx()
            y = self.winfo_rooty() - self.root.winfo_rooty() + self.winfo_height()
            self.listbox.place(x=x, y=y)
            self.listbox.lift()

    def hide_listbox(self):
        if self.listbox:
            self.listbox.destroy()
            self.listbox = None

    def comparison(self):
        pattern = self.var.get().lower()
        return [w for w in self.autocomplete_list if pattern in w.lower()]


def convert_xml(input_file, output_path):
    print("[DEBUG] Започва генериране на stayTransportDeclaration XML...")
    tree = ET.parse(input_file)
    root = tree.getroot()

    try:
        ukn_eADD = root.findtext('declarationReference/ukn_eADD', '')
        date = date_entry.get()
        fuelAmount = root.findtext('fuel/fuelAmount', '')
        fuelKNCode = root.findtext('fuel/fuelKNCode', '')

        storage_type = root.findtext('transport/storage/type', 'other_no_ESFP')
        transporter_eik = root.findtext('transport/transporter/bgCompany/eik', '')

        # === Превозни средства ===
        tugcistern_elements = root.findall('transport/transportation/tugcistern')
        tugcisterns = [el.text for el in tugcistern_elements if el.text]

        tug_value = root.findtext('transport/transportation/tug', '').strip()
        registration_numbers = tugcisterns if tugcisterns else ([tug_value] if tug_value else [])

        if not registration_numbers:
            raise ValueError("Не е подаден нито един регистрационен номер!")

        # Данни за шофьора и доставчика
        receiver = root.find('receiverPerson/bgPerson')
        receiver_egn = receiver.findtext('egn', '') if receiver is not None else ''
        receiver_fname = receiver.findtext('firstName', '') if receiver is not None else ''
        receiver_lname = receiver.findtext('lastName', '') if receiver is not None else ''

        # Данни от GUI
        domain = region_code_var.get()
        municipality = municipality_code_var.get()
        city = city_code_var.get()
        address = address_entry.get().strip() if address_entry.get() != "Адрес" else ""
        address_number = number_entry.get().strip() if number_entry.get() != "№" else ""

        if not all([ukn_eADD, date, fuelAmount, fuelKNCode, domain, municipality, city, address, address_number]):
            raise ValueError("Липсват задължителни данни за генериране на XML!")

        nsmap = {"xsi": "http://www.w3.org/2001/XMLSchema-instance"}
        ET.register_namespace('xsi', nsmap['xsi'])

        stay_root = ET.Element("stayTransportDeclaration", {
            "{http://www.w3.org/2001/XMLSchema-instance}noNamespaceSchemaLocation": "baseDeclarationSchema_v1.3.xsd"
        })

        decl_ref = ET.SubElement(stay_root, "declarationReference")
        ET.SubElement(decl_ref, "ukn_eADD").text = ukn_eADD
        ET.SubElement(decl_ref, "date").text = date

        fuel = ET.SubElement(stay_root, "fuel")
        ET.SubElement(fuel, "fuelAmount").text = fuelAmount
        ET.SubElement(fuel, "fuelKNCode").text = fuelKNCode

        transport = ET.SubElement(stay_root, "transport")

        location = ET.SubElement(transport, "location")
        ET.SubElement(location, "domain").text = domain
        ET.SubElement(location, "municipality").text = municipality
        ET.SubElement(location, "city").text = city

        storage = ET.SubElement(transport, "storage")
        ET.SubElement(storage, "type").text = storage_type
        ET.SubElement(storage, "address").text = address
        ET.SubElement(storage, "addressNumber").text = address_number

        transporter = ET.SubElement(transport, "transporter")
        bgCompany = ET.SubElement(transporter, "bgCompany")
        ET.SubElement(bgCompany, "eik").text = transporter_eik

        # === Превоз ===
        transportation = ET.SubElement(transport, "transportation", {
            "{http://www.w3.org/2001/XMLSchema-instance}type": "AutoTransportationType"
        })

        # Добавяне на tugcistern елементи (ако има)
        for reg in tugcisterns:
            ET.SubElement(transportation, "tugcistern").text = reg

        # Ако няма tugcistern, използваме tug
        if not tugcisterns and tug_value:
            ET.SubElement(transportation, "tug").text = tug_value

        # Шофьор
        drivers = ET.SubElement(transportation, "drivers")
        driver_bg = ET.SubElement(drivers, "bgPerson")
        ET.SubElement(driver_bg, "egn").text = receiver_egn
        ET.SubElement(driver_bg, "firstName").text = receiver_fname
        ET.SubElement(driver_bg, "lastName").text = receiver_lname

        # Лице по доставка
        deliver_person = ET.SubElement(stay_root, "deliverPerson")
        deliver_bg = ET.SubElement(deliver_person, "bgPerson")
        ET.SubElement(deliver_bg, "egn").text = receiver_egn
        ET.SubElement(deliver_bg, "firstName").text = receiver_fname
        ET.SubElement(deliver_bg, "lastName").text = receiver_lname

        # Получател
        receiver_person = ET.SubElement(stay_root, "receiverPerson")
        receiver_bg = ET.SubElement(receiver_person, "bgPerson")
        ET.SubElement(receiver_bg, "egn").text = receiver_egn
        ET.SubElement(receiver_bg, "firstName").text = receiver_fname
        ET.SubElement(receiver_bg, "lastName").text = receiver_lname

        # Запис
        rough_string = ET.tostring(stay_root, 'utf-8')
        reparsed = minidom.parseString(rough_string)
        pretty_xml = reparsed.toprettyxml(indent="  ")

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(pretty_xml)

        print(f"[DEBUG] XML файлът е записан успешно: {output_path}")
        return output_path

    except Exception as e:
        print(f"[ERROR] {e}")
        raise ValueError(f"Грешка при обработката на XML: {e}")


# === Интерфейсни функции ===
def browse_file():
    global selected_file
    file_path = filedialog.askopenfilename(filetypes=[("XML файлове", "*.xml")])
    if file_path:
        selected_file = file_path
        file_label.config(text=f"Избран файл:\n{os.path.basename(file_path)}")


def choose_save_location():
    global output_file_path
    file_path = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML файлове", "*.xml")])
    if file_path:
        output_file_path = file_path
        save_label.config(text=f"Изходен файл:\n{os.path.basename(file_path)}")


def generate_output():
    print("selected_file:", selected_file)
    print("output_file_path:", output_file_path)
    print("region_code:", region_code_var.get())
    print("municipality_code:", municipality_code_var.get())
    print("city_code:", city_code_var.get())
    print("address:", address_entry.get())
    print("number:", number_entry.get())
    if not selected_file:
        messagebox.showwarning("Липсва файл", "Моля, първо изберете XML файл.")
        return
    if not output_file_path:
        messagebox.showwarning("Липсва място за запис", "Моля, изберете къде да се запази изходният XML файл.")
        return
    if not region_code_var.get() or not municipality_code_var.get() or not city_code_var.get():
        messagebox.showwarning("Липсва код", "Моля, изберете валидна област, община и населено място.")
        return
    if address_entry.get() == "" or address_entry.get() == "Адрес":
        messagebox.showwarning("Липсва адрес", "Моля, въведете адрес.")
        return
    if number_entry.get() == "" or number_entry.get() == "№":
        messagebox.showwarning("Липсва номер", "Моля, въведете номер.")
        return
    try:
        output = convert_xml(selected_file, output_file_path)
        messagebox.showinfo("Успех", f"Файлът е създаден:\n{output}")
        clear_fields()
    except Exception as e:
        messagebox.showerror("Грешка", str(e))


# Функция за изчистване на всички полета
def clear_fields():
    # Дата (reset до днешна дата)
    try:
        date_entry.set_date('today')
    except Exception:
        pass
    # Адрес
    address_entry.delete(0, tk.END)
    address_entry.insert(0, "   Адрес")
    address_entry.config(fg="gray")
    # Номер
    number_entry.delete(0, tk.END)
    number_entry.insert(0, "   №")
    number_entry.config(fg="gray")
    # Autocomplete полета и кодове
    region_entry.delete(0, tk.END)
    region_entry.insert(0, "   Област")
    region_entry.config(fg="gray")
    region_code_var.set("")
    municipality_entry.delete(0, tk.END)
    municipality_entry.insert(0, "   Община")
    municipality_entry.config(fg="gray")
    municipality_code_var.set("")
    city_entry.delete(0, tk.END)
    city_entry.insert(0, "   Населено място")
    city_entry.config(fg="gray")
    city_code_var.set("")
    # File labels
    file_label.config(text="Няма избран файл")
    save_label.config(text="")


# --- Въвеждане на нови стойности за ширини ---
FIELD_WIDTH = 18
CODE_WIDTH = 6
ADDRESS_WIDTH = 40
NUMBER_WIDTH = 16
DATE_WIDTH = 26

# --- Създаване на прозорец (ТРЯБВА ДА Е ПРЕДИ ВСЯКА УПОТРЕБА НА root) ---
root = tk.Tk()
root.iconbitmap(resource_path("data/eStayGen_ico.ico"))
root.title("Генератор на stayTransportDeclaration XML")
root.geometry("1150x750")

# Глобални настройки за размери и шрифт
ENTRY_WIDTH = 18
ENTRY_FONT = ("Arial", 14)
ENTRY_PADX = 5
ENTRY_PADY = 8
CODE_WIDTH = 6
CODE_FONT = ("Arial", 14)
BTN_FONT = ("Arial", 16, "bold")

# --- UI ---
title = tk.Label(root, text="Генератор на XML за НАП - ЕДД за престой", font=("Arial", 16, "bold"))
title.pack(pady=10)

# (Премахнато: старо поле и label за дата)

file_label = tk.Label(root, text="Няма избран файл", font=("Arial", 12))
file_label.pack(pady=5)

button_frame = tk.Frame(root)
button_frame.pack(pady=5)

browse_btn = tk.Button(button_frame, text="Избери входен файл ЕДП", command=browse_file, font=("Arial", 12))
browse_btn.pack(side=tk.LEFT, padx=10)

save_btn = tk.Button(button_frame, text="Избери място за запис", command=choose_save_location, font=("Arial", 12))
save_btn.pack(side=tk.LEFT, padx=10)

save_label = tk.Label(root)
save_label.pack(pady=5)

location_label = tk.Label(root, text="Място за престой", font=("Arial", 14, "bold"))
location_label.pack(pady=(20, 5))

row1_frame = tk.Frame(root)
row1_frame.pack(pady=12)

FIELD_WIDTH = 18
CODE_WIDTH = 6
FIELD_FONT = ("Arial", 14)
CODE_FONT = ("Arial", 14)

# Област
region_entry = AutocompleteEntry(DOMAIN_LIST, row1_frame, font=FIELD_FONT, width=FIELD_WIDTH)
region_entry.grid(row=0, column=0, padx=12, pady=6, sticky="ew")
region_entry.insert(0, "   Област")


def clear_region_placeholder(event):
    if region_entry.get().strip() == "Област" or region_entry.get().strip() == "":
        region_entry.delete(0, tk.END)
        region_entry.config(fg="black")


region_entry.bind("<FocusIn>", clear_region_placeholder)


def restore_region_placeholder(event):
    if region_entry.get().strip() == "":
        region_entry.insert(0, "   Област")
        region_entry.config(fg="gray")


region_entry.bind("<FocusOut>", restore_region_placeholder)
region_entry.config(fg="gray")
region_code_var = tk.StringVar()
region_code_entry = tk.Entry(row1_frame, font=CODE_FONT, width=CODE_WIDTH, textvariable=region_code_var,
                             state='readonly', justify='center')
region_code_entry.grid(row=0, column=1, padx=(0, 12), pady=6, sticky="ew")


def on_region_select(value):
    region_code_var.set(DOMAIN_DICT.get(value, ''))


region_entry.set_on_select(on_region_select)

# Община
municipality_entry = AutocompleteEntry(MUNICIPALITY_LIST, row1_frame, font=FIELD_FONT, width=FIELD_WIDTH)
municipality_entry.grid(row=0, column=2, padx=12, pady=6, sticky="ew")
municipality_entry.insert(0, "   Община")


def clear_municipality_placeholder(event):
    if municipality_entry.get().strip() == "Община" or municipality_entry.get().strip() == "":
        municipality_entry.delete(0, tk.END)
        municipality_entry.config(fg="black")


municipality_entry.bind("<FocusIn>", clear_municipality_placeholder)


def restore_municipality_placeholder(event):
    if municipality_entry.get().strip() == "":
        municipality_entry.insert(0, "   Община")
        municipality_entry.config(fg="gray")


municipality_entry.bind("<FocusOut>", restore_municipality_placeholder)
municipality_entry.config(fg="gray")
municipality_code_var = tk.StringVar()
municipality_code_entry = tk.Entry(row1_frame, font=CODE_FONT, width=CODE_WIDTH, textvariable=municipality_code_var,
                                   state='readonly', justify='center')
municipality_code_entry.grid(row=0, column=3, padx=(0, 12), pady=6, sticky="ew")


def on_municipality_select(value):
    municipality_code_var.set(MUNICIPALITY_DICT.get(value, ''))


municipality_entry.set_on_select(on_municipality_select)

# Населено място
city_entry = AutocompleteEntry(CITY_LIST, row1_frame, font=FIELD_FONT, width=FIELD_WIDTH)
city_entry.grid(row=0, column=4, padx=12, pady=6, sticky="ew")
city_entry.insert(0, "   Населено място")


def clear_city_placeholder(event):
    if city_entry.get().strip() == "Населено място" or city_entry.get().strip() == "":
        city_entry.delete(0, tk.END)
        city_entry.config(fg="black")


city_entry.bind("<FocusIn>", clear_city_placeholder)


def restore_city_placeholder(event):
    if city_entry.get().strip() == "":
        city_entry.insert(0, "   Населено място")
        city_entry.config(fg="gray")


city_entry.bind("<FocusOut>", restore_city_placeholder)
city_entry.config(fg="gray")
city_code_var = tk.StringVar()
city_code_entry = tk.Entry(row1_frame, font=CODE_FONT, width=CODE_WIDTH, textvariable=city_code_var, state='readonly',
                           justify='center')
city_code_entry.grid(row=0, column=5, padx=(0, 12), pady=6, sticky="ew")


def on_city_select(value):
    city_code_var.set(CITY_DICT.get(value, ''))


city_entry.set_on_select(on_city_select)

# Контейнер за ред 2: Дата, Адрес и Номер
row2_frame = tk.Frame(root)
row2_frame.pack(pady=12)

ADDRESS_WIDTH = 40
NUMBER_WIDTH = 16
DATE_WIDTH = 26  # по-голямо поле за дата

# Поле за дата (DateEntry) на първо място, по-голямо и подравнено с 'Област'
date_entry = DateEntry(row2_frame, font=("Arial", 12), width=DATE_WIDTH, date_pattern="dd.mm.yyyy")
date_entry.grid(row=0, column=0, padx=(0, 8), pady=6, sticky="ew")

# Адрес
address_entry = tk.Entry(row2_frame, font=FIELD_FONT, width=ADDRESS_WIDTH)
address_entry.grid(row=0, column=1, padx=(0, 8), pady=6, sticky="ew")
address_entry.insert(0, "   Адрес")


def clear_address_placeholder(event):
    if address_entry.get().strip() == "Адрес" or address_entry.get().strip() == "":
        address_entry.delete(0, tk.END)
        address_entry.config(fg="black")


address_entry.bind("<FocusIn>", clear_address_placeholder)


def restore_address_placeholder(event):
    if address_entry.get().strip() == "":
        address_entry.insert(0, "   Адрес")
        address_entry.config(fg="gray")


address_entry.bind("<FocusOut>", restore_address_placeholder)
address_entry.config(fg="gray")

# Номер
number_entry = tk.Entry(row2_frame, font=FIELD_FONT, width=NUMBER_WIDTH)
number_entry.grid(row=0, column=2, padx=(0, 8), pady=6, sticky="ew")
number_entry.insert(0, "№")


def clear_number_placeholder(event):
    if number_entry.get().strip() == "№" or number_entry.get().strip() == "":
        number_entry.delete(0, tk.END)
        number_entry.config(fg="black")


number_entry.bind("<FocusIn>", clear_number_placeholder)


def restore_number_placeholder(event):
    if number_entry.get().strip() == "":
        number_entry.insert(0, "   №")
        number_entry.config(fg="gray")


number_entry.bind("<FocusOut>", restore_number_placeholder)
number_entry.config(fg="gray")

# --- Премествам бутона най-отдолу ---
# След транспортните редове:
generate_btn = tk.Button(root, text="Генерирай XML", command=generate_output, font=("Arial", 14, "bold"), bg="#4CAF50",
                         fg="white", width=14, height=1, cursor="arrow")
generate_btn.pack(pady=12)

# ТУК ПРЕМАХВАМ СТАРИЯ FOOTER:
# footer = tk.Label(root, text="2025 Plamen Svetoslavov eStayGen v1.0", font=("Arial", 10), fg="#888888")
# footer.pack(side=tk.BOTTOM, pady=8)

SAVED_ADDRESSES_PATH = os.path.join(os.getcwd(), 'saved_addresses.json')


# --- Функции за зареждане и запис на адреси ---
def load_saved_addresses():
    try:
        with open(SAVED_ADDRESSES_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return [{} for _ in range(5)]


def save_addresses(addresses):
    with open(SAVED_ADDRESSES_PATH, 'w', encoding='utf-8') as f:
        json.dump(addresses, f, ensure_ascii=False, indent=2)


# --- Транспортни адреси ---
# Нови по-малки размери за транспортните редове
COMPANY_WIDTH_SMALL = 12
FIELD_WIDTH_SMALL = 10
CODE_WIDTH_SMALL = 5
ADDRESS_WIDTH_SMALL = 14
NUMBER_WIDTH_SMALL = 6

transport_frame = tk.Frame(root)
transport_frame.pack(pady=10)

transport_entries = []
saved_addresses = load_saved_addresses()

for i in range(5):
    row = {}
    row_frame = tk.Frame(transport_frame)
    row_frame.pack(pady=2, fill='x')


    # --- Placeholder logic for all fields ---
    def make_placeholder(entry, placeholder):
        def clear_placeholder(event, entry=entry, placeholder=placeholder):
            if entry.get() == placeholder:
                entry.delete(0, tk.END)
                entry.config(fg="black")

        def restore_placeholder(event, entry=entry, placeholder=placeholder):
            if entry.get().strip() == "":
                entry.insert(0, placeholder)
                entry.config(fg="gray")

        entry.bind("<FocusIn>", clear_placeholder)
        entry.bind("<FocusOut>", restore_placeholder)
        # При първоначално създаване
        if entry.get().strip() == "":
            entry.insert(0, placeholder)
            entry.config(fg="gray")


    # Транспортна фирма (първа колона)
    row['company'] = tk.Entry(row_frame, font=("Arial", 12), width=COMPANY_WIDTH_SMALL)
    row['company'].delete(0, tk.END)
    row['company'].insert(0, saved_addresses[i].get('company', f"Тр. Фирма {i + 1}"))
    row['company'].config(fg="gray" if not saved_addresses[i].get('company') else "black")
    row['company'].grid(row=0, column=0, padx=2)
    make_placeholder(row['company'], f"Тр. Фирма {i + 1}")

    # Област
    row['region'] = AutocompleteEntry(DOMAIN_LIST, row_frame, font=("Arial", 12), width=FIELD_WIDTH_SMALL)
    row['region'].grid(row=0, column=1, padx=2)
    row['region_code_var'] = tk.StringVar()
    row['region_code'] = tk.Entry(row_frame, font=("Arial", 12), width=8, textvariable=row['region_code_var'],
                                  state='readonly', justify='center')
    row['region_code'].grid(row=0, column=2, padx=2)


    def on_region_select_local(value, var=row['region_code_var']):
        var.set(DOMAIN_DICT.get(value, ''))


    row['region'].set_on_select(on_region_select_local)
    row['region'].delete(0, tk.END)
    row['region'].insert(0, saved_addresses[i].get('region', 'Област'))
    row['region'].config(fg="gray" if not saved_addresses[i].get('region') else "black")
    row['region_code_var'].set(saved_addresses[i].get('region_code', ''))
    make_placeholder(row['region'], 'Област')

    # Община
    row['municipality'] = AutocompleteEntry(MUNICIPALITY_LIST, row_frame, font=("Arial", 12), width=FIELD_WIDTH_SMALL)
    row['municipality'].grid(row=0, column=3, padx=2)
    row['municipality_code_var'] = tk.StringVar()
    row['municipality_code'] = tk.Entry(row_frame, font=("Arial", 12), width=8,
                                        textvariable=row['municipality_code_var'], state='readonly', justify='center')
    row['municipality_code'].grid(row=0, column=4, padx=2)


    def on_municipality_select_local(value, var=row['municipality_code_var']):
        var.set(MUNICIPALITY_DICT.get(value, ''))


    row['municipality'].set_on_select(on_municipality_select_local)
    row['municipality'].delete(0, tk.END)
    row['municipality'].insert(0, saved_addresses[i].get('municipality', 'Община'))
    row['municipality'].config(fg="gray" if not saved_addresses[i].get('municipality') else "black")
    row['municipality_code_var'].set(saved_addresses[i].get('municipality_code', ''))
    make_placeholder(row['municipality'], 'Община')

    # Населено място
    row['city'] = AutocompleteEntry(CITY_LIST, row_frame, font=("Arial", 12), width=FIELD_WIDTH_SMALL)
    row['city'].grid(row=0, column=5, padx=2)
    row['city_code_var'] = tk.StringVar()
    row['city_code'] = tk.Entry(row_frame, font=("Arial", 12), width=8, textvariable=row['city_code_var'],
                                state='readonly', justify='center')
    row['city_code'].grid(row=0, column=6, padx=2)


    def on_city_select_local(value, var=row['city_code_var']):
        var.set(CITY_DICT.get(value, ''))


    row['city'].set_on_select(on_city_select_local)
    row['city'].delete(0, tk.END)
    row['city'].insert(0, saved_addresses[i].get('city', 'Населено място'))
    row['city'].config(fg="gray" if not saved_addresses[i].get('city') else "black")
    row['city_code_var'].set(saved_addresses[i].get('city_code', ''))
    make_placeholder(row['city'], 'Населено място')

    # Адрес
    row['address'] = tk.Entry(row_frame, font=("Arial", 12), width=ADDRESS_WIDTH_SMALL)
    row['address'].delete(0, tk.END)
    row['address'].insert(0, saved_addresses[i].get('address', f"Адрес {i + 1}"))
    row['address'].config(fg="gray" if not saved_addresses[i].get('address') else "black")
    row['address'].grid(row=0, column=7, padx=2)
    make_placeholder(row['address'], f"Адрес {i + 1}")

    # Номер
    row['number'] = tk.Entry(row_frame, font=("Arial", 12), width=NUMBER_WIDTH_SMALL)
    row['number'].delete(0, tk.END)
    row['number'].insert(0, saved_addresses[i].get('number', f"№{i + 1}"))
    row['number'].config(fg="gray" if not saved_addresses[i].get('number') else "black")
    row['number'].grid(row=0, column=8, padx=2)
    make_placeholder(row['number'], f"№{i + 1}")


    def apply_address(
            region_entry=row['region'], region_code_var=row['region_code_var'],
            municipality_entry=row['municipality'], municipality_code_var=row['municipality_code_var'],
            city_entry=row['city'], city_code_var=row['city_code_var'],
            addr_entry=row['address'], num_entry=row['number']):

        # Apply to main fields for XML
        region_entry_val = region_entry.get()
        municipality_entry_val = municipality_entry.get()
        city_entry_val = city_entry.get()
        address_val = addr_entry.get()
        number_val = num_entry.get()

        # Main region
        region_entry_main.delete(0, tk.END)
        region_entry_main.insert(0, region_entry_val)
        region_entry_main.config(fg="black")
        region_code_var_main.set(region_code_var.get())
        if hasattr(region_entry_main, 'on_select_callback') and region_entry_main.on_select_callback:
            region_entry_main.on_select_callback(region_entry_val)

        # Main municipality
        municipality_entry_main.delete(0, tk.END)
        municipality_entry_main.insert(0, municipality_entry_val)
        municipality_entry_main.config(fg="black")
        municipality_code_var_main.set(municipality_code_var.get())
        if hasattr(municipality_entry_main, 'on_select_callback') and municipality_entry_main.on_select_callback:
            municipality_entry_main.on_select_callback(municipality_entry_val)

        # Main city
        city_entry_main.delete(0, tk.END)
        city_entry_main.insert(0, city_entry_val)
        city_entry_main.config(fg="black")
        city_code_var_main.set(city_code_var.get())
        if hasattr(city_entry_main, 'on_select_callback') and city_entry_main.on_select_callback:
            city_entry_main.on_select_callback(city_entry_val)

        # Main address
        address_entry.delete(0, tk.END)
        address_entry.insert(0, address_val)
        address_entry.config(fg="black")

        # Main number
        number_entry.delete(0, tk.END)
        number_entry.insert(0, number_val)
        number_entry.config(fg="black")


    region_entry_main = region_entry
    region_code_var_main = region_code_var
    municipality_entry_main = municipality_entry
    municipality_code_var_main = municipality_code_var
    city_entry_main = city_entry
    city_code_var_main = city_code_var

    apply_btn = tk.Button(row_frame, text="Приложи адрес", font=("Arial", 9), width=13, command=apply_address,
                          bg="#2196F3", fg="white")
    apply_btn.grid(row=0, column=11, padx=8)


    def clear_row(idx=i):
        for key, entry in transport_entries[idx].items():
            if isinstance(entry, tk.Entry):
                if key == 'company':
                    entry.delete(0, tk.END)
                    entry.insert(0, f"Тр. Фирма {idx + 1}")
                    entry.config(fg="gray")
                elif key == 'address':
                    entry.delete(0, tk.END)
                    entry.insert(0, f"Адрес {idx + 1}")
                    entry.config(fg="gray")
                elif key == 'number':
                    entry.delete(0, tk.END)
                    entry.insert(0, f"№{idx + 1}")
                    entry.config(fg="gray")
                elif key == 'region':
                    entry.delete(0, tk.END)
                    entry.insert(0, "Област")
                    entry.config(fg="gray")
                elif key == 'municipality':
                    entry.delete(0, tk.END)
                    entry.insert(0, "Община")
                    entry.config(fg="gray")
                elif key == 'city':
                    entry.delete(0, tk.END)
                    entry.insert(0, "Населено място")
                    entry.config(fg="gray")
            elif isinstance(entry, tk.StringVar):
                entry.set("")
        # Изтриване от файла
        saved_addresses[idx] = {}
        save_addresses(saved_addresses)


    del_btn = tk.Button(row_frame, text="Изтрий адрес", font=("Arial", 9), width=13, command=clear_row, bg="#F44336",
                        fg="white")
    del_btn.grid(row=0, column=12, padx=8)


    # --- Автоматично запазване при промяна ---
    def save_row(event=None, idx=i, row=row):
        if row is None:
            row = row
        saved_addresses[idx] = {
            'region': row['region'].get(),
            'region_code': row['region_code_var'].get(),
            'municipality': row['municipality'].get(),
            'municipality_code': row['municipality_code_var'].get(),
            'city': row['city'].get(),
            'city_code': row['city_code_var'].get(),
            'company': row['company'].get(),
            'address': row['address'].get(),
            'number': row['number'].get().strip(),
        }
        save_addresses(saved_addresses)


    for key in ['region', 'municipality', 'city', 'company', 'address', 'number']:
        row[key].bind('<FocusOut>', save_row)

    transport_entries.append(row)


def on_btn_enter(event):
    generate_btn.config(cursor="hand2")


def on_btn_leave(event):
    generate_btn.config(cursor="arrow")


def on_close():
    for idx, row in enumerate(transport_entries):
        saved_addresses[idx] = {
            'region': row['region'].get(),
            'region_code': row['region_code_var'].get(),
            'municipality': row['municipality'].get(),
            'municipality_code': row['municipality_code_var'].get(),
            'city': row['city'].get(),
            'city_code': row['city_code_var'].get(),
            'company': row['company'].get(),
            'address': row['address'].get(),
            'number': row['number'].get().strip(),
        }
    save_addresses(saved_addresses)
    root.destroy()




generate_btn.bind("<Enter>", on_btn_enter)
generate_btn.bind("<Leave>", on_btn_leave)

# --- Footer -----
# Най-отдолу:
footer = tk.Label(root, text="2025 Plamen Svetoslavov eStayGen v1.1", font=("Arial", 10), fg="#888888")
footer.pack(side=tk.BOTTOM, pady=8)

root.iconbitmap(resource_path("data/eStayGen_ico.ico"))

root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()
